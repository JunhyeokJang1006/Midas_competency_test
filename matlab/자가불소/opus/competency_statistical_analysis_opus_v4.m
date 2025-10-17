%% 인재유형 종합 분석 시스템 - 최종 통합 버전
% HR Talent Type Comprehensive Analysis System - Final Integrated Version
% 목적: 1) 인재 유형별 역량점수 상위요인 프로파일링 (with radar chart)
%      2) 상관 분석을 이용한 가중치 부여
%      3) 고성과자-저성과자 로지스틱 회귀 분석
%      4) 다중레이블 및 이진 분류 머신러닝 통합

clear; clc; close all;

%% ========================================================================
%                          PART 1: 초기 설정 및 데이터 로딩
% =========================================================================

fprintf('╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║       인재유형 종합 분석 시스템 - 최종 통합 버전         ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

% 전역 설정 - 한국어 폰트 및 출판용 그래픽 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);
set(0, 'DefaultLineLineWidth', 2);
set(0, 'DefaultFigureColor', 'white');
set(0, 'DefaultAxesBox', 'off');
set(0, 'DefaultAxesTickDir', 'out');

% 컬러 팔레트 설정 (논문용)
colors = struct();
colors.primary = [0.2, 0.4, 0.8];     % 진한 파란색
colors.secondary = [0.8, 0.3, 0.3];   % 붉은색
colors.tertiary = [0.3, 0.7, 0.3];    % 녹색
colors.quaternary = [0.9, 0.6, 0.2];  % 주황색
colors.gray = [0.5, 0.5, 0.5];        % 회색

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.output_dir = pwd;
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 성과 순위 정의 (위장형 소화성 제외)
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 1]);

% 제외할 인재유형
config.excluded_types = {'위장형 소화성'};

%% 1.1 데이터 로딩
fprintf('【STEP 1】 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

try
    % HR 데이터 로딩
    fprintf('▶ HR 데이터 로딩 중...\n');
    hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
    fprintf('  ✓ HR 데이터: %d명 로드 완료\n', height(hr_data));
    
    % 역량검사 데이터 로딩
    fprintf('▶ 역량검사 데이터 로딩 중...\n');
    comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
    comp_total = readtable(config.comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 상위항목 데이터: %d명\n', height(comp_upper));
    fprintf('  ✓ 종합점수 데이터: %d명\n', height(comp_total));
    
catch ME
    error('데이터 로딩 실패: %s', ME.message);
end

%% 1.2 인재유형 데이터 추출 및 정제
fprintf('\n【STEP 2】 인재유형 데이터 추출 및 정제\n');
fprintf('────────────────────────────────────────────\n');

% 인재유형 컬럼 찾기
talent_col_idx = find(contains(hr_data.Properties.VariableNames, {'인재유형', '인재', '유형'}), 1);
if isempty(talent_col_idx)
    error('인재유형 컬럼을 찾을 수 없습니다.');
end

talent_col_name = hr_data.Properties.VariableNames{talent_col_idx};
fprintf('▶ 인재유형 컬럼: %s\n', talent_col_name);

% 빈 값 및 제외 유형 제거
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_idx}), :);

% 위장형 소화성 제외
excluded_mask = false(height(hr_clean), 1);
for i = 1:length(config.excluded_types)
    excluded_mask = excluded_mask | strcmp(hr_clean{:, talent_col_idx}, config.excluded_types{i});
end
hr_clean = hr_clean(~excluded_mask, :);

fprintf('▶ 유효한 인재유형 데이터: %d명 (위장형 소화성 제외)\n', height(hr_clean));

% 인재유형 분포 분석
talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('\n인재유형 분포:\n');
for i = 1:length(unique_types)
    fprintf('  • %-20s: %3d명 (%5.1f%%)\n', ...
        unique_types{i}, type_counts(i), type_counts(i)/sum(type_counts)*100);
end

%% 1.3 역량 데이터 처리
fprintf('\n【STEP 3】 역량 데이터 처리\n');
fprintf('────────────────────────────────────────────\n');

% ID 컬럼 찾기
comp_id_col = find(contains(lower(comp_upper.Properties.VariableNames), {'id', '사번'}), 1);
if isempty(comp_id_col)
    error('역량 데이터에서 ID 컬럼을 찾을 수 없습니다.');
end

% 유효한 역량 컬럼 추출
valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)
    col_data = comp_upper{:, i};
    if isnumeric(col_data) && ~all(isnan(col_data))
        valid_data = col_data(~isnan(col_data));
        if length(valid_data) >= 5 && var(valid_data) > 0 && ...
           all(valid_data >= 0) && all(valid_data <= 100)
            valid_comp_cols{end+1} = comp_upper.Properties.VariableNames{i};
            valid_comp_indices(end+1) = i;
        end
    end
end

fprintf('▶ 유효한 역량 항목: %d개\n', length(valid_comp_cols));

%% 1.4 ID 매칭 및 데이터 통합
fprintf('\n【STEP 4】 데이터 매칭 및 통합\n');
fprintf('────────────────────────────────────────────\n');

% ID 표준화
hr_ids_str = arrayfun(@(x) sprintf('%.0f', x), hr_clean.ID, 'UniformOutput', false);
comp_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_upper{:, comp_id_col}, 'UniformOutput', false);

% 교집합 찾기
[matched_ids, hr_idx, comp_idx] = intersect(hr_ids_str, comp_ids_str);

if length(matched_ids) < 10
    error('매칭된 데이터가 부족합니다: %d명', length(matched_ids));
end

fprintf('▶ 매칭 성공: %d명\n', length(matched_ids));

% 매칭된 데이터 추출
matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, valid_comp_indices);
matched_talent_types = matched_hr{:, talent_col_idx};

% 종합점수 매칭
comp_total_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_total{:, 1}, 'UniformOutput', false);
[~, ~, total_idx] = intersect(matched_ids, comp_total_ids_str);

if ~isempty(total_idx)
    total_scores = comp_total{total_idx, end};
    fprintf('▶ 종합점수 통합: %d명\n', length(total_idx));
else
    total_scores = [];
end

%% ========================================================================
%            PART 2: 인재유형별 역량 프로파일링 (Radar Chart)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║        PART 2: 인재유형별 역량 프로파일링                ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 2.1 유형별 프로파일 계산
fprintf('【STEP 5】 인재유형별 역량 프로파일 계산\n');
fprintf('────────────────────────────────────────────\n');

unique_matched_types = unique(matched_talent_types);
n_types = length(unique_matched_types);

% 프로파일 통계 테이블
profile_stats = table();
profile_stats.TalentType = unique_matched_types;
profile_stats.Count = zeros(n_types, 1);
profile_stats.CompetencyMean = zeros(n_types, 1);
profile_stats.CompetencyStd = zeros(n_types, 1);
profile_stats.PerformanceRank = zeros(n_types, 1);

% 각 유형별 상세 프로파일
type_profiles = zeros(n_types, length(valid_comp_cols));

for i = 1:n_types
    type_name = unique_matched_types{i};
    type_mask = strcmp(matched_talent_types, type_name);
    
    profile_stats.Count(i) = sum(type_mask);
    
    type_comp_data = matched_comp{type_mask, :};
    profile_stats.CompetencyMean(i) = nanmean(type_comp_data(:));
    profile_stats.CompetencyStd(i) = nanstd(type_comp_data(:));
    
    if config.performance_ranking.isKey(type_name)
        profile_stats.PerformanceRank(i) = config.performance_ranking(type_name);
    end
    
    type_profiles(i, :) = nanmean(type_comp_data, 1);
end

% 상위 역량 선정 (분산이 큰 상위 12개)
comp_variance = var(table2array(matched_comp), 0, 1, 'omitnan');
[~, var_idx] = sort(comp_variance, 'descend');
top_comp_idx = var_idx(1:min(12, length(var_idx)));
top_comp_names = valid_comp_cols(top_comp_idx);

%% 2.2 논문 수준 Radar Chart 생성
fprintf('\n【STEP 6】 레이더 차트 생성\n');
fprintf('────────────────────────────────────────────\n');

% Figure 1: 인재유형별 역량 프로파일 레이더 차트
fig1 = figure('Position', [50, 50, 1800, 1200], 'Color', 'white');

% 전체 평균 프로파일
overall_mean_profile = nanmean(table2array(matched_comp), 1);

% 성과 순위별 정렬
[~, sort_idx] = sort(profile_stats.PerformanceRank, 'descend');
sorted_types = unique_matched_types(sort_idx);

% 서브플롯 레이아웃
n_rows = ceil(sqrt(n_types));
n_cols = ceil(n_types / n_rows);

for i = 1:n_types
    subplot(n_rows, n_cols, i);
    
    type_idx = find(strcmp(unique_matched_types, sorted_types{i}));
    type_profile = type_profiles(type_idx, top_comp_idx);
    baseline = overall_mean_profile(top_comp_idx);
    
    % 레이더 차트 그리기 (v2 방식)
    createRadarChart(type_profile, baseline, top_comp_names, ...
                    sorted_types{i}, colors.primary);
    
    % 성과 순위 표시
    perf_rank = profile_stats.PerformanceRank(type_idx);
    text(0, -1.5, sprintf('성과순위: %d', perf_rank), ...
         'HorizontalAlignment', 'center', 'FontWeight', 'bold', 'FontSize', 10);
end

sgtitle('인재유형별 역량 프로파일 분석', 'FontSize', 18, 'FontWeight', 'bold');

%% ========================================================================
%                 PART 3: 상관분석 기반 가중치 계산
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║         PART 3: 상관분석 기반 가중치 계산                ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 3.1 성과점수 기반 상관분석
fprintf('【STEP 7】 역량-성과 상관분석\n');
fprintf('────────────────────────────────────────────\n');

% 성과점수 할당
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    type_name = matched_talent_types{i};
    if config.performance_ranking.isKey(type_name)
        performance_scores(i) = config.performance_ranking(type_name);
    end
end

valid_perf_idx = performance_scores > 0;
valid_performance = performance_scores(valid_perf_idx);
valid_competencies = table2array(matched_comp(valid_perf_idx, :));

% 상관계수 계산
n_competencies = length(valid_comp_cols);
correlation_results = table();
correlation_results.Competency = valid_comp_cols';
correlation_results.Correlation = zeros(n_competencies, 1);
correlation_results.PValue = zeros(n_competencies, 1);
correlation_results.Significance = cell(n_competencies, 1);

for i = 1:n_competencies
    comp_scores = valid_competencies(:, i);
    valid_idx = ~isnan(comp_scores);
    
    if sum(valid_idx) >= 10
        [r, p] = corr(comp_scores(valid_idx), valid_performance(valid_idx), 'Type', 'Spearman');
        correlation_results.Correlation(i) = r;
        correlation_results.PValue(i) = p;
        
        % 유의성 표시
        if p < 0.001
            correlation_results.Significance{i} = '***';
        elseif p < 0.01
            correlation_results.Significance{i} = '**';
        elseif p < 0.05
            correlation_results.Significance{i} = '*';
        else
            correlation_results.Significance{i} = '';
        end
    end
end

% 가중치 계산
positive_corr = max(0, correlation_results.Correlation);
weights_corr = positive_corr / sum(positive_corr);
correlation_results.Weight = weights_corr * 100;

correlation_results = sortrows(correlation_results, 'Correlation', 'descend');

%% 3.2 상관분석 시각화
% Figure 2: 상관분석 결과
fig2 = figure('Position', [100, 100, 1600, 900], 'Color', 'white');

% 상위 15개 역량의 상관계수와 가중치
subplot(2, 2, [1, 2]);
top_15 = correlation_results(1:min(15, height(correlation_results)), :);
x = 1:height(top_15);

yyaxis left
bar(x, top_15.Correlation, 'FaceColor', colors.primary, 'EdgeColor', 'none');
ylabel('상관계수', 'FontSize', 12, 'FontWeight', 'bold');
ylim([-0.2, max(top_15.Correlation)*1.2]);

yyaxis right
plot(x, top_15.Weight, '-o', 'Color', colors.secondary, 'LineWidth', 2, ...
     'MarkerFaceColor', colors.secondary, 'MarkerSize', 8);
ylabel('가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');

set(gca, 'XTick', x, 'XTickLabel', top_15.Competency, 'XTickLabelRotation', 45);
xlabel('역량 항목', 'FontSize', 12, 'FontWeight', 'bold');
title('역량-성과 상관분석 및 가중치', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
box off;

% 누적 가중치
subplot(2, 2, 3);
cumulative_weight = cumsum(correlation_results.Weight);
plot(cumulative_weight, 'LineWidth', 2.5, 'Color', colors.tertiary);
hold on;
plot([1, length(cumulative_weight)], [80, 80], '--', 'Color', colors.gray, 'LineWidth', 2);
xlabel('역량 개수', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('누적 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('누적 설명력 분석', 'FontSize', 14, 'FontWeight', 'bold');
legend('누적 가중치', '80% 기준선', 'Location', 'southeast', 'FontSize', 10);
grid on;
box off;

% 유의성 분포
subplot(2, 2, 4);
sig_counts = [sum(strcmp(correlation_results.Significance, '***')), ...
              sum(strcmp(correlation_results.Significance, '**')), ...
              sum(strcmp(correlation_results.Significance, '*')), ...
              sum(strcmp(correlation_results.Significance, ''))];
pie(sig_counts, {'p<0.001', 'p<0.01', 'p<0.05', 'n.s.'});
title('상관계수 유의성 분포', 'FontSize', 14, 'FontWeight', 'bold');
colormap(gca, [colors.primary; colors.secondary; colors.tertiary; colors.gray]);

%% ========================================================================
%            PART 4: 고성과자-저성과자 로지스틱 회귀
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║      PART 4: 고성과자-저성과자 로지스틱 회귀 분석       ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 4.1 이진 분류 레이블 생성
fprintf('【STEP 8】 고성과자-저성과자 분류\n');
fprintf('────────────────────────────────────────────\n');

% 중앙값 기준 이진 분류
performance_threshold = median(valid_performance);
high_perf_labels = double(valid_performance > performance_threshold);

fprintf('▶ 고성과자 (상위 50%%): %d명\n', sum(high_perf_labels == 1));
fprintf('▶ 저성과자 (하위 50%%): %d명\n', sum(high_perf_labels == 0));

% 데이터 정규화
X_norm = normalize(valid_competencies, 'range');
y_binary = high_perf_labels;

% 훈련/테스트 분할
rng(42); % 재현성
test_ratio = 0.2;
n_samples = length(y_binary);

high_indices = find(y_binary == 1);
low_indices = find(y_binary == 0);

n_high_test = round(length(high_indices) * test_ratio);
n_low_test = round(length(low_indices) * test_ratio);

high_test_idx = high_indices(randperm(length(high_indices), n_high_test));
low_test_idx = low_indices(randperm(length(low_indices), n_low_test));

test_indices = [high_test_idx; low_test_idx];
train_indices = setdiff(1:n_samples, test_indices);

X_train = X_norm(train_indices, :);
y_train = y_binary(train_indices);
X_test = X_norm(test_indices, :);
y_test = y_binary(test_indices);

%% 4.2 로지스틱 회귀 분석
fprintf('\n【STEP 9】 로지스틱 회귀 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% 로지스틱 회귀
[B, FitInfo] = lassoglm(X_train, y_train, 'binomial', ...
                        'Alpha', 0.5, 'NumLambda', 100, 'CV', 5);

% 최적 람다 선택
[~, idx_min] = min(FitInfo.Deviance);
B_optimal = [FitInfo.Intercept(idx_min); B(:, idx_min)];

% 예측
y_pred_prob = glmval(B_optimal, X_test, 'logit');
y_pred = double(y_pred_prob > 0.5);

% 성능 평가
logit_accuracy = sum(y_pred == y_test) / length(y_test);
logit_precision = sum(y_pred == 1 & y_test == 1) / sum(y_pred == 1);
logit_recall = sum(y_pred == 1 & y_test == 1) / sum(y_test == 1);
logit_f1 = 2 * (logit_precision * logit_recall) / (logit_precision + logit_recall);

fprintf('로지스틱 회귀 성능:\n');
fprintf('  • 정확도: %.3f\n', logit_accuracy);
fprintf('  • 정밀도: %.3f\n', logit_precision);
fprintf('  • 재현율: %.3f\n', logit_recall);
fprintf('  • F1-Score: %.3f\n', logit_f1);

% 중요 변수 식별
important_vars = find(abs(B_optimal(2:end)) > 0.01);
logit_importance = abs(B_optimal(2:end)) / sum(abs(B_optimal(2:end)));

% logit_importance를 valid_comp_cols 크기에 맞춤
if length(logit_importance) ~= length(valid_comp_cols)
    logit_importance_temp = zeros(length(valid_comp_cols), 1);
    logit_importance_temp(1:min(length(logit_importance), length(valid_comp_cols))) = ...
        logit_importance(1:min(length(logit_importance), length(valid_comp_cols)));
    logit_importance = logit_importance_temp;
end

% logit_importance_norm 정의
logit_importance_norm = logit_importance / sum(logit_importance);

%% 4.3 ROC 곡선 및 성능 시각화
% Figure 3: 로지스틱 회귀 결과
fig3 = figure('Position', [150, 150, 1400, 800], 'Color', 'white');

% ROC 곡선
subplot(2, 3, 1);
[fpr, tpr, ~, auc_value] = perfcurve(y_test, y_pred_prob, 1);
plot(fpr, tpr, 'LineWidth', 2.5, 'Color', colors.primary);
hold on;
plot([0, 1], [0, 1], '--', 'Color', colors.gray, 'LineWidth', 1.5);
xlabel('위양성률', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('진양성률', 'FontSize', 12, 'FontWeight', 'bold');
title(sprintf('ROC 곡선 (AUC = %.3f)', auc_value), 'FontSize', 14, 'FontWeight', 'bold');
legend('로지스틱 회귀', '기준선', 'Location', 'southeast', 'FontSize', 10);
grid on;
axis square;

% 혼동 행렬
subplot(2, 3, 2);
conf_matrix = confusionmat(y_test, y_pred);
imagesc(conf_matrix);
colormap(gca, flipud(gray));
colorbar;
xlabel('예측값', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('실제값', 'FontSize', 12, 'FontWeight', 'bold');
title('혼동 행렬', 'FontSize', 14, 'FontWeight', 'bold');
set(gca, 'XTick', [1, 2], 'XTickLabel', {'저성과', '고성과'}, ...
         'YTick', [1, 2], 'YTickLabel', {'저성과', '고성과'});

% 값 표시
for i = 1:2
    for j = 1:2
        if conf_matrix(i,j) > max(conf_matrix(:))/2
            text_color = 'w';
        else
            text_color = 'k';
        end
        text(j, i, num2str(conf_matrix(i,j)), ...
             'HorizontalAlignment', 'center', ...
             'Color', text_color, ...
             'FontSize', 14, 'FontWeight', 'bold');
    end
end

% 계수 중요도
subplot(2, 3, [3, 6]);
[sorted_imp, imp_idx] = sort(logit_importance, 'descend');
top_20_idx = imp_idx(1:min(20, length(imp_idx)));
barh(20:-1:max(1, 21-length(top_20_idx)), sorted_imp(1:min(20, length(sorted_imp))), ...
     'FaceColor', colors.secondary, 'EdgeColor', 'none');
set(gca, 'YTick', 1:min(20, length(top_20_idx)), ...
         'YTickLabel', flip(valid_comp_cols(top_20_idx)));
xlabel('중요도', 'FontSize', 12, 'FontWeight', 'bold');
title('로지스틱 회귀 변수 중요도 (상위 20개)', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
box off;

% 예측 확률 분포
subplot(2, 3, 4);
histogram(y_pred_prob(y_test == 0), 20, 'FaceColor', colors.primary, ...
          'FaceAlpha', 0.6, 'EdgeColor', 'none');
hold on;
histogram(y_pred_prob(y_test == 1), 20, 'FaceColor', colors.secondary, ...
          'FaceAlpha', 0.6, 'EdgeColor', 'none');
xlabel('예측 확률', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('빈도', 'FontSize', 12, 'FontWeight', 'bold');
title('예측 확률 분포', 'FontSize', 14, 'FontWeight', 'bold');
legend('실제 저성과자', '실제 고성과자', 'Location', 'north', 'FontSize', 10);
grid on;
box off;

% 성능 메트릭
subplot(2, 3, 5);
metrics = [logit_accuracy; logit_precision; logit_recall; logit_f1];
bar(metrics, 'FaceColor', colors.tertiary, 'EdgeColor', 'none');
set(gca, 'XTickLabel', {'정확도', '정밀도', '재현율', 'F1-Score'});
ylabel('점수', 'FontSize', 12, 'FontWeight', 'bold');
title('성능 메트릭', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0, 1]);
grid on;

for i = 1:length(metrics)
    text(i, metrics(i) + 0.02, sprintf('%.3f', metrics(i)), ...
         'HorizontalAlignment', 'center', 'FontWeight', 'bold');
end

%% ========================================================================
%         PART 5: 다중레이블 및 고급 머신러닝 (클래스 불균형 대응)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║       PART 5: 다중레이블 머신러닝 (클래스 불균형 대응)   ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 5.1 클래스 불균형 처리 (SMOTE)
fprintf('【STEP 10】 SMOTE를 통한 클래스 균형화\n');
fprintf('────────────────────────────────────────────\n');

X = table2array(matched_comp);
y = matched_talent_types;
[y_unique, ~, y_encoded] = unique(y);

% 클래스별 샘플 수
class_counts = histcounts(y_encoded, 1:(length(y_unique)+1));
imbalance_ratio = max(class_counts) / min(class_counts);

fprintf('클래스 불균형 비율: %.1f:1\n', imbalance_ratio);

% SMOTE 적용
X_normalized = normalize(X, 'range');
target_samples = round(median(class_counts));

X_balanced = [];
y_balanced = [];

for class = 1:length(y_unique)
    class_idx = find(y_encoded == class);
    X_class = X_normalized(class_idx, :);
    n_samples = length(class_idx);
    
    if n_samples < target_samples
        % 오버샘플링
        n_synthetic = target_samples - n_samples;
        
        if n_samples >= 2
            k_neighbors = min(5, n_samples-1);
            X_synthetic = zeros(n_synthetic, size(X_class, 2));
            
            for i = 1:n_synthetic
                base_idx = randi(n_samples);
                base_sample = X_class(base_idx, :);
                
                distances = sqrt(sum((X_class - base_sample).^2, 2));
                [~, sorted_idx] = sort(distances);
                neighbor_idx = sorted_idx(randi([2, min(k_neighbors+1, n_samples)]));
                neighbor_sample = X_class(neighbor_idx, :);
                
                lambda = 0.3 + 0.4 * rand();
                noise_level = 0.01 * std(X_class(:));
                X_synthetic(i, :) = base_sample + lambda * (neighbor_sample - base_sample) + ...
                                   noise_level * randn(1, size(X_class, 2));
            end
            
            X_balanced = [X_balanced; X_class; X_synthetic];
            y_balanced = [y_balanced; repmat(class, n_samples + n_synthetic, 1)];
        else
            X_balanced = [X_balanced; repmat(X_class, target_samples, 1)];
            y_balanced = [y_balanced; repmat(class, target_samples, 1)];
        end
    elseif n_samples > target_samples * 1.5
        % 언더샘플링
        sample_idx = randsample(n_samples, target_samples);
        X_balanced = [X_balanced; X_class(sample_idx, :)];
        y_balanced = [y_balanced; repmat(class, target_samples, 1)];
    else
        X_balanced = [X_balanced; X_class];
        y_balanced = [y_balanced; repmat(class, n_samples, 1)];
    end
end

fprintf('균형화 완료: %d → %d 샘플\n', length(y_encoded), length(y_balanced));

%% 5.2 훈련/테스트 분할
test_ratio = 0.2;
test_indices = [];
for class = 1:length(y_unique)
    class_indices = find(y_encoded == class);
    n_test = max(1, round(length(class_indices) * test_ratio));
    rand_perm = randperm(length(class_indices));
    test_indices = [test_indices; class_indices(rand_perm(1:n_test))];
end
train_indices = setdiff(1:length(y_encoded), test_indices);

X_test_multi = X_normalized(test_indices, :);
y_test_multi = y_encoded(test_indices);

X_train_multi = X_balanced;
y_train_multi = y_balanced;

%% 5.3 앙상블 모델 구축
fprintf('\n【STEP 11】 앙상블 머신러닝 모델 구축\n');
fprintf('────────────────────────────────────────────\n');

% 1. Random Forest
fprintf('1. Random Forest 학습 중...\n');
rf_model = TreeBagger(200, X_train_multi, y_train_multi, ...
    'Method', 'classification', ...
    'MinLeafSize', 2, ...
    'MaxNumSplits', 30, ...
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', valid_comp_cols);

[y_pred_rf_cell, rf_scores] = predict(rf_model, X_test_multi);
y_pred_rf = cellfun(@str2double, y_pred_rf_cell);

if iscell(rf_scores)
    rf_probs = cell2mat(rf_scores);
else
    rf_probs = rf_scores;
end

rf_accuracy = sum(y_pred_rf == y_test_multi) / length(y_test_multi);
rf_importance = rf_model.OOBPermutedPredictorDeltaError;

% 2. Gradient Boosting
fprintf('2. Gradient Boosting 학습 중...\n');

% 클래스 가중치 계산
unique_train_classes = unique(y_train_multi);
class_costs = zeros(length(unique_train_classes), length(unique_train_classes));
for i = 1:length(unique_train_classes)
    for j = 1:length(unique_train_classes)
        if i ~= j
            class_i_count = sum(y_train_multi == unique_train_classes(i));
            total_samples = length(y_train_multi);
            cost_weight = total_samples / (length(unique_train_classes) * class_i_count);
            class_costs(i, j) = cost_weight;
        end
    end
end

gb_model = fitcensemble(X_train_multi, y_train_multi, ...
    'Method', 'AdaBoostM2', ...
    'NumLearningCycles', 100, ...
    'LearnRate', 0.1, ...
    'Learners', templateTree('MaxNumSplits', 20, 'MinLeafSize', 5), ...
    'Cost', class_costs);

[y_pred_gb, gb_scores] = predict(gb_model, X_test_multi);
gb_accuracy = sum(y_pred_gb == y_test_multi) / length(y_test_multi);
gb_importance = predictorImportance(gb_model);

% 3. 클래스별 전문가 모델
fprintf('3. 클래스별 전문가 모델 학습 중...\n');
class_experts = cell(length(y_unique), 1);
expert_accuracies = zeros(length(y_unique), 1);

for class = 1:length(y_unique)
    if sum(y_train_multi == class) >= 5
        y_binary = double(y_train_multi == class);
        
        class_experts{class} = fitcensemble(X_train_multi, y_binary, ...
            'Method', 'RUSBoost', ...
            'NumLearningCycles', 50, ...
            'Learners', templateTree('MaxNumSplits', 10));
        
        y_pred_binary = predict(class_experts{class}, X_test_multi);
        y_test_binary = double(y_test_multi == class);
        expert_accuracies(class) = sum(y_pred_binary == y_test_binary) / length(y_test_binary);
    end
end

% 앙상블 예측
expert_probs = zeros(length(y_test_multi), length(y_unique));
for class = 1:length(y_unique)
    if ~isempty(class_experts{class})
        [~, scores] = predict(class_experts{class}, X_test_multi);
        if size(scores, 2) >= 2
            expert_probs(:, class) = scores(:, 2);
        end
    end
end

% 가중 앙상블
model_weights = [0.4, 0.3, 0.3];
ensemble_probs = model_weights(1) * rf_probs + model_weights(2) * gb_scores + ...
                model_weights(3) * expert_probs;
ensemble_probs = ensemble_probs ./ sum(ensemble_probs, 2);
[~, y_pred_ensemble] = max(ensemble_probs, [], 2);

ensemble_accuracy = sum(y_pred_ensemble == y_test_multi) / length(y_test_multi);

fprintf('\n모델 성능 비교:\n');
fprintf('  • Random Forest: %.3f\n', rf_accuracy);
fprintf('  • Gradient Boosting: %.3f\n', gb_accuracy);
fprintf('  • 가중 앙상블: %.3f\n', ensemble_accuracy);

%% 5.4 클래스별 성능 평가
conf_ensemble = confusionmat(y_test_multi, y_pred_ensemble);

class_metrics = table();
class_metrics.TalentType = y_unique;
class_metrics.Precision = zeros(length(y_unique), 1);
class_metrics.Recall = zeros(length(y_unique), 1);
class_metrics.F1Score = zeros(length(y_unique), 1);
class_metrics.Support = zeros(length(y_unique), 1);

for i = 1:length(y_unique)
    tp = conf_ensemble(i, i);
    fp = sum(conf_ensemble(:, i)) - tp;
    fn = sum(conf_ensemble(i, :)) - tp;
    
    class_metrics.Precision(i) = tp / (tp + fp + eps);
    class_metrics.Recall(i) = tp / (tp + fn + eps);
    class_metrics.F1Score(i) = 2 * (class_metrics.Precision(i) * class_metrics.Recall(i)) / ...
                              (class_metrics.Precision(i) + class_metrics.Recall(i) + eps);
    class_metrics.Support(i) = sum(y_test_multi == i);
end

balanced_accuracy = mean(class_metrics.Recall);
macro_f1 = mean(class_metrics.F1Score);

%% 5.5 통합 Feature Importance
rf_importance_norm = rf_importance / sum(rf_importance);
gb_importance_norm = gb_importance / sum(gb_importance);

% 차원 맞추기
rf_importance_norm = rf_importance_norm(:);
gb_importance_norm = gb_importance_norm(:);
weights_corr = weights_corr(:);
logit_importance_norm = logit_importance_norm(:);

% 모든 벡터를 같은 크기로 맞춤
n_comp = length(valid_comp_cols);
if length(rf_importance_norm) ~= n_comp
    rf_importance_norm = rf_importance_norm(1:min(n_comp, length(rf_importance_norm)));
    if length(rf_importance_norm) < n_comp
        rf_importance_norm = [rf_importance_norm; zeros(n_comp - length(rf_importance_norm), 1)];
    end
end

if length(gb_importance_norm) ~= n_comp
    gb_importance_norm = gb_importance_norm(1:min(n_comp, length(gb_importance_norm)));
    if length(gb_importance_norm) < n_comp
        gb_importance_norm = [gb_importance_norm; zeros(n_comp - length(gb_importance_norm), 1)];
    end
end

if length(weights_corr) ~= n_comp
    weights_corr = weights_corr(1:min(n_comp, length(weights_corr)));
    if length(weights_corr) < n_comp
        weights_corr = [weights_corr; zeros(n_comp - length(weights_corr), 1)];
    end
end

if length(logit_importance_norm) ~= n_comp
    logit_importance_norm = logit_importance_norm(1:min(n_comp, length(logit_importance_norm)));
    if length(logit_importance_norm) < n_comp
        logit_importance_norm = [logit_importance_norm; zeros(n_comp - length(logit_importance_norm), 1)];
    end
end

final_importance = (0.3 * rf_importance_norm + 0.3 * gb_importance_norm + ...
                   0.2 * weights_corr + 0.2 * logit_importance_norm);
final_importance = final_importance / sum(final_importance);

%% 5.6 다중레이블 분류 시각화
% Figure 4: 다중레이블 분류 결과
fig4 = figure('Position', [200, 200, 1600, 1000], 'Color', 'white');

% 혼동 행렬
subplot(2, 3, [1, 2, 4, 5]);
imagesc(conf_ensemble);
colormap(gca, flipud(gray));
colorbar;
xlabel('예측 클래스', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('실제 클래스', 'FontSize', 12, 'FontWeight', 'bold');
title('다중레이블 분류 혼동 행렬', 'FontSize', 14, 'FontWeight', 'bold');
set(gca, 'XTick', 1:length(y_unique), 'XTickLabel', y_unique, ...
         'YTick', 1:length(y_unique), 'YTickLabel', y_unique, ...
         'XTickLabelRotation', 45);

% 각 셀에 숫자 표시
for i = 1:size(conf_ensemble, 1)
    for j = 1:size(conf_ensemble, 2)
        if conf_ensemble(i,j) > max(conf_ensemble(:))/2
            text_color = 'w';
        else
            text_color = 'k';
        end
        text(j, i, num2str(conf_ensemble(i,j)), ...
             'HorizontalAlignment', 'center', ...
             'Color', text_color, ...
             'FontSize', 10, 'FontWeight', 'bold');
    end
end

% 모델별 정확도 비교
subplot(2, 3, 3);
model_names = {'Random\nForest', 'Gradient\nBoosting', '가중\n앙상블'};
accuracies = [rf_accuracy, gb_accuracy, ensemble_accuracy];
bar(accuracies, 'FaceColor', colors.quaternary, 'EdgeColor', 'none');
set(gca, 'XTickLabel', model_names);
ylabel('정확도', 'FontSize', 12, 'FontWeight', 'bold');
title('모델 성능 비교', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0, 1]);
grid on;

for i = 1:length(accuracies)
    text(i, accuracies(i) + 0.02, sprintf('%.3f', accuracies(i)), ...
         'HorizontalAlignment', 'center', 'FontWeight', 'bold');
end

% 클래스별 F1-Score
subplot(2, 3, 6);
bar(class_metrics.F1Score, 'FaceColor', colors.primary, 'EdgeColor', 'none');
set(gca, 'XTickLabel', class_metrics.TalentType, 'XTickLabelRotation', 45);
ylabel('F1-Score', 'FontSize', 12, 'FontWeight', 'bold');
title('클래스별 F1-Score', 'FontSize', 14, 'FontWeight', 'bold');
grid on;

sgtitle('다중레이블 분류 성능 분석', 'FontSize', 18, 'FontWeight', 'bold');

%% ========================================================================
%                    PART 6: 최종 통합 분석 및 보고서
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║          PART 6: 최종 통합 분석 및 보고서 생성           ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 6.1 최종 Feature Importance 시각화
% Figure 5: 통합 Feature Importance
fig5 = figure('Position', [250, 250, 1600, 900], 'Color', 'white');

% Feature Importance 테이블 생성
importance_table = table();
importance_table.Competency = valid_comp_cols';
importance_table.RF_Importance = rf_importance_norm * 100;
importance_table.GB_Importance = gb_importance_norm * 100;
importance_table.Correlation = weights_corr * 100;
importance_table.Logistic = logit_importance_norm * 100;
importance_table.Final_Importance = final_importance * 100;

importance_table = sortrows(importance_table, 'Final_Importance', 'descend');

% 상위 20개 역량의 종합 중요도
subplot(2, 1, 1);
top_20 = importance_table(1:min(20, height(importance_table)), :);
bar_data = [top_20.RF_Importance, top_20.GB_Importance, ...
           top_20.Correlation, top_20.Logistic];

bar(bar_data, 'grouped');
set(gca, 'XTickLabel', top_20.Competency, 'XTickLabelRotation', 45);
ylabel('중요도 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('통합 Feature Importance 분석 (상위 20개)', 'FontSize', 14, 'FontWeight', 'bold');
legend('Random Forest', 'Gradient Boosting', '상관분석', '로지스틱 회귀', ...
       'Location', 'northeast', 'FontSize', 10);
grid on;
colormap(gca, [colors.primary; colors.secondary; colors.tertiary; colors.quaternary]);

% 최종 중요도
subplot(2, 1, 2);
barh(20:-1:1, top_20.Final_Importance, 'FaceColor', colors.primary, 'EdgeColor', 'none');
set(gca, 'YTick', 1:20, 'YTickLabel', flip(top_20.Competency));
xlabel('최종 통합 중요도 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('최종 역량 중요도 순위', 'FontSize', 14, 'FontWeight', 'bold');
grid on;

sgtitle('Feature Importance 종합 분석', 'FontSize', 18, 'FontWeight', 'bold');

%% 6.2 Excel 보고서 생성
fprintf('\n【STEP 12】 종합 Excel 보고서 생성\n');
fprintf('────────────────────────────────────────────\n');

output_filename = sprintf('talent_analysis_report_%s.xlsx', config.timestamp);

try
    % Sheet 1: 요약 정보
    summary_info = table();
    summary_info.항목 = {'분석 일시'; '전체 샘플 수'; '인재유형 수'; '역량항목 수'; ...
                       '클래스 불균형 비율'; '균형화 후 샘플 수'};
    summary_info.값 = {datestr(now); num2str(length(matched_ids)); ...
                     num2str(n_types); num2str(length(valid_comp_cols)); ...
                     sprintf('%.1f:1', imbalance_ratio); num2str(length(y_balanced))};
    writetable(summary_info, output_filename, 'Sheet', '요약정보');
    
    % Sheet 2: 인재유형 프로파일
    writetable(profile_stats, output_filename, 'Sheet', '인재유형_프로파일');
    
    % Sheet 3: 상관분석 결과
    writetable(correlation_results(1:min(30, height(correlation_results)), :), ...
              output_filename, 'Sheet', '상관분석_결과');
    
    % Sheet 4: 로지스틱 회귀 성능
    logit_performance = table();
    logit_performance.메트릭 = {'정확도'; '정밀도'; '재현율'; 'F1-Score'; 'AUC'};
    logit_performance.값 = [logit_accuracy; logit_precision; logit_recall; ...
                          logit_f1; auc_value];
    writetable(logit_performance, output_filename, 'Sheet', '로지스틱회귀_성능');
    
    % Sheet 5: 다중레이블 분류 성능
    ml_performance = table();
    ml_performance.모델 = {'Random Forest'; 'Gradient Boosting'; '가중 앙상블'};
    ml_performance.정확도 = [rf_accuracy; gb_accuracy; ensemble_accuracy];
    writetable(ml_performance, output_filename, 'Sheet', '다중레이블_모델성능');
    
    % Sheet 6: 클래스별 성능
    writetable(class_metrics, output_filename, 'Sheet', '클래스별_성능');
    
    % Sheet 7: Feature Importance
    writetable(importance_table, output_filename, 'Sheet', 'Feature_Importance');
    
    % Sheet 8: 종합 메트릭
    overall_metrics = table();
    overall_metrics.메트릭 = {'균형 정확도'; 'Macro F1-Score'; '로지스틱 AUC'; ...
                            '앙상블 정확도'; '상위3개 역량 설명력'};
    top3_importance = sum(importance_table.Final_Importance(1:3));
    overall_metrics.값 = [balanced_accuracy; macro_f1; auc_value; ...
                        ensemble_accuracy; top3_importance/100];
    writetable(overall_metrics, output_filename, 'Sheet', '종합_메트릭');
    
    fprintf('✓ Excel 보고서 저장 완료: %s\n', output_filename);
    
catch ME
    fprintf('⚠ Excel 저장 실패: %s\n', ME.message);
end

%% 6.3 Figure 저장
fprintf('\n【STEP 13】 그래프 저장\n');
fprintf('────────────────────────────────────────────\n');

fig_names = {'인재유형_프로파일', '상관분석_결과', '로지스틱회귀_분석', ...
            '다중레이블_분류', 'Feature_Importance'};
figures = [fig1, fig2, fig3, fig4, fig5];

for i = 1:length(figures)
    filename = sprintf('%s_%s.png', fig_names{i}, config.timestamp);
    saveas(figures(i), filename, 'png');
    fprintf('  ✓ %s 저장 완료\n', filename);
end

%% 6.4 최종 보고서 출력
fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                    최종 분석 보고서                       ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

fprintf('📊 데이터 요약\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • 전체 분석 대상: %d명\n', length(matched_ids));
fprintf('  • 인재유형: %d개\n', n_types);
fprintf('  • 역량항목: %d개\n', length(valid_comp_cols));
fprintf('  • 클래스 불균형 비율: %.1f:1\n', imbalance_ratio);

fprintf('\n🎯 주요 발견사항\n');
fprintf('────────────────────────────────────────────\n');
[~, best_type_idx] = max(profile_stats.PerformanceRank);
fprintf('  1. 최고 성과 인재유형: %s\n', profile_stats.TalentType{best_type_idx});
fprintf('  2. 핵심 예측 역량 Top 3:\n');
for i = 1:3
    fprintf('     - %s (%.2f%%)\n', importance_table.Competency{i}, ...
            importance_table.Final_Importance(i));
end

fprintf('\n📈 모델 성능\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • 로지스틱 회귀 (이진분류)\n');
fprintf('    - 정확도: %.1f%%, F1-Score: %.3f, AUC: %.3f\n', ...
        logit_accuracy*100, logit_f1, auc_value);
fprintf('  • 앙상블 모델 (다중분류)\n');
fprintf('    - 정확도: %.1f%%, 균형정확도: %.1f%%, Macro F1: %.3f\n', ...
        ensemble_accuracy*100, balanced_accuracy*100, macro_f1);

fprintf('\n💡 실무 적용 권장사항\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  1. 채용 스크리닝: 상위 3개 역량 (%s, %s, %s) 중점 평가\n', ...
        importance_table.Competency{1}, importance_table.Competency{2}, ...
        importance_table.Competency{3});
fprintf('  2. 인재 육성: %s 유형 벤치마킹 프로그램 개발\n', ...
        profile_stats.TalentType{best_type_idx});
fprintf('  3. 모델 활용: ');
if ensemble_accuracy > 0.7
    fprintf('우수한 성능으로 실무 적용 가능\n');
elseif ensemble_accuracy > 0.5
    fprintf('보조 도구로 활용 권장\n');
else
    fprintf('추가 데이터 수집 후 재학습 필요\n');
end

fprintf('\n🔄 향후 개선 방향\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  1. 소수 클래스 데이터 추가 수집 (최소 10명/클래스)\n');
fprintf('  2. 시계열 데이터 활용한 성과 예측 모델 개발\n');
fprintf('  3. 분기별 모델 재학습 및 업데이트\n');

fprintf('\n════════════════════════════════════════════════════════════\n');
fprintf('         분석 완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('════════════════════════════════════════════════════════════\n\n');

%% Helper Functions

function createRadarChart(data, baseline, labels, title_text, color)
    % 레이더 차트 생성 함수 (v2 방식)
    n_vars = length(data);
    angles = linspace(0, 2*pi, n_vars+1);
    
    % 데이터 정규화 (0-1 스케일)
    data_norm = data / 100;
    baseline_norm = baseline / 100;
    
    % 순환을 위해 첫 번째 값을 마지막에 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];
    
    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);
    
    hold on;
    
    % 그리드 그리기
    for r = 0.2:0.2:1
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
    end
    
    % 방사선 그리기
    for i = 1:n_vars
        plot([0, cos(angles(i))], [0, sin(angles(i))], 'k:', ...
             'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
    end
    
    % 기준선 (전체 평균)
    plot(x_base, y_base, '--', 'Color', [0.5 0.5 0.5], 'LineWidth', 1.5);
    
    % 데이터 플롯
    patch(x_data, y_data, color, 'FaceAlpha', 0.3, 'EdgeColor', color, 'LineWidth', 2);
    
    % 데이터 포인트
    scatter(x_data(1:end-1), y_data(1:end-1), 30, color, 'filled');
    
    % 레이블
    label_radius = 1.15;
    for i = 1:n_vars
        [lx, ly] = pol2cart(angles(i), label_radius);
        text(lx, ly, sprintf('%s\n%.1f', labels{i}, data(i)), ...
             'HorizontalAlignment', 'center', 'FontSize', 8);
    end
    
    % 제목
    title(title_text, 'FontSize', 11, 'FontWeight', 'bold');
    
    axis equal;
    axis([-1.3 1.3 -1.3 1.3]);
    axis off;
    hold off;
end