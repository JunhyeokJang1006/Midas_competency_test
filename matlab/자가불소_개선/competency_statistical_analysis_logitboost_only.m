%% 개선된 HR 인재유형 분석 시스템 - LogitBoost 전용 버전
% 주요 특징:
% 1. 개별 레이더 차트 생성 (통일된 스케일)
% 2. 이진 분류 머신러닝 (LogitBoost만 사용)
% 3. 하이퍼파라미터 튜닝 및 교차검증
% 4. 고도화된 상관 기반 가중치 시스템
% 5. LogitBoost Feature Importance와 상관분석 통합
% 6. 학습된 모델 저장 및 재사용 기능

clear; clc; close all;

% 경고 메시지 억제 (선택적)
warning('off', 'stats:cvpartition:KFoldMissingGrp');
warning('off', 'stats:cvpartition:HoldOutMissingGrp');

%% ========================================================================
%                          PART 1: 초기 설정 및 데이터 로딩
% =========================================================================

fprintf('╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║    개선된 HR 인재유형 분석 시스템 - LogitBoost 버전      ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);
set(0, 'DefaultLineLineWidth', 2);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.output_dir = pwd;
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
config.model_file = fullfile(config.output_dir, 'trained_logitboost_model.mat');
config.use_saved_model = true;  % 저장된 모델 사용 여부

% 성과 순위 정의 (위장형 소화성 제외)
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 1]);

% 고성과자/저성과자 정의 (이진 분류용)
config.high_performers = {'자연성', '성실한 가연성', '유능한 불연성'};
config.low_performers = {'무능한 불연성', '소화성'};
config.excluded_types = {'위장형 소화성', '유익한 불연성', '게으른 가연성'}; % 중간 그룹 제외

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

% 빈 값 제거
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_idx}), :);

% 위장형 소화성 제외
excluded_mask = false(height(hr_clean), 1);
for i = 1:length(config.excluded_types)
    excluded_mask = excluded_mask | strcmp(hr_clean{:, talent_col_idx}, config.excluded_types{i});
end
hr_clean = hr_clean(~excluded_mask, :);

% 인재유형 분포 분석
talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('\n전체 인재유형 분포:\n');
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
        if length(valid_data) >= 5
            % 분산이 0인 경우도 처리
            data_var = var(valid_data);
            if (data_var > 0 || length(unique(valid_data)) > 1) && ...
               all(valid_data >= 0) && all(valid_data <= 100)
                valid_comp_cols{end+1} = comp_upper.Properties.VariableNames{i};
                valid_comp_indices(end+1) = i;
            end
        end
    end
end

if isempty(valid_comp_cols)
    error('유효한 역량 컬럼을 찾을 수 없습니다. 데이터를 확인해주세요.');
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
    fprintf('⚠ 종합점수 데이터 없음\n');
end

%% ========================================================================
%            PART 2: 개선된 레이더 차트 (개별 Figure, 통일 스케일)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║     PART 2: 개선된 레이더 차트 (통일 스케일)             ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 2.1 유형별 프로파일 계산 및 스케일 범위 설정
fprintf('【STEP 5】 유형별 프로파일 계산 및 스케일 설정\n');
fprintf('────────────────────────────────────────────\n');

unique_matched_types = unique(matched_talent_types);
n_types = length(unique_matched_types);

% 프로파일 계산
type_profiles = zeros(n_types, length(valid_comp_cols));
profile_stats = table();
profile_stats.TalentType = unique_matched_types;
profile_stats.Count = zeros(n_types, 1);
profile_stats.CompetencyMean = zeros(n_types, 1);
profile_stats.CompetencyStd = zeros(n_types, 1);
profile_stats.TotalScoreMean = zeros(n_types, 1);
profile_stats.PerformanceRank = zeros(n_types, 1);

for i = 1:n_types
    type_name = unique_matched_types{i};
    type_mask = strcmp(matched_talent_types, type_name);
    type_comp_data = matched_comp{type_mask, :};
    type_profiles(i, :) = nanmean(type_comp_data, 1);

    % 통계 정보 수집
    profile_stats.Count(i) = sum(type_mask);
    profile_stats.CompetencyMean(i) = nanmean(type_comp_data(:));
    profile_stats.CompetencyStd(i) = nanstd(type_comp_data(:));

    % 종합점수 통계
    if ~isempty(total_scores)
        type_total_scores = total_scores(type_mask);
        profile_stats.TotalScoreMean(i) = nanmean(type_total_scores);
    end

    % 성과 순위
    if config.performance_ranking.isKey(type_name)
        profile_stats.PerformanceRank(i) = config.performance_ranking(type_name);
    end
end

% 상위 12개 주요 역량 선정 (분산 기준)
comp_variance = var(table2array(matched_comp), 0, 1, 'omitnan');
[~, var_idx] = sort(comp_variance, 'descend');
top_comp_idx = var_idx(1:min(12, length(var_idx)));
top_comp_names = valid_comp_cols(top_comp_idx);

% 전체 평균 프로파일
overall_mean_profile = nanmean(table2array(matched_comp), 1);

% 통일된 스케일 범위 계산 (모든 유형의 최소/최대값)
all_profile_data = type_profiles(:, top_comp_idx);
global_min = min(all_profile_data(:)) - 5;  % 여유값 5점
global_max = max(all_profile_data(:)) + 5;  % 여유값 5점

fprintf('▶ 통일 스케일 범위: %.1f ~ %.1f\n', global_min, global_max);
fprintf('▶ 선정된 주요 역량: %d개\n', length(top_comp_idx));

%% 2.2 개별 레이더 차트 생성
fprintf('\n【STEP 6】 개별 레이더 차트 생성\n');
fprintf('────────────────────────────────────────────\n');

% 컬러맵 설정
colors = lines(n_types);

for i = 1:n_types
    % 새로운 Figure 창 생성
    fig = figure('Position', [100 + (i-1)*50, 100 + (i-1)*30, 800, 800], ...
                 'Color', 'white', ...
                 'Name', sprintf('인재유형: %s', unique_matched_types{i}));

    % 해당 유형의 프로파일 데이터
    type_profile = type_profiles(i, top_comp_idx);
    baseline = overall_mean_profile(top_comp_idx);

    % 개선된 레이더 차트 그리기
    createEnhancedRadarChart(type_profile, baseline, top_comp_names, ...
                            unique_matched_types{i}, colors(i,:), ...
                            global_min, global_max);

    % 추가 정보 표시
    if config.performance_ranking.isKey(unique_matched_types{i})
        perf_rank = config.performance_ranking(unique_matched_types{i});
        text(0.5, -0.05, sprintf('성과순위: %d', perf_rank), ...
             'Units', 'normalized', ...
             'HorizontalAlignment', 'center', ...
             'FontWeight', 'bold', 'FontSize', 14);
    end

    % Figure 저장
    saveas(fig, sprintf('radar_chart_%s_%s.png', ...
           strrep(unique_matched_types{i}, ' ', '_'), config.timestamp));

    fprintf('  ✓ %s 차트 생성 완료\n', unique_matched_types{i});
end

%% ========================================================================
%                    PART 3: 고도화된 상관 기반 가중치 분석
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║              PART 3: 고도화된 상관 기반 가중치 분석      ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 3.1 성과점수 기반 상관분석
fprintf('【STEP 7】 성과점수 기반 상관분석\n');
fprintf('────────────────────────────────────────────\n');

% 각 개인의 성과점수 할당
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    type_name = matched_talent_types{i};
    if config.performance_ranking.isKey(type_name)
        performance_scores(i) = config.performance_ranking(type_name);
    end
end

% 유효한 데이터만 선택
valid_perf_idx = performance_scores > 0;
valid_performance = performance_scores(valid_perf_idx);
valid_competencies = table2array(matched_comp(valid_perf_idx, :));

fprintf('▶ 성과점수 할당 완료: %d명\n', sum(valid_perf_idx));

%% 3.2 역량별 상관계수 계산
fprintf('\n【STEP 8】 역량-성과 상관분석\n');
fprintf('────────────────────────────────────────────\n');

% 상관계수 계산
n_competencies = length(valid_comp_cols);
correlation_results = table();
correlation_results.Competency = valid_comp_cols';
correlation_results.Correlation = zeros(n_competencies, 1);
correlation_results.PValue = zeros(n_competencies, 1);
correlation_results.Significance = cell(n_competencies, 1);
correlation_results.HighPerf_Mean = zeros(n_competencies, 1);
correlation_results.LowPerf_Mean = zeros(n_competencies, 1);
correlation_results.Difference = zeros(n_competencies, 1);
correlation_results.EffectSize = zeros(n_competencies, 1);

% 성과 상위/하위 그룹 분류 (상위 25%, 하위 25%)
perf_q75 = quantile(valid_performance, 0.75);
perf_q25 = quantile(valid_performance, 0.25);
high_perf_idx = valid_performance >= perf_q75;
low_perf_idx = valid_performance <= perf_q25;

for i = 1:n_competencies
    comp_scores = valid_competencies(:, i);
    valid_idx = ~isnan(comp_scores);

    if sum(valid_idx) >= 10
        % 상관계수 계산 (Spearman)
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

        % 그룹별 평균
        correlation_results.HighPerf_Mean(i) = nanmean(comp_scores(high_perf_idx));
        correlation_results.LowPerf_Mean(i) = nanmean(comp_scores(low_perf_idx));
        correlation_results.Difference(i) = correlation_results.HighPerf_Mean(i) - ...
                                           correlation_results.LowPerf_Mean(i);

        % Effect Size (Cohen's d)
        high_scores = comp_scores(high_perf_idx & valid_idx);
        low_scores = comp_scores(low_perf_idx & valid_idx);
        if length(high_scores) > 1 && length(low_scores) > 1
            pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                              (length(low_scores)-1)*var(low_scores)) / ...
                              (length(high_scores) + length(low_scores) - 2));
            correlation_results.EffectSize(i) = (mean(high_scores) - mean(low_scores)) / pooled_std;
        end
    end
end

% 가중치 계산
positive_corr = max(0, correlation_results.Correlation);
weights_corr = positive_corr / (sum(positive_corr) + eps);
correlation_results.Weight = weights_corr * 100;

correlation_results = sortrows(correlation_results, 'Correlation', 'descend');

% 다중 비교 보정 (Bonferroni)
correlation_results.PValue_Corrected = correlation_results.PValue * n_competencies;
correlation_results.PValue_Corrected = min(correlation_results.PValue_Corrected, 1);

fprintf('\n상위 10개 성과 예측 역량:\n');
fprintf('%-25s | 상관계수 | p-값 | 효과크기 | 가중치(%%)\n', '역량');
fprintf('%s\n', repmat('-', 75, 1));

for i = 1:min(10, height(correlation_results))
    fprintf('%-25s | %8.4f%s | %6.4f | %8.2f | %7.2f\n', ...
        correlation_results.Competency{i}, ...
        correlation_results.Correlation(i), correlation_results.Significance{i}, ...
        correlation_results.PValue_Corrected(i), ...
        correlation_results.EffectSize(i), ...
        correlation_results.Weight(i));
end

%% 3.3 상관분석 시각화
% Figure 2: 상관분석 결과
colors_vis = struct('primary', [0.2, 0.4, 0.8], 'secondary', [0.8, 0.3, 0.2], ...
               'tertiary', [0.3, 0.7, 0.4], 'gray', [0.5, 0.5, 0.5]);

fig2 = figure('Position', [100, 100, 1600, 900], 'Color', 'white');

% 상위 15개 역량의 상관계수와 가중치
subplot(2, 2, [1, 2]);
top_15 = correlation_results(1:min(15, height(correlation_results)), :);
x = 1:height(top_15);

yyaxis left
bar(x, top_15.Correlation, 'FaceColor', colors_vis.primary, 'EdgeColor', 'none');
ylabel('상관계수', 'FontSize', 12, 'FontWeight', 'bold');
ylim([-0.2, max(top_15.Correlation)*1.2]);

yyaxis right
plot(x, top_15.Weight, '-o', 'Color', colors_vis.secondary, 'LineWidth', 2, ...
     'MarkerFaceColor', colors_vis.secondary, 'MarkerSize', 8);
ylabel('가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');

set(gca, 'XTick', x, 'XTickLabel', top_15.Competency, 'XTickLabelRotation', 45);
xlabel('역량 항목', 'FontSize', 12, 'FontWeight', 'bold');
title('역량-성과 상관분석 및 가중치', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
box off;

% 누적 가중치
subplot(2, 2, [3 ,4]);
cumulative_weight = cumsum(correlation_results.Weight);
plot(cumulative_weight, 'LineWidth', 2.5, 'Color', colors_vis.tertiary);
hold on;
plot([1, length(cumulative_weight)], [80, 80], '--', 'Color', colors_vis.gray, 'LineWidth', 2);
xlabel('역량 개수', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('누적 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('누적 설명력 분석', 'FontSize', 14, 'FontWeight', 'bold');
legend('누적 가중치', '80% 기준선', 'Location', 'southeast', 'FontSize', 10);
grid on;
box off;

sgtitle('역량-성과 상관분석 종합 결과', 'FontSize', 16, 'FontWeight', 'bold');

%% ========================================================================
%         PART 4: LogitBoost 이진 분류 머신러닝
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║         PART 4: LogitBoost 이진 분류 머신러닝            ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 4.1 이진 레이블 생성
fprintf('【STEP 9】 고성과자/저성과자 이진 레이블 생성\n');
fprintf('────────────────────────────────────────────\n');

% 이진 분류용 데이터 필터링
binary_mask = false(length(matched_talent_types), 1);
binary_labels = zeros(length(matched_talent_types), 1);

for i = 1:length(matched_talent_types)
    type = matched_talent_types{i};

    if ismember(type, config.high_performers)
        binary_mask(i) = true;
        binary_labels(i) = 1;  % 고성과자
    elseif ismember(type, config.low_performers)
        binary_mask(i) = true;
        binary_labels(i) = 0;  % 저성과자
    end
end

% 이진 분류용 데이터 추출
X_binary = table2array(matched_comp(binary_mask, :));
y_binary = binary_labels(binary_mask);

fprintf('▶ 고성과자: %d명\n', sum(y_binary == 1));
fprintf('▶ 저성과자: %d명\n', sum(y_binary == 0));
fprintf('▶ 제외된 중간그룹: %d명\n', sum(~binary_mask));

% 최소 데이터 수 확인
min_class_size = min(sum(y_binary == 1), sum(y_binary == 0));
if min_class_size < 10
    warning('클래스 당 샘플 수가 너무 적습니다 (최소: %d). 결과가 불안정할 수 있습니다.', min_class_size);
end

% 데이터 정규화 (NaN 처리 포함)
X_binary_norm = X_binary;
for col = 1:size(X_binary, 2)
    col_data = X_binary(:, col);
    if any(~isnan(col_data))
        col_min = min(col_data(~isnan(col_data)));
        col_max = max(col_data(~isnan(col_data)));
        if col_max > col_min
            X_binary_norm(:, col) = (col_data - col_min) / (col_max - col_min);
        else
            X_binary_norm(:, col) = 0.5;  % 모든 값이 같으면 0.5로 설정
        end
    end
end

% NaN을 평균값으로 대체
for col = 1:size(X_binary_norm, 2)
    nan_idx = isnan(X_binary_norm(:, col));
    if any(nan_idx)
        col_mean = mean(X_binary_norm(~nan_idx, col));
        if isnan(col_mean)
            col_mean = 0.5;
        end
        X_binary_norm(nan_idx, col) = col_mean;
    end
end

%% 4.2 교차검증 설정
fprintf('\n【STEP 10】 교차검증 설정\n');
fprintf('────────────────────────────────────────────\n');

% 최종 테스트 세트 분리 (고정)
try
    test_partition = cvpartition(y_binary, 'HoldOut', 0.2, 'Stratify', true);
X_train_final = X_binary_norm(test_partition.training, :);
y_train_final = y_binary(test_partition.training);
X_test_final = X_binary_norm(test_partition.test, :);
y_test_final = y_binary(test_partition.test);
catch ME
    % Stratify가 실패하면 일반 HoldOut 사용
    fprintf('⚠ Stratified HoldOut 실패. 일반 HoldOut 사용\n');
    test_partition = cvpartition(y_binary, 'HoldOut', 0.2);
    X_train_final = X_binary_norm(test_partition.training, :);
    y_train_final = y_binary(test_partition.training);
    X_test_final = X_binary_norm(test_partition.test, :);
    y_test_final = y_binary(test_partition.test);
end

% 5-fold 교차검증 (훈련 세트에 대해서만)
k_folds = min(5, min(sum(y_train_final == 0), sum(y_train_final == 1)));  % 클래스 크기에 따라 조정
if k_folds < 2
    fprintf('⚠ 데이터가 부족하여 교차검증을 수행할 수 없습니다.\n');
    cv = [];
else
    try
        cv = cvpartition(y_train_final, 'KFold', k_folds, 'Stratify', true);
    catch
        cv = cvpartition(y_train_final, 'KFold', k_folds);
    end
end

fprintf('▶ 테스트 세트 분리 완료 (훈련: %d명, 테스트: %d명)\n', ...
    length(y_train_final), length(y_test_final));
fprintf('▶ %d-fold 교차검증 설정 완료\n', k_folds);

%% 4.3 LogitBoost 모델 학습 또는 로드
fprintf('\n【STEP 11】 LogitBoost 모델 준비\n');
fprintf('────────────────────────────────────────────\n');

% 저장된 모델 확인 및 로드
model_loaded = false;
if config.use_saved_model && exist(config.model_file, 'file')
    try
        fprintf('▶ 저장된 모델 로드 시도 중...\n');
        saved_data = load(config.model_file);
        
        % 저장된 모델의 특징 수와 현재 데이터의 특징 수 비교
        if size(saved_data.X_train_info, 2) == size(X_train_final, 2)
            final_logit_model = saved_data.final_logit_model;
            best_logit_params = saved_data.best_logit_params;
            logit_importance = saved_data.logit_importance;
            model_training_date = saved_data.training_date;
            model_loaded = true;
            
            fprintf('✓ 모델 로드 성공!\n');
            fprintf('  - 학습 날짜: %s\n', model_training_date);
            fprintf('  - 하이퍼파라미터:\n');
            fprintf('    • NumLearningCycles: %d\n', best_logit_params.NumCycles);
            fprintf('    • LearnRate: %.2f\n', best_logit_params.LearnRate);
            fprintf('    • MaxNumSplits: %d\n', best_logit_params.MaxSplits);
            fprintf('    • CV 정확도: %.3f (±%.3f)\n', best_logit_params.CV_Accuracy, best_logit_params.CV_Std);
        else
            fprintf('⚠ 저장된 모델의 특징 수가 일치하지 않습니다. 새로 학습합니다.\n');
        end
    catch ME
        fprintf('⚠ 모델 로드 실패: %s\n', ME.message);
        fprintf('  새로운 모델을 학습합니다.\n');
    end
end

% 모델이 로드되지 않았으면 새로 학습
if ~model_loaded
    fprintf('\n【STEP 11-1】 LogitBoost 하이퍼파라미터 튜닝\n');
    fprintf('────────────────────────────────────────────\n');
    
    % 하이퍼파라미터 그리드
    num_cycles_grid = [50, 100, 150, 200];
    learn_rate_grid = [0.1, 0.2, 0.3, 0.5];
    max_splits_grid = [10, 20, 30, 40];
    
    best_logit_accuracy = 0;
    best_logit_params = struct();
    
    fprintf('▶ 그리드 서치 진행 중...\n');
    total_combinations = length(num_cycles_grid) * length(learn_rate_grid) * length(max_splits_grid);
    current_combo = 0;
    
    for nc = num_cycles_grid
        for lr = learn_rate_grid
            for ms = max_splits_grid
                current_combo = current_combo + 1;
                cv_accuracies = zeros(k_folds, 1);
                
                for fold = 1:k_folds
                    train_idx = training(cv, fold);
                    val_idx = test(cv, fold);
                    
                    X_train = X_train_final(train_idx, :);
                    y_train = y_train_final(train_idx);
                    X_val = X_train_final(val_idx, :);
                    y_val = y_train_final(val_idx);
                    
                    try
                        model = fitcensemble(X_train, y_train, ...
                            'Method', 'LogitBoost', ...
                            'NumLearningCycles', nc, ...
                            'LearnRate', lr, ...
                            'Learners', templateTree('MaxNumSplits', ms));
                        
                        y_pred = predict(model, X_val);
                        cv_accuracies(fold) = sum(y_pred == y_val) / length(y_val);
                    catch ME
                        fprintf('  ⚠ Fold %d 학습 실패: %s\n', fold, ME.message);
                        cv_accuracies(fold) = 0;
                    end
                end
                
                mean_accuracy = mean(cv_accuracies(cv_accuracies > 0));  % 0이 아닌 값들만 평균
                
                if mean_accuracy > best_logit_accuracy && ~isnan(mean_accuracy)
                    best_logit_accuracy = mean_accuracy;
                    best_logit_params.NumCycles = nc;
                    best_logit_params.LearnRate = lr;
                    best_logit_params.MaxSplits = ms;
                    best_logit_params.CV_Accuracy = mean_accuracy;
                    best_logit_params.CV_Std = std(cv_accuracies(cv_accuracies > 0));
                end
                
                if mod(current_combo, 10) == 0
                    fprintf('  진행률: %d/%d (%.1f%%)\n', current_combo, total_combinations, ...
                        current_combo/total_combinations*100);
                end
            end
        end
    end
    
    % 최적 파라미터가 없으면 기본값 사용
    if isempty(fieldnames(best_logit_params))
        fprintf('⚠ 최적 파라미터를 찾지 못했습니다. 기본값을 사용합니다.\n');
        best_logit_params.NumCycles = 100;
        best_logit_params.LearnRate = 0.1;
        best_logit_params.MaxSplits = 20;
        best_logit_params.CV_Accuracy = NaN;
        best_logit_params.CV_Std = NaN;
    end
    
    fprintf('\n✓ 최적 하이퍼파라미터 결정:\n');
    fprintf('  - NumLearningCycles: %d\n', best_logit_params.NumCycles);
    fprintf('  - LearnRate: %.2f\n', best_logit_params.LearnRate);
    fprintf('  - MaxNumSplits: %d\n', best_logit_params.MaxSplits);
    if ~isnan(best_logit_params.CV_Accuracy)
        fprintf('  - CV 정확도: %.3f (±%.3f)\n', best_logit_params.CV_Accuracy, best_logit_params.CV_Std);
    end
    
    fprintf('\n【STEP 11-2】 최종 LogitBoost 모델 학습\n');
    fprintf('────────────────────────────────────────────\n');
    
    try
        final_logit_model = fitcensemble(X_train_final, y_train_final, ...
            'Method', 'LogitBoost', ...
            'NumLearningCycles', best_logit_params.NumCycles, ...
            'LearnRate', best_logit_params.LearnRate, ...
            'Learners', templateTree('MaxNumSplits', best_logit_params.MaxSplits));
        
        % Feature Importance 추출
        logit_importance = predictorImportance(final_logit_model);
        
        fprintf('✓ 최종 모델 학습 완료\n');
        
        % 모델 저장
        fprintf('\n▶ 학습된 모델 저장 중...\n');
        training_date = datestr(now, 'yyyy-mm-dd HH:MM:SS');
        X_train_info = X_train_final;  % 데이터 정보 저장
        
        save(config.model_file, 'final_logit_model', 'best_logit_params', ...
             'logit_importance', 'training_date', 'X_train_info', ...
             'valid_comp_cols');
        fprintf('✓ 모델 저장 완료: %s\n', config.model_file);
        
    catch ME
        error('모델 학습 실패: %s', ME.message);
    end
end

%% 4.4 최종 성능 평가
fprintf('\n【STEP 12】 최종 모델 성능 평가\n');
fprintf('────────────────────────────────────────────\n');

% 테스트 세트 예측
try
    y_pred_test = predict(final_logit_model, X_test_final);
    test_accuracy = sum(y_pred_test == y_test_final) / length(y_test_final);
catch ME
    fprintf('⚠ 예측 실패: %s\n', ME.message);
    y_pred_test = zeros(size(y_test_final));
    test_accuracy = 0;
end

% 혼동 행렬 계산
confusion_mat = confusionmat(y_test_final, y_pred_test);
precision = confusion_mat(2,2) / (confusion_mat(1,2) + confusion_mat(2,2) + eps);
recall = confusion_mat(2,2) / (confusion_mat(2,1) + confusion_mat(2,2) + eps);
f1_score = 2 * (precision * recall) / (precision + recall + eps);

fprintf('\n최종 테스트 세트 성능:\n');
fprintf('┌─────────────────┬──────────┐\n');
fprintf('│      지표       │    값    │\n');
fprintf('├─────────────────┼──────────┤\n');
fprintf('│ 정확도          │  %.3f   │\n', test_accuracy);
fprintf('│ 정밀도          │  %.3f   │\n', precision);
fprintf('│ 재현율          │  %.3f   │\n', recall);
fprintf('│ F1-Score        │  %.3f   │\n', f1_score);
fprintf('└─────────────────┴──────────┘\n');

fprintf('\n혼동 행렬:\n');
fprintf('                예측\n');
fprintf('              저성과  고성과\n');
fprintf('실제 저성과  %6d  %6d\n', confusion_mat(1,1), confusion_mat(1,2));
fprintf('     고성과  %6d  %6d\n', confusion_mat(2,1), confusion_mat(2,2));

%% ========================================================================
%              PART 5: 통합 가중치 시스템 (상관분석 + LogitBoost)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║     PART 5: 통합 가중치 시스템 (상관분석 + LogitBoost)   ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 5.1 Feature Importance 정규화 및 통합
fprintf('【STEP 13】 Feature Importance 통합\n');
fprintf('────────────────────────────────────────────\n');

% 벡터 차원 확인 및 조정
n_features = length(valid_comp_cols);

% LogitBoost importance 정규화
if ~exist('logit_importance', 'var') || isempty(logit_importance)
    fprintf('⚠ Feature importance가 없습니다. 균등 가중치를 사용합니다.\n');
    logit_importance = ones(1, n_features) / n_features;
end

% logit_importance를 올바른 크기로 조정
logit_importance = logit_importance(:);  % 열 벡터로 변환
if length(logit_importance) ~= n_features
    fprintf('⚠ Feature importance 크기 조정 중... (%d -> %d)\n', length(logit_importance), n_features);
    temp_importance = zeros(n_features, 1);
    min_len = min(length(logit_importance), n_features);
    temp_importance(1:min_len) = logit_importance(1:min_len);
    logit_importance = temp_importance;
end
logit_importance_norm = logit_importance / (sum(logit_importance) + eps);

% 상관분석 가중치 정규화
corr_weights = zeros(n_features, 1);
for i = 1:n_features
    comp_name = valid_comp_cols{i};
    idx = find(strcmp(correlation_results.Competency, comp_name), 1);
    if ~isempty(idx)
        corr_weights(i) = correlation_results.Weight(idx) / 100;
    end
end
corr_weights = corr_weights / (sum(corr_weights) + eps);

% 최종 통합 가중치 (상관분석 40% + LogitBoost 60%)
final_weights = 0.4 * corr_weights + 0.6 * logit_importance_norm;
final_weights = final_weights / (sum(final_weights) + eps);

%% 5.2 통합 가중치 비교 테이블 생성
weight_comparison = table();
weight_comparison.Competency = valid_comp_cols(:);  % 열 벡터로 확인
weight_comparison.Correlation = corr_weights * 100;
weight_comparison.LogitBoost = logit_importance_norm * 100;
weight_comparison.Final = final_weights * 100;

% 최종 가중치 기준 정렬
weight_comparison = sortrows(weight_comparison, 'Final', 'descend');

fprintf('\n상위 15개 역량 통합 가중치 비교:\n');
fprintf('%-25s | 상관(%%) | LogitBoost(%%) | 최종(%%)\n', '역량');
fprintf('%s\n', repmat('─', 65, 1));

for i = 1:min(15, height(weight_comparison))
    fprintf('%-25s | %7.2f | %13.2f | %8.2f\n', ...
        weight_comparison.Competency{i}, ...
        weight_comparison.Correlation(i), ...
        weight_comparison.LogitBoost(i), ...
        weight_comparison.Final(i));
end

%% 5.3 통합 시각화
fprintf('\n【STEP 14】 통합 결과 시각화\n');
fprintf('────────────────────────────────────────────\n');

figure('Position', [100, 100, 1600, 1000], 'Color', 'white');

% 1. LogitBoost 성능 지표
subplot(2, 3, 1);
metrics = [test_accuracy, precision, recall, f1_score];
metric_names = {'정확도', '정밀도', '재현율', 'F1-Score'};
bar(metrics * 100, 'FaceColor', [0.3, 0.6, 0.9]);
set(gca, 'XTickLabel', metric_names, 'XTickLabelRotation', 45);
ylabel('성능 (%)');
title('LogitBoost 모델 성능 지표');
ylim([0, 100]);
grid on;
for i = 1:length(metrics)
    text(i, metrics(i)*100 + 2, sprintf('%.1f%%', metrics(i)*100), ...
        'HorizontalAlignment', 'center', 'FontSize', 10);
end

% 2. 상위 15개 역량의 방법별 가중치 비교
subplot(2, 3, [2, 3]);
top_15 = weight_comparison(1:min(15, height(weight_comparison)), :);
if ismember('LogitBoost', top_15.Properties.VariableNames)
    bar_data = [top_15.Correlation, top_15.LogitBoost];
    legend_labels = {'상관분석', 'LogitBoost'};
else
    bar_data = top_15.Correlation;
    legend_labels = {'상관분석'};
end

bar(bar_data);
set(gca, 'XTickLabel', top_15.Competency, 'XTickLabelRotation', 45);
ylabel('가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('상관분석 vs LogitBoost Feature Importance (상위 15개)', 'FontSize', 14, 'FontWeight', 'bold');
legend(legend_labels, 'Location', 'northeast');
grid on;

% 3. 상관분석 vs LogitBoost 산점도
subplot(2, 3, 4);
if ismember('LogitBoost', weight_comparison.Properties.VariableNames)
    scatter(weight_comparison.Correlation, weight_comparison.LogitBoost, ...
            50, 'filled', 'MarkerFaceColor', [0.3, 0.6, 0.9]);
    xlabel('상관분석 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
    ylabel('LogitBoost 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
    title('상관분석 vs LogitBoost 가중치', 'FontSize', 14, 'FontWeight', 'bold');
    grid on;
    hold on;
    max_val = max([weight_comparison.Correlation; weight_comparison.LogitBoost]);
    plot([0, max_val], [0, max_val], 'r--', 'LineWidth', 1.5);
else
    bar(weight_comparison.Correlation(1:min(10, height(weight_comparison))));
    xlabel('상위 10개 역량', 'FontSize', 12, 'FontWeight', 'bold');
    ylabel('상관분석 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
    title('상관분석 가중치', 'FontSize', 14, 'FontWeight', 'bold');
    grid on;
end

% 4. 교차검증 성능
subplot(2, 3, 5);
if isfield(best_logit_params, 'CV_Accuracy') && ~isnan(best_logit_params.CV_Accuracy)
    cv_results = [best_logit_params.CV_Accuracy*100, test_accuracy*100];
    cv_labels = {'CV 정확도', '테스트 정확도'};
else
    cv_results = test_accuracy*100;
    cv_labels = {'테스트 정확도'};
end
bar(cv_results, 'FaceColor', [0.8, 0.3, 0.3]);
set(gca, 'XTickLabel', cv_labels);
ylabel('정확도 (%)');
title('모델 성능', 'FontSize', 12);
ylim([0, 100]);
grid on;
for i = 1:length(cv_results)
    if ~isnan(cv_results(i))
        text(i, cv_results(i) + 2, sprintf('%.1f%%', cv_results(i)), ...
            'HorizontalAlignment', 'center', 'FontSize', 10);
    end
end

% 5. 최종 통합 가중치
subplot(2, 3, 6);
n_top = min(15, height(weight_comparison));
top_15_final = weight_comparison(1:n_top, :);
barh(n_top:-1:1, top_15_final.Final, 'FaceColor', [0.2, 0.7, 0.4]);
set(gca, 'YTick', 1:n_top, 'YTickLabel', flip(top_15_final.Competency));
xlabel('최종 통합 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('최종 역량 중요도 순위', 'FontSize', 14, 'FontWeight', 'bold');
grid on;

sgtitle('HR 인재유형 분석 결과: 상관분석 + LogitBoost', 'FontSize', 16, 'FontWeight', 'bold');

%% ========================================================================
%                      PART 6: 결과 저장 및 최종 보고서
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║          PART 6: 결과 저장 및 최종 종합 보고서           ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 6.1 Excel 보고서 생성
fprintf('【STEP 15】 Excel 종합 보고서 생성\n');
fprintf('────────────────────────────────────────────\n');

output_filename = sprintf('hr_analysis_logitboost_%s.xlsx', config.timestamp);

try
    % Sheet 1: 인재유형 프로파일
    writetable(profile_stats, output_filename, 'Sheet', '인재유형프로파일');

    % Sheet 2: 상관분석 결과
    writetable(correlation_results, output_filename, 'Sheet', '상관분석결과');

    % Sheet 3: 모델 성능
    model_performance = table();
    model_performance.Metric = {'정확도'; '정밀도'; '재현율'; 'F1-Score'; 'CV_정확도'};
    model_performance.Value = [test_accuracy; precision; recall; f1_score; best_logit_params.CV_Accuracy];
    writetable(model_performance, output_filename, 'Sheet', '모델성능');

    % Sheet 4: 통합 가중치 비교
    writetable(weight_comparison, output_filename, 'Sheet', '통합가중치비교');

    % Sheet 5: 하이퍼파라미터
    hyperparams = table();
    hyperparams.Parameter = {'NumLearningCycles'; 'LearnRate'; 'MaxNumSplits'; 'CV_Accuracy'; 'CV_Std'};
    hyperparams.Value = {best_logit_params.NumCycles; best_logit_params.LearnRate; ...
                         best_logit_params.MaxSplits; best_logit_params.CV_Accuracy; ...
                         best_logit_params.CV_Std};
    writetable(hyperparams, output_filename, 'Sheet', '하이퍼파라미터');

    fprintf('✓ Excel 종합 보고서 저장 완료: %s\n', output_filename);

catch ME
    fprintf('⚠ Excel 저장 실패: %s\n', ME.message);
end

%% 6.2 MATLAB 파일 저장
analysis_results = struct();
analysis_results.profile_stats = profile_stats;
analysis_results.correlation_results = correlation_results;
analysis_results.model = final_logit_model;
analysis_results.performance = struct('accuracy', test_accuracy, 'precision', precision, ...
                                      'recall', recall, 'f1_score', f1_score);
analysis_results.weights = weight_comparison;
analysis_results.hyperparams = best_logit_params;
analysis_results.config = config;
analysis_results.confusion_matrix = confusion_mat;

save(sprintf('hr_analysis_logitboost_%s.mat', config.timestamp), 'analysis_results');
fprintf('✓ MATLAB 파일 저장 완료\n');

%% 6.3 최종 종합 보고서
fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                  최종 종합 분석 보고서                    ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

fprintf('📊 데이터 요약\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • 전체 매칭 데이터: %d명\n', length(matched_ids));
fprintf('  • 이진 분류 데이터: %d명 (고성과자 %d, 저성과자 %d)\n', ...
        length(y_binary), sum(y_binary==1), sum(y_binary==0));
fprintf('  • 역량 항목: %d개\n', length(valid_comp_cols));

fprintf('\n🤖 LogitBoost 모델 성능\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • 테스트 정확도: %.1f%%\n', test_accuracy*100);
fprintf('  • 정밀도: %.1f%%\n', precision*100);
fprintf('  • 재현율: %.1f%%\n', recall*100);
fprintf('  • F1-Score: %.1f%%\n', f1_score*100);
if isfield(best_logit_params, 'CV_Accuracy') && ~isnan(best_logit_params.CV_Accuracy)
    fprintf('  • 교차검증 정확도: %.1f%% (±%.1f%%)\n', ...
            best_logit_params.CV_Accuracy*100, best_logit_params.CV_Std*100);
end

fprintf('\n⭐ 핵심 역량 Top 5 (최종 통합 가중치)\n');
fprintf('────────────────────────────────────────────\n');
for i = 1:min(5, height(weight_comparison))
    if ismember('LogitBoost', weight_comparison.Properties.VariableNames)
        fprintf('  %d. %-25s: %5.2f%% (상관: %4.1f%%, ML: %4.1f%%)\n', i, ...
                weight_comparison.Competency{i}, ...
                weight_comparison.Final(i), ...
                weight_comparison.Correlation(i), ...
                weight_comparison.LogitBoost(i));
    else
        fprintf('  %d. %-25s: %5.2f%% (상관: %4.1f%%)\n', i, ...
                weight_comparison.Competency{i}, ...
                weight_comparison.Final(i), ...
                weight_comparison.Correlation(i));
    end
end

fprintf('\n📈 방법론 통합 성과\n');
fprintf('────────────────────────────────────────────\n');

% 상관분석과 ML의 일치도 계산
top5_corr_idx = 1:min(5, height(correlation_results));
top5_corr = correlation_results.Competency(top5_corr_idx);
if ismember('LogitBoost', weight_comparison.Properties.VariableNames)
    [~, ml_idx] = sort(weight_comparison.LogitBoost, 'descend');
    top5_ml = weight_comparison.Competency(ml_idx(1:min(5, length(ml_idx))));
    agreement = length(intersect(top5_corr, top5_ml));
else
    top5_ml = {};
    agreement = 0;
end

fprintf('  • 상위 5개 역량 일치도: %d/5 (%.0f%%)\n', agreement, agreement*20);
fprintf('  • 모델 신뢰도: ');
if test_accuracy > 0.75
    fprintf('높음 (실무 즉시 적용 가능)\n');
elseif test_accuracy > 0.65
    fprintf('중간 (보조 도구로 활용)\n');
else
    fprintf('낮음 (추가 데이터 필요)\n');
end

fprintf('\n✅ 실무 적용 권장사항\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  1. 1차 스크리닝: %s, %s, %s\n', ...
        weight_comparison.Competency{1}, ...
        weight_comparison.Competency{2}, ...
        weight_comparison.Competency{3});
fprintf('  2. 정밀 평가: 상위 5개 역량 통합 점수 활용\n');
fprintf('  3. 통합 가중치: 상관분석(40%%) + LogitBoost(60%%) 적용\n');
fprintf('  4. 업데이트 주기: 분기별 재학습 권장\n');

fprintf('\n🔬 기술적 특징\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • LogitBoost 앙상블 학습 (Boosting 기법)\n');
fprintf('  • 하이퍼파라미터 그리드 서치 최적화\n');
fprintf('  • 5-fold 교차검증\n');
fprintf('  • 다중 비교 보정 (Bonferroni)\n');
fprintf('  • Effect Size 기반 실질적 유의성 평가\n');

fprintf('\n════════════════════════════════════════════════════════════\n');
fprintf('           분석 완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('════════════════════════════════════════════════════════════\n\n');

%% Helper Function: 개선된 레이더 차트
function createEnhancedRadarChart(data, baseline, labels, title_text, color, min_val, max_val)
    % 개선된 레이더 차트 생성 (통일된 스케일 적용)
    n_vars = length(data);
    angles = linspace(0, 2*pi, n_vars+1);

    % 스케일 정규화 (통일된 범위 사용)
    data_norm = (data - min_val) / (max_val - min_val);
    baseline_norm = (baseline - min_val) / (max_val - min_val);

    % 순환을 위해 첫 번째 값을 마지막에 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];

    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);

    hold on;

    % 그리드 그리기 (5단계)
    grid_levels = 5;
    for i = 1:grid_levels
        r = i / grid_levels;
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);

        % 그리드 레이블 (실제 값으로 표시)
        grid_value = min_val + (max_val - min_val) * r;
        text(0, r, sprintf('%.0f', grid_value), ...
             'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'bottom', ...
             'FontSize', 9, 'Color', [0.4 0.4 0.4]);
    end

    % 방사선 그리기
    for i = 1:n_vars
        plot([0, cos(angles(i))], [0, sin(angles(i))], 'k:', ...
             'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);
    end

    % 기준선 (전체 평균)
    plot(x_base, y_base, '--', 'Color', [0.7 0.7 0.7], 'LineWidth', 2);

    % 데이터 플롯
    patch(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2.5);

    % 데이터 포인트
    scatter(x_data(1:end-1), y_data(1:end-1), 60, color, 'filled', ...
            'MarkerEdgeColor', 'white', 'LineWidth', 1);

    % 레이블 및 값
    label_radius = 1.25;
    for i = 1:n_vars
        [lx, ly] = pol2cart(angles(i), label_radius);

        % 차이값 계산
        diff_val = data(i) - baseline(i);
        diff_str = sprintf('%+.1f', diff_val);
        if diff_val > 0
            diff_color = [0, 0.5, 0];
        else
            diff_color = [0.8, 0, 0];
        end

        text(lx, ly, sprintf('%s\n%.1f\n(%s)', labels{i}, data(i), diff_str), ...
             'HorizontalAlignment', 'center', ...
             'FontSize', 10, 'FontWeight', 'bold');
    end

    % 제목
    title(title_text, 'FontSize', 16, 'FontWeight', 'bold');

    % 범례
    legend({'평균선', '해당 유형'}, 'Location', 'best', 'FontSize', 10);

    axis equal;
    axis([-1.5 1.5 -1.5 1.5]);
    axis off;
    hold off;
end
