%% ========================================================================
%                   소화성 인재 특성 분석 전용 코드
%          (One-Class SVM + Logistic Regression + t-test)
% =========================================================================
% 목적: 소화성 인재만의 고유한 역량 프로파일 파악
% 방법: 1) One-Class 모델로 소화성 패턴 학습
%       2) 소화성 vs 나머지 t-test 비교
%       3) 소화성 판별 가중치 도출
% =========================================================================

clear; clc; close all;
rng(42);

%% ========================================================================
%                          PART 1: 기본 설정 및 데이터 로딩
% =========================================================================

fprintf('\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║            소화성 인재 특성 분석 시스템                  ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 1.1 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\소화성분석';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 출력 폴더 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

% 한글 폰트 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');

%% 1.2 데이터 로딩
fprintf('【STEP 1】 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

hr_data = readtable(config.hr_file, 'VariableNamingRule', 'preserve');
comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');

fprintf('  ✓ HR 데이터: %d명\n', height(hr_data));
fprintf('  ✓ 역량 데이터: %d명\n', height(comp_upper));

%% 1.3 신뢰가능성 필터링
reliability_col = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col)
    unreliable = strcmp(comp_upper{:, reliability_col}, '신뢰불가');
    comp_upper = comp_upper(~unreliable, :);
    fprintf('  신뢰불가 제외: %d명 → %d명\n', sum(unreliable), height(comp_upper));
end

%% 1.4 역량 컬럼 추출
fprintf('\n【STEP 2】 역량 데이터 추출\n');
fprintf('────────────────────────────────────────────\n');

valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)
    col_name = comp_upper.Properties.VariableNames{i};
    col_data = comp_upper{:, i};
    
    if isnumeric(col_data) && ~all(isnan(col_data))
        valid_data = col_data(~isnan(col_data));
        if length(valid_data) >= 5 && var(valid_data) > 0
            valid_comp_cols{end+1} = col_name;
            valid_comp_indices(end+1) = i;
        end
    end
end

fprintf('  ✓ 유효 역량: %d개\n', length(valid_comp_cols));

%% 1.5 인재유형 매칭
fprintf('\n【STEP 3】 인재유형 매칭\n');
fprintf('────────────────────────────────────────────\n');

talent_col = find(contains(hr_data.Properties.VariableNames, '인재유형'), 1);
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col}), :);

% 위장형 소화성 제외
excluded = strcmp(hr_clean{:, talent_col}, '위장형 소화성');
hr_clean = hr_clean(~excluded, :);

% ID 매칭
hr_ids = arrayfun(@(x) sprintf('%.0f', x), hr_clean.ID, 'UniformOutput', false);
comp_ids = arrayfun(@(x) sprintf('%.0f', x), comp_upper{:, 1}, 'UniformOutput', false);

[matched_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_ids);

matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, valid_comp_indices);
matched_types = matched_hr{:, talent_col};

fprintf('  ✓ 매칭 성공: %d명\n', length(matched_ids));

%% 1.6 결측치 처리 (완전한 케이스만 사용)
X_raw = table2array(matched_comp);
complete_cases = ~any(isnan(X_raw), 2);

X_clean = X_raw(complete_cases, :);
types_clean = matched_types(complete_cases);

fprintf('  결측치 제거: %d명 → %d명\n', size(X_raw, 1), size(X_clean, 1));

%% ========================================================================
%                    PART 2: 소화성 vs 나머지 그룹 분석
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║              PART 2: 소화성 vs 나머지 비교 분석          ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 2.1 그룹 분리
fprintf('【STEP 4】 소화성 vs 나머지 그룹 분리\n');
fprintf('────────────────────────────────────────────\n');

% 소화성 그룹
sohwa_idx = strcmp(types_clean, '소화성');
normal_idx = ~sohwa_idx;

X_sohwa = X_clean(sohwa_idx, :);
X_normal = X_clean(normal_idx, :);

n_sohwa = sum(sohwa_idx);
n_normal = sum(normal_idx);

fprintf('  소화성: %d명 (%.1f%%)\n', n_sohwa, n_sohwa/(n_sohwa+n_normal)*100);
fprintf('  나머지: %d명 (%.1f%%)\n', n_normal, n_normal/(n_sohwa+n_normal)*100);

% 인재유형별 분포
unique_types = unique(types_clean);
fprintf('\n  인재유형 분포:\n');
for i = 1:length(unique_types)
    count = sum(strcmp(types_clean, unique_types{i}));
    fprintf('    - %s: %d명\n', unique_types{i}, count);
end

%% 2.2 역량별 t-test 비교
fprintf('\n【STEP 5】 역량별 t-test 분석\n');
fprintf('────────────────────────────────────────────\n');

ttest_results = table();
ttest_results.Competency = valid_comp_cols';
ttest_results.Sohwa_Mean = nanmean(X_sohwa, 1)';
ttest_results.Sohwa_Std = nanstd(X_sohwa, 0, 1)';
ttest_results.Normal_Mean = nanmean(X_normal, 1)';
ttest_results.Normal_Std = nanstd(X_normal, 0, 1)';
ttest_results.Mean_Diff = ttest_results.Sohwa_Mean - ttest_results.Normal_Mean;

n_comps = length(valid_comp_cols);
ttest_results.t_stat = zeros(n_comps, 1);
ttest_results.p_value = zeros(n_comps, 1);
ttest_results.Cohen_d = zeros(n_comps, 1);
ttest_results.Significance = cell(n_comps, 1);

for i = 1:n_comps
    sohwa_scores = X_sohwa(:, i);
    normal_scores = X_normal(:, i);
    
    % t-test
    [h, p, ~, stats] = ttest2(sohwa_scores, normal_scores);
    ttest_results.t_stat(i) = stats.tstat;
    ttest_results.p_value(i) = p;
    
    % Cohen's d
    pooled_std = sqrt(((length(sohwa_scores)-1)*var(sohwa_scores) + ...
                      (length(normal_scores)-1)*var(normal_scores)) / ...
                      (length(sohwa_scores) + length(normal_scores) - 2));
    ttest_results.Cohen_d(i) = ttest_results.Mean_Diff(i) / pooled_std;
    
    % 유의성 표시
    if p < 0.001
        ttest_results.Significance{i} = '***';
    elseif p < 0.01
        ttest_results.Significance{i} = '**';
    elseif p < 0.05
        ttest_results.Significance{i} = '*';
    else
        ttest_results.Significance{i} = '';
    end
end

% Cohen's d 절대값으로 정렬
ttest_results = sortrows(ttest_results, 'Cohen_d', 'ascend');  % 음수가 큰 순서

fprintf('\n소화성 특징 역량 (Cohen''s d < -0.5 & p < 0.05):\n');
fprintf('%-25s | 소화성 | 나머지 | 차이 | Cohen''s d | p-value\n', '역량');
fprintf('%s\n', repmat('-', 80, 1));

significant_idx = ttest_results.Cohen_d < -0.5 & ttest_results.p_value < 0.05;

for i = 1:sum(significant_idx)
    row = ttest_results(i, :);
    fprintf('%-25s | %6.1f | %6.1f | %+5.1f | %+7.3f | %.4f%s\n', ...
        row.Competency{1}, ...
        row.Sohwa_Mean, row.Normal_Mean, row.Mean_Diff, ...
        row.Cohen_d, row.p_value, row.Significance{1});
end

%% ========================================================================
%                    PART 3: One-Class 모델 학습
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║              PART 3: One-Class 모델 학습                  ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 3.1 One-Class SVM
fprintf('【STEP 6】 One-Class SVM 학습\n');
fprintf('────────────────────────────────────────────\n');

% 표준화
X_sohwa_z = zscore(X_sohwa);

% One-Class SVM (이상치 비율 10%)
try
    ocsvm_model = fitcsvm(X_sohwa_z, ones(n_sohwa, 1), ...
        'KernelFunction', 'rbf', ...
        'KernelScale', 'auto', ...
        'Nu', 0.1, ...  % 이상치 비율
        'Standardize', false);  % 이미 표준화됨
    
    fprintf('  ✓ One-Class SVM 학습 완료\n');
    
    % 전체 데이터에 적용
    X_all_z = zscore(X_clean);
    [labels, scores] = predict(ocsvm_model, X_all_z);
    
    % 소화성 탐지율
    sohwa_detected = sum(labels(sohwa_idx) == 1);
    normal_rejected = sum(labels(normal_idx) == -1);
    
    fprintf('  소화성 정탐지: %d/%d (%.1f%%)\n', ...
        sohwa_detected, n_sohwa, sohwa_detected/n_sohwa*100);
    fprintf('  정상 오탐지: %d/%d (%.1f%%)\n', ...
        normal_rejected, n_normal, normal_rejected/n_normal*100);
    
    ocsvm_success = true;
catch ME
    fprintf('  ⚠ One-Class SVM 실패: %s\n', ME.message);
    ocsvm_success = false;
end

%% 3.2 One-Class Logistic Regression (소화성 vs 나머지)
fprintf('\n【STEP 7】 소화성 판별 로지스틱 회귀\n');
fprintf('────────────────────────────────────────────\n');

% 레이블 생성 (소화성=1, 나머지=0)
y_binary = double(sohwa_idx);

% 표준화
X_all_z = zscore(X_clean);

% 클래스 가중치 (소화성이 적으므로 높은 가중치)
class_weights = [sum(y_binary==1)/length(y_binary), ...
                 sum(y_binary==0)/length(y_binary)];
sample_weights = zeros(size(y_binary));
sample_weights(y_binary == 1) = 1 / class_weights(1);
sample_weights(y_binary == 0) = 1 / class_weights(2);

% 로지스틱 회귀 학습
try
    logit_model = fitclinear(X_all_z, y_binary, ...
        'Learner', 'logistic', ...
        'Regularization', 'lasso', ...
        'Lambda', 1e-4, ...
        'Weights', sample_weights);
    
    fprintf('  ✓ 로지스틱 회귀 학습 완료\n');
    
    % 계수 추출
    coefficients = logit_model.Beta;
    
    % 소화성 특징 가중치 (음수 계수 = 소화성에서 낮음)
    sohwa_weights = -coefficients;  % 역방향
    sohwa_weights(sohwa_weights < 0) = 0;  % 양수만
    
    if sum(sohwa_weights) > 0
        sohwa_weights = sohwa_weights / sum(sohwa_weights) * 100;
    end
    
    % 성능 평가
    [pred_labels, pred_scores] = predict(logit_model, X_all_z);
    
    TP = sum(pred_labels == 1 & y_binary == 1);
    TN = sum(pred_labels == 0 & y_binary == 0);
    FP = sum(pred_labels == 1 & y_binary == 0);
    FN = sum(pred_labels == 0 & y_binary == 1);
    
    precision = TP / (TP + FP);
    recall = TP / (TP + FN);
    f1 = 2 * (precision * recall) / (precision + recall);
    
    fprintf('  정밀도: %.3f | 재현율: %.3f | F1: %.3f\n', precision, recall, f1);
    
    logit_success = true;
catch ME
    fprintf('  ⚠ 로지스틱 회귀 실패: %s\n', ME.message);
    logit_success = false;
end

%% ========================================================================
%                    PART 4: 소화성 판별 가중치 도출
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║            PART 4: 소화성 판별 가중치 도출               ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 4.1 가중치 통합
fprintf('【STEP 8】 소화성 판별 가중치 통합\n');
fprintf('────────────────────────────────────────────\n');

% 방법 1: t-test 효과크기 기반
ttest_weights = abs(ttest_results.Cohen_d);
ttest_weights(ttest_weights < 0) = 0;
ttest_weights = ttest_weights / sum(ttest_weights) * 100;

% 방법 2: 로지스틱 회귀 기반 (성공 시)
if logit_success
    logit_weights = sohwa_weights;
else
    logit_weights = zeros(size(ttest_weights));
end

% 앙상블 가중치 (두 방법 평균)
if logit_success
    ensemble_weights = (ttest_weights + logit_weights) / 2;
else
    ensemble_weights = ttest_weights;
end

% 결과 테이블
weight_table = table();
weight_table.Competency = ttest_results.Competency;
weight_table.Ttest_Weight = ttest_weights;
if logit_success
    weight_table.Logit_Weight = logit_weights;
end
weight_table.Ensemble_Weight = ensemble_weights;
weight_table.Sohwa_Mean = ttest_results.Sohwa_Mean;
weight_table.Normal_Mean = ttest_results.Normal_Mean;

% 앙상블 가중치로 정렬
weight_table = sortrows(weight_table, 'Ensemble_Weight', 'descend');

fprintf('\n소화성 판별 핵심 역량 (가중치 > 5%%):\n');
fprintf('%-25s | 가중치 | 소화성 | 나머지 | 차이\n', '역량');
fprintf('%s\n', repmat('-', 70, 1));

top_weights = weight_table(weight_table.Ensemble_Weight > 5, :);
for i = 1:height(top_weights)
    fprintf('%-25s | %6.2f%% | %6.1f | %6.1f | %+5.1f\n', ...
        top_weights.Competency{i}, ...
        top_weights.Ensemble_Weight(i), ...
        top_weights.Sohwa_Mean(i), ...
        top_weights.Normal_Mean(i), ...
        top_weights.Sohwa_Mean(i) - top_weights.Normal_Mean(i));
end

%% 4.2 소화성 점수 계산
fprintf('\n【STEP 9】 소화성 점수 계산 및 임계값 설정\n');
fprintf('────────────────────────────────────────────\n');

% 가중 점수 계산
sohwa_scores = X_clean * (ensemble_weights / 100);

% ROC 분석
[X_roc, Y_roc, T_roc, AUC] = perfcurve(y_binary, sohwa_scores, 1);

% Youden's J로 최적 임계값
J = Y_roc - X_roc;
[~, opt_idx] = max(J);
optimal_threshold = T_roc(opt_idx);

fprintf('  AUC: %.3f\n', AUC);
fprintf('  최적 임계값: %.2f\n', optimal_threshold);
fprintf('    - 민감도: %.3f\n', Y_roc(opt_idx));
fprintf('    - 특이도: %.3f\n', 1 - X_roc(opt_idx));

% 실제 소화성 점수 분포
sohwa_group_scores = sohwa_scores(sohwa_idx);
normal_group_scores = sohwa_scores(normal_idx);

fprintf('\n  점수 분포:\n');
fprintf('    소화성: %.2f ± %.2f\n', mean(sohwa_group_scores), std(sohwa_group_scores));
fprintf('    나머지: %.2f ± %.2f\n', mean(normal_group_scores), std(normal_group_scores));

% 실용적 임계값 제안 (백분위 기준)
threshold_p90 = prctile(sohwa_group_scores, 90);  % 소화성 90% 포함
threshold_p75 = prctile(sohwa_group_scores, 75);  % 소화성 75% 포함

fprintf('\n  실용 임계값 제안:\n');
fprintf('    보수적 (P75): %.2f (소화성 75%% 탐지)\n', threshold_p75);
fprintf('    균형적 (P90): %.2f (소화성 90%% 탐지)\n', threshold_p90);

%% ========================================================================
%                          PART 5: 시각화
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                    PART 5: 시각화                         ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 5.1 종합 결과 Figure
fig = figure('Position', [100, 100, 1600, 1200], 'Color', 'white');

% [1] 역량별 평균 비교
subplot(2, 3, 1);
x = 1:length(valid_comp_cols);
bar(x, [ttest_results.Normal_Mean, ttest_results.Sohwa_Mean]);
set(gca, 'XTick', x, 'XTickLabel', ttest_results.Competency, 'XTickLabelRotation', 45);
legend({'나머지', '소화성'}, 'Location', 'best');
ylabel('평균 점수');
title('역량별 평균 비교');
grid on;

% [2] Cohen's d (효과크기)
subplot(2, 3, 2);
barh(1:length(valid_comp_cols), ttest_results.Cohen_d);
set(gca, 'YTick', 1:length(valid_comp_cols), 'YTickLabel', ttest_results.Competency);
xlabel('Cohen''s d');
title('소화성 vs 나머지 효과크기');
xline(-0.5, '--r', 'LineWidth', 2);
xline(-0.8, '--r', 'LineWidth', 2);
grid on;

% [3] p-value
subplot(2, 3, 3);
bar(-log10(ttest_results.p_value));
set(gca, 'XTick', x, 'XTickLabel', ttest_results.Competency, 'XTickLabelRotation', 45);
ylabel('-log10(p-value)');
title('통계적 유의성');
yline(-log10(0.05), '--g', 'LineWidth', 2);
yline(-log10(0.01), '--', 'Color', [1, 0.5, 0], 'LineWidth', 2);
grid on;

% [4] 가중치 비교
subplot(2, 3, 4);
top_n = min(10, height(weight_table));
barh(1:top_n, weight_table.Ensemble_Weight(1:top_n));
set(gca, 'YTick', 1:top_n, 'YTickLabel', weight_table.Competency(1:top_n));
xlabel('가중치 (%)');
title('소화성 판별 가중치 (Top 10)');
grid on;

% [5] 점수 분포
subplot(2, 3, 5);
histogram(normal_group_scores, 20, 'FaceColor', [0.3, 0.7, 0.9], 'FaceAlpha', 0.7);
hold on;
histogram(sohwa_group_scores, 20, 'FaceColor', [0.9, 0.3, 0.3], 'FaceAlpha', 0.7);
xline(optimal_threshold, '--k', 'LineWidth', 2);
xline(threshold_p75, ':', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);
xlabel('소화성 점수');
ylabel('빈도');
title('소화성 점수 분포');
legend({'나머지', '소화성', '최적 임계값', 'P75 임계값'}, 'Location', 'best');
grid on;

% [6] ROC 곡선
subplot(2, 3, 6);
plot(X_roc, Y_roc, 'LineWidth', 3);
hold on;
plot([0, 1], [0, 1], '--k');
plot(X_roc(opt_idx), Y_roc(opt_idx), 'ro', 'MarkerSize', 10, 'LineWidth', 2);
xlabel('위양성률');
ylabel('민감도');
title(sprintf('ROC 곡선 (AUC=%.3f)', AUC));
legend({'ROC', '무작위', '최적점'}, 'Location', 'best');
grid on;
axis square;

sgtitle('소화성 인재 특성 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');

% 저장
saveas(fig, fullfile(config.output_dir, 'sohwa_analysis_results.png'));
fprintf('  ✓ 시각화 저장 완료\n');

%% ========================================================================
%                        PART 6: 결과 저장
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                  PART 6: 결과 저장                        ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 6.1 엑셀 저장
fprintf('【STEP 10】 엑셀 결과 파일 생성\n');
fprintf('────────────────────────────────────────────\n');

excel_file = fullfile(config.output_dir, 'sohwa_analysis_results.xlsx');

try
    % Sheet 1: 요약
    summary_table = table();
    summary_table.Item = {'분석일시'; '총샘플수'; '소화성수'; '나머지수'; 'AUC'; '최적임계값'; 'P75임계값'};
    summary_table.Value = {datestr(now); length(y_binary); n_sohwa; n_normal; ...
                          sprintf('%.3f', AUC); sprintf('%.2f', optimal_threshold); ...
                          sprintf('%.2f', threshold_p75)};
    writetable(summary_table, excel_file, 'Sheet', '요약');
    
    % Sheet 2: t-test 결과
    writetable(ttest_results, excel_file, 'Sheet', 't-test결과');
    
    % Sheet 3: 가중치
    writetable(weight_table, excel_file, 'Sheet', '소화성판별가중치');
    
    % Sheet 4: 개별 점수
    individual_results = table();
    individual_results.ID = (1:length(y_binary))';
    individual_results.TalentType = types_clean;
    individual_results.IsSohwa = y_binary;
    individual_results.SohwaScore = sohwa_scores;
    individual_results.Predicted = sohwa_scores >= optimal_threshold;
    writetable(individual_results, excel_file, 'Sheet', '개별점수');
    
    fprintf('  ✓ 엑셀 저장 완료: %s\n', excel_file);
catch ME
    fprintf('  ⚠ 엑셀 저장 실패: %s\n', ME.message);
end

%% 6.2 MAT 파일 저장
results = struct();
results.config = config;
results.ttest_results = ttest_results;
results.weight_table = weight_table;
results.threshold = struct('optimal', optimal_threshold, 'p75', threshold_p75, 'p90', threshold_p90);
results.performance = struct('AUC', AUC, 'precision', precision, 'recall', recall, 'f1', f1);
if ocsvm_success
    results.ocsvm_model = ocsvm_model;
end
if logit_success
    results.logit_model = logit_model;
end

mat_file = fullfile(config.output_dir, 'sohwa_analysis_results.mat');
save(mat_file, 'results');
fprintf('  ✓ MAT 저장 완료: %s\n', mat_file);

%% ========================================================================
%                          최종 요약 출력
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                    최종 요약                              ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

fprintf('【소화성 인재 특징】\n');
fprintf('────────────────────────────────────────────\n');
fprintf('샘플 수: %d명 (전체의 %.1f%%)\n\n', n_sohwa, n_sohwa/(n_sohwa+n_normal)*100);

fprintf('핵심 차별화 역량 (효과크기 큰 순서, p<0.05):\n');
sig_features = ttest_results(ttest_results.p_value < 0.05 & abs(ttest_results.Cohen_d) > 0.5, :);
for i = 1:min(5, height(sig_features))
    fprintf('  %d. %s\n', i, sig_features.Competency{i});
    fprintf('     - 소화성: %.1f점, 나머지: %.1f점 (차이: %.1f점)\n', ...
        sig_features.Sohwa_Mean(i), sig_features.Normal_Mean(i), sig_features.Mean_Diff(i));
    fprintf('     - 효과크기: %.3f\n\n', sig_features.Cohen_d(i));
end

fprintf('【소화성 판별 시스템】\n');
fprintf('────────────────────────────────────────────\n');
fprintf('가중치 상위 3개 역량:\n');
for i = 1:min(3, height(weight_table))
    fprintf('  %d. %s: %.2f%%\n', i, weight_table.Competency{i}, weight_table.Ensemble_Weight(i));
end

fprintf('\n판별 기준:\n');
fprintf('  • 보수적 (P75): %.2f점 이하 → 소화성 의심\n', threshold_p75);
fprintf('  • 최적 (Youden): %.2f점 이하 → 소화성 의심\n', optimal_threshold);
fprintf('  • 공격적 (P90): %.2f점 이하 → 소화성 의심\n\n', threshold_p90);

fprintf('성능 지표:\n');
fprintf('  • AUC: %.3f\n', AUC);
if logit_success
    fprintf('  • 정밀도: %.3f\n', precision);
    fprintf('  • 재현율: %.3f\n', recall);
    fprintf('  • F1 스코어: %.3f\n\n', f1);
end

fprintf('【활용 방법】\n');
fprintf('────────────────────────────────────────────\n');
fprintf('1. 역량 점수 확인\n');
fprintf('2. 가중치 곱셈: Σ(역량×가중치/100)\n');
fprintf('3. 임계값 비교:\n');
fprintf('   - %.2f점 이하: 소화성 리스크 높음 → 추가 확인\n', threshold_p75);
fprintf('   - %.2f-%.2f점: 경계선 → 면접 중점 확인\n', threshold_p75, threshold_p90);
fprintf('   - %.2f점 이상: 안전\n\n', threshold_p90);

fprintf('✅ 소화성 인재 특성 분석 완료!\n');
fprintf('📊 결과 저장 위치: %s\n\n', config.output_dir);