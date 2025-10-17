%% ========================================================================
%         소화성 인재 특성 분석 + 역량검사 가중치 통합 시스템 (Revised)
%          (Logistic Regression 기반 확률 조정 방식)
% =========================================================================
% 목적: 1) 소화성 인재의 역량 프로파일 분석
%       2) Logistic Regression으로 소화성 확률 계산
%       3) 기존 역량검사 가중치 점수에 확률 기반 조정 적용
%       4) Excel 적용 가능한 계수 테이블 출력
% =========================================================================
% 개선사항:
%   - 기존 역량검사 가중치 파일 연동
%   - Logistic 확률 기반 점수 조정 로직 추가
%   - Excel 수식 재현을 위한 계수 테이블 생성
%   - 실제 적용 예시 및 검증 시트 추가
% =========================================================================


clear; clc; close all;
rng(42, 'twister');

%% ========================================================================
%                          PART 1: 기본 설정 및 데이터 로딩
% =========================================================================

fprintf('\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║      소화성 인재 특성 분석 + 가중치 통합 시스템          ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 1.1 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\소화성분석';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 기존 역량검사 가중치 파일 (최신 파일 자동 탐색)
weight_files = dir('D:\project\HR데이터\결과\자가불소_revised\역량검사_가중치적용점수_*.xlsx');
if ~isempty(weight_files)
    [~, latest_idx] = max([weight_files.datenum]);
    config.weight_file = fullfile(weight_files(latest_idx).folder, weight_files(latest_idx).name);
    fprintf('📁 역량검사 가중치 파일: %s\n', weight_files(latest_idx).name);
else
    warning('역량검사 가중치 파일을 찾을 수 없습니다. 기본값 사용.');
    config.weight_file = '';
end

% 출력 폴더 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

% 한글 폰트 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');

%% 1.2 기존 역량검사 가중치 로드
fprintf('\n【STEP 1】 기존 역량검사 가중치 로드\n');
fprintf('────────────────────────────────────────────\n');

if ~isempty(config.weight_file) && exist(config.weight_file, 'file')
    try
        weight_info = readtable(config.weight_file, 'Sheet', '가중치_상세정보', 'VariableNamingRule', 'preserve');
        base_weights_table = weight_info(:, 1:2);
        base_weights_table.Properties.VariableNames = {'역량명', '가중치_퍼센트'};

        fprintf('  ✓ 가중치 로드 성공: %d개 역량\n', height(base_weights_table));
        disp(base_weights_table);
    catch ME
        warning('가중치 파일 읽기 실패: %s', ME.message);
        base_weights_table = table();
    end
else
    fprintf('  ⚠ 가중치 파일 없음, 분석만 진행\n');
    base_weights_table = table();
end

%% 1.3 데이터 로딩
fprintf('\n【STEP 2】 HR 및 역량 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

hr_data = readtable(config.hr_file, 'VariableNamingRule', 'preserve');
comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');

fprintf('  ✓ HR 데이터: %d명\n', height(hr_data));
fprintf('  ✓ 역량 데이터: %d명\n', height(comp_upper));

%% 1.4 신뢰가능성 필터링
reliability_col = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col)
    unreliable = strcmp(comp_upper{:, reliability_col}, '신뢰불가');
    comp_upper = comp_upper(~unreliable, :);
    fprintf('  신뢰불가 제외: %d명 → %d명\n', sum(unreliable), height(comp_upper));
end

%% 1.5 역량 컬럼 추출
fprintf('\n【STEP 3】 역량 데이터 추출\n');
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

%% 1.6 인재유형 매칭
fprintf('\n【STEP 4】 인재유형 매칭\n');
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

%% 1.7 결측치 처리
X_raw = table2array(matched_comp);
complete_cases = ~any(isnan(X_raw), 2);

X_clean = X_raw(complete_cases, :);
types_clean = matched_types(complete_cases);
ids_clean = matched_ids(complete_cases);

fprintf('  결측치 제거: %d명 → %d명\n', size(X_raw, 1), size(X_clean, 1));

%% ========================================================================
%                    PART 2: 소화성 vs 나머지 그룹 분석
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║              PART 2: 소화성 vs 나머지 비교 분석          ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 2.1 그룹 분리
fprintf('【STEP 5】 소화성 vs 나머지 그룹 분리\n');
fprintf('────────────────────────────────────────────\n');

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
fprintf('\n【STEP 6】 역량별 t-test 분석\n');
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

    [h, p, ~, stats] = ttest2(sohwa_scores, normal_scores);
    ttest_results.t_stat(i) = stats.tstat;
    ttest_results.p_value(i) = p;

    pooled_std = sqrt(((length(sohwa_scores)-1)*var(sohwa_scores) + ...
                      (length(normal_scores)-1)*var(normal_scores)) / ...
                      (length(sohwa_scores) + length(normal_scores) - 2));
    ttest_results.Cohen_d(i) = ttest_results.Mean_Diff(i) / pooled_std;

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

ttest_results = sortrows(ttest_results, 'Cohen_d', 'ascend');

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
%                    PART 3: Logistic Regression 모델 학습
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║       PART 3: Logistic Regression 소화성 판별 모델       ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 3.1 Logistic Regression
fprintf('【STEP 7】 소화성 판별 로지스틱 회귀 학습\n');
fprintf('────────────────────────────────────────────\n');

% 레이블 생성
y_binary = double(sohwa_idx);

% 표준화 (Z-score)
X_all_z = zscore(X_clean);
mu_comp = mean(X_clean, 1);
sigma_comp = std(X_clean, 0, 1);

% 클래스 가중치
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
    beta_coefficients = logit_model.Beta;
    beta_bias = logit_model.Bias;

    % 성능 평가
    [pred_labels, pred_scores] = predict(logit_model, X_all_z);

    % 확률 계산 (Sigmoid)
    logits = X_all_z * beta_coefficients + beta_bias;
    sohwa_probabilities = 1 ./ (1 + exp(-logits));

    % 혼동 행렬
    TP = sum(pred_labels == 1 & y_binary == 1);
    TN = sum(pred_labels == 0 & y_binary == 0);
    FP = sum(pred_labels == 1 & y_binary == 0);
    FN = sum(pred_labels == 0 & y_binary == 1);

    accuracy = (TP + TN) / length(y_binary);
    precision = TP / (TP + FP);
    recall = TP / (TP + FN);
    f1 = 2 * (precision * recall) / (precision + recall);

    fprintf('  정확도: %.3f | 정밀도: %.3f | 재현율: %.3f | F1: %.3f\n', ...
        accuracy, precision, recall, f1);

    logit_success = true;
catch ME
    fprintf('  ⚠ 로지스틱 회귀 실패: %s\n', ME.message);
    logit_success = false;
    beta_coefficients = [];
    beta_bias = 0;
    sohwa_probabilities = zeros(size(y_binary));
end

%% ========================================================================
%                PART 4: 기존 가중치 점수 + 소화성 확률 조정
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║        PART 4: 기존 가중치 점수 + 소화성 확률 조정       ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 4.1 기존 가중치 점수 계산
fprintf('【STEP 8】 기존 역량검사 가중치 점수 계산\n');
fprintf('────────────────────────────────────────────\n');

if ~isempty(base_weights_table)
    % 역량명 매칭
    base_scores = zeros(size(X_clean, 1), 1);
    matched_weights = zeros(1, n_comps);

    for i = 1:n_comps
        comp_name = valid_comp_cols{i};
        weight_idx = find(strcmp(base_weights_table.('역량명'), comp_name), 1);

        if ~isempty(weight_idx)
            matched_weights(i) = base_weights_table.('가중치_퍼센트')(weight_idx);
        end
    end

    % Z-score로 정규화 후 가중치 적용
    X_clean_z = zscore(X_clean);
    X_clean_z(isnan(X_clean_z)) = 0;

    weighted_scores = X_clean_z * (matched_weights' / 100);

    % 0-100 범위로 변환
    min_score = min(weighted_scores);
    max_score = max(weighted_scores);
    base_scores = 100 * (weighted_scores - min_score) / (max_score - min_score);

    fprintf('  ✓ 기존 가중치 점수 계산 완료\n');
    fprintf('    - 평균: %.2f ± %.2f\n', mean(base_scores), std(base_scores));
    fprintf('    - 범위: [%.2f, %.2f]\n', min(base_scores), max(base_scores));
else
    fprintf('  ⚠ 가중치 정보 없음, 균등 가중치 사용\n');
    base_scores = mean(X_clean, 2);
end

%% 4.2 소화성 확률 기반 점수 조정
fprintf('\n【STEP 9】 소화성 확률 기반 점수 조정\n');
fprintf('────────────────────────────────────────────\n');

if logit_success
    % 조정 강도 파라미터 (HR이 조정 가능)
    adjustment_strength = 0.5;  % 0~1 사이 (기본값 0.5)

    % 조정 점수 계산
    adjusted_scores = base_scores .* (1 - sohwa_probabilities * adjustment_strength);

    fprintf('  ✓ 조정 강도: %.2f\n', adjustment_strength);
    fprintf('  ✓ 조정 점수 계산 완료\n');
    fprintf('    - 조정 전 평균: %.2f ± %.2f\n', mean(base_scores), std(base_scores));
    fprintf('    - 조정 후 평균: %.2f ± %.2f\n', mean(adjusted_scores), std(adjusted_scores));

    % 소화성 그룹의 점수 변화
    sohwa_base_mean = mean(base_scores(sohwa_idx));
    sohwa_adjusted_mean = mean(adjusted_scores(sohwa_idx));
    normal_base_mean = mean(base_scores(normal_idx));
    normal_adjusted_mean = mean(adjusted_scores(normal_idx));

    fprintf('\n  그룹별 점수 변화:\n');
    fprintf('    소화성: %.2f → %.2f (%.2f점 감소)\n', ...
        sohwa_base_mean, sohwa_adjusted_mean, sohwa_base_mean - sohwa_adjusted_mean);
    fprintf('    나머지: %.2f → %.2f (%.2f점 감소)\n', ...
        normal_base_mean, normal_adjusted_mean, normal_base_mean - normal_adjusted_mean);
else
    adjusted_scores = base_scores;
    fprintf('  ⚠ Logistic 모델 없음, 조정 없이 기존 점수 사용\n');
end

%% ========================================================================
%                          PART 5: Excel 출력용 테이블 생성
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║            PART 5: Excel 출력용 테이블 생성               ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 5.1 Logistic 계수 테이블 (Excel 수식용)
fprintf('【STEP 10】 Logistic 계수 테이블 생성\n');
fprintf('────────────────────────────────────────────\n');

if logit_success
    logit_coef_table = table();
    logit_coef_table.('역량명') = ['절편(Bias)'; valid_comp_cols'];
    logit_coef_table.('Logistic계수') = [beta_bias; beta_coefficients];
    logit_coef_table.('설명') = [{'모델 절편'}; repmat({'역량별 계수'}, n_comps, 1)];

    % Z-score 변환 파라미터 추가
    zscore_params = table();
    zscore_params.('역량명') = valid_comp_cols';
    zscore_params.('평균') = mu_comp';
    zscore_params.('표준편차') = sigma_comp';

    fprintf('  ✓ Logistic 계수 테이블 생성 완료\n');
    fprintf('    - 절편: %.4f\n', beta_bias);
    fprintf('    - 계수 범위: [%.4f, %.4f]\n', min(beta_coefficients), max(beta_coefficients));
else
    logit_coef_table = table();
    zscore_params = table();
    fprintf('  ⚠ Logistic 모델 없음, 계수 테이블 생성 불가\n');
end

%% 5.2 개별 결과 테이블
fprintf('\n【STEP 11】 개별 결과 테이블 생성\n');
fprintf('────────────────────────────────────────────\n');

individual_results = table();
individual_results.('ID') = ids_clean;
individual_results.('인재유형') = types_clean;
individual_results.('소화성여부') = y_binary;

% 역량 점수 추가
for i = 1:n_comps
    individual_results.(valid_comp_cols{i}) = X_clean(:, i);
end

individual_results.('기존가중치점수') = base_scores;

if logit_success
    individual_results.('소화성확률') = sohwa_probabilities;
    individual_results.('조정후점수') = adjusted_scores;
    individual_results.('점수변화') = base_scores - adjusted_scores;
end

fprintf('  ✓ 개별 결과 테이블 생성 완료: %d명\n', height(individual_results));

%% 5.3 적용 가이드 테이블
fprintf('\n【STEP 12】 Excel 적용 가이드 생성\n');
fprintf('────────────────────────────────────────────\n');

application_guide = table();
application_guide.('단계') = {'1단계'; '2단계'; '3단계'; '4단계'; '5단계'};
application_guide.('작업내용') = {
    '역량 점수를 Z-score로 변환: (점수 - 평균) / 표준편차';
    'Logistic 계산: 절편 + Σ(Z-score × 계수)';
    '소화성 확률 계산: 1 / (1 + EXP(-Logistic값))';
    '기존 가중치 점수 계산';
    '조정 점수 = 기존점수 × (1 - 소화성확률 × 조정강도)'
};
application_guide.('Excel수식예시') = {
    '=(B2 - VLOOKUP("전략성", 정규화파라미터!A:B, 2)) / VLOOKUP("전략성", 정규화파라미터!A:C, 3)';
    '=VLOOKUP("절편", Logistic계수!A:B, 2) + SUMPRODUCT(Z점수범위, VLOOKUP(역량명범위, Logistic계수!A:B, 2))';
    '=1 / (1 + EXP(-D2))';
    '=SUMPRODUCT(역량점수범위, 가중치범위)';
    '=F2 * (1 - E2 * 0.5)'
};

fprintf('  ✓ 적용 가이드 생성 완료\n');

%% ========================================================================
%                          PART 6: 시각화
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                    PART 6: 시각화                         ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

fprintf('【STEP 13】 종합 결과 시각화\n');
fprintf('────────────────────────────────────────────\n');

fig = figure('Position', [100, 100, 1800, 1200], 'Color', 'white');

% [1] 역량별 평균 비교
subplot(2, 4, 1);
x = 1:length(valid_comp_cols);
bar(x, [ttest_results.Normal_Mean, ttest_results.Sohwa_Mean]);
set(gca, 'XTick', x, 'XTickLabel', ttest_results.Competency, 'XTickLabelRotation', 45);
legend({'나머지', '소화성'}, 'Location', 'best');
ylabel('평균 점수');
title('역량별 평균 비교');
grid on;

% [2] Cohen's d
subplot(2, 4, 2);
barh(1:length(valid_comp_cols), ttest_results.Cohen_d);
set(gca, 'YTick', 1:length(valid_comp_cols), 'YTickLabel', ttest_results.Competency);
xlabel('Cohen''s d');
title('효과크기 (소화성 vs 나머지)');
xline(-0.5, '--r', 'LineWidth', 1.5);
xline(-0.8, '--r', 'LineWidth', 2);
grid on;

% [3] Logistic 계수
if logit_success
    subplot(2, 4, 3);
    barh(1:n_comps, beta_coefficients);
    set(gca, 'YTick', 1:n_comps, 'YTickLabel', valid_comp_cols);
    xlabel('Logistic 계수');
    title('소화성 판별 계수');
    xline(0, '--k', 'LineWidth', 1);
    grid on;
end

% [4] 소화성 확률 분포
if logit_success
    subplot(2, 4, 4);
    histogram(sohwa_probabilities(normal_idx), 20, 'FaceColor', [0.3, 0.7, 0.9], 'FaceAlpha', 0.7);
    hold on;
    histogram(sohwa_probabilities(sohwa_idx), 20, 'FaceColor', [0.9, 0.3, 0.3], 'FaceAlpha', 0.7);
    xlabel('소화성 확률');
    ylabel('빈도');
    title('소화성 확률 분포');
    legend({'나머지', '소화성'}, 'Location', 'best');
    grid on;
end

% [5] 기존 점수 분포
subplot(2, 4, 5);
histogram(base_scores(normal_idx), 20, 'FaceColor', [0.3, 0.7, 0.9], 'FaceAlpha', 0.7);
hold on;
histogram(base_scores(sohwa_idx), 20, 'FaceColor', [0.9, 0.3, 0.3], 'FaceAlpha', 0.7);
xlabel('기존 가중치 점수');
ylabel('빈도');
title('기존 점수 분포');
legend({'나머지', '소화성'}, 'Location', 'best');
grid on;

% [6] 조정 후 점수 분포
if logit_success
    subplot(2, 4, 6);
    histogram(adjusted_scores(normal_idx), 20, 'FaceColor', [0.3, 0.7, 0.9], 'FaceAlpha', 0.7);
    hold on;
    histogram(adjusted_scores(sohwa_idx), 20, 'FaceColor', [0.9, 0.3, 0.3], 'FaceAlpha', 0.7);
    xlabel('조정 후 점수');
    ylabel('빈도');
    title('소화성 확률 조정 후 점수 분포');
    legend({'나머지', '소화성'}, 'Location', 'best');
    grid on;
end

% [7] 점수 변화 (Before vs After)
if logit_success
    subplot(2, 4, 7);
    scatter(base_scores(normal_idx), adjusted_scores(normal_idx), 50, 'o', 'MarkerFaceColor', [0.3, 0.7, 0.9], 'MarkerEdgeColor', 'none', 'MarkerFaceAlpha', 0.5);
    hold on;
    scatter(base_scores(sohwa_idx), adjusted_scores(sohwa_idx), 100, 's', 'MarkerFaceColor', [0.9, 0.3, 0.3], 'MarkerEdgeColor', 'k', 'LineWidth', 1.5);
    plot([0, 100], [0, 100], '--k', 'LineWidth', 1.5);
    xlabel('기존 점수');
    ylabel('조정 후 점수');
    title('점수 변화 (Before vs After)');
    legend({'나머지', '소화성', '변화없음'}, 'Location', 'best');
    grid on;
    axis equal;
    xlim([0, 100]);
    ylim([0, 100]);
end

% [8] 점수 변화량 분포
if logit_success
    subplot(2, 4, 8);
    score_change = base_scores - adjusted_scores;
    histogram(score_change(normal_idx), 20, 'FaceColor', [0.3, 0.7, 0.9], 'FaceAlpha', 0.7);
    hold on;
    histogram(score_change(sohwa_idx), 20, 'FaceColor', [0.9, 0.3, 0.3], 'FaceAlpha', 0.7);
    xlabel('점수 감소량 (기존 - 조정)');
    ylabel('빈도');
    title('소화성 확률 조정에 따른 점수 변화');
    legend({'나머지', '소화성'}, 'Location', 'best');
    grid on;
end

sgtitle('소화성 인재 분석 + 가중치 통합 결과', 'FontSize', 16, 'FontWeight', 'bold');

% 저장
fig_file = fullfile(config.output_dir, sprintf('sohwa_integrated_analysis_%s.png', config.timestamp));
saveas(fig, fig_file);
fprintf('  ✓ 시각화 저장: %s\n', fig_file);

%% ========================================================================
%                        PART 7: 결과 저장
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                  PART 7: 결과 저장                        ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 7.1 엑셀 저장
fprintf('【STEP 14】 Excel 결과 파일 생성\n');
fprintf('────────────────────────────────────────────\n');

excel_file = fullfile(config.output_dir, sprintf('sohwa_integrated_results_%s.xlsx', config.timestamp));

try
    % Sheet 1: 요약
    summary_table = table();
    summary_table.('항목') = {'분석일시'; '총샘플수'; '소화성수'; '나머지수'; ...
                          '기존점수평균_소화성'; '기존점수평균_나머지'; ...
                          '조정후평균_소화성'; '조정후평균_나머지'; ...
                          '조정강도'; 'Logistic정확도'; 'Logistic정밀도'; 'LogisticF1'};
    if logit_success
        summary_table.('값') = {datestr(now); length(y_binary); n_sohwa; n_normal; ...
                           sprintf('%.2f', sohwa_base_mean); sprintf('%.2f', normal_base_mean); ...
                           sprintf('%.2f', sohwa_adjusted_mean); sprintf('%.2f', normal_adjusted_mean); ...
                           sprintf('%.2f', adjustment_strength); ...
                           sprintf('%.3f', accuracy); sprintf('%.3f', precision); sprintf('%.3f', f1)};
    else
        summary_table.('값') = {datestr(now); length(y_binary); n_sohwa; n_normal; ...
                           sprintf('%.2f', mean(base_scores(sohwa_idx))); sprintf('%.2f', mean(base_scores(normal_idx))); ...
                           'N/A'; 'N/A'; 'N/A'; 'N/A'; 'N/A'; 'N/A'};
    end
    writetable(summary_table, excel_file, 'Sheet', '요약');

    % Sheet 2: t-test 결과
    writetable(ttest_results, excel_file, 'Sheet', 't-test결과');

    % Sheet 3: Logistic 계수
    if logit_success
        writetable(logit_coef_table, excel_file, 'Sheet', 'Logistic계수');
    end

    % Sheet 4: Z-score 정규화 파라미터
    if logit_success
        writetable(zscore_params, excel_file, 'Sheet', 'Z정규화파라미터');
    end

    % Sheet 5: 개별 결과
    writetable(individual_results, excel_file, 'Sheet', '개별결과');

    % Sheet 6: Excel 적용 가이드
    writetable(application_guide, excel_file, 'Sheet', 'Excel적용가이드');

    % Sheet 7: 기존 역량검사 가중치
    if ~isempty(base_weights_table)
        writetable(base_weights_table, excel_file, 'Sheet', '기존역량검사가중치');
    end

    fprintf('  ✓ Excel 저장 완료: %s\n', excel_file);
catch ME
    fprintf('  ⚠ Excel 저장 실패: %s\n', ME.message);
end

%% 7.2 MAT 파일 저장
results = struct();
results.config = config;
results.ttest_results = ttest_results;
results.individual_results = individual_results;
results.base_weights_table = base_weights_table;

if logit_success
    results.logit_model = logit_model;
    results.logit_coef_table = logit_coef_table;
    results.zscore_params = zscore_params;
    results.adjustment_strength = adjustment_strength;
    results.performance = struct('accuracy', accuracy, 'precision', precision, ...
                                 'recall', recall, 'f1', f1);
end

mat_file = fullfile(config.output_dir, sprintf('sohwa_integrated_results_%s.mat', config.timestamp));
save(mat_file, 'results');
fprintf('  ✓ MAT 저장 완료: %s\n', mat_file);

%% ========================================================================
%                          최종 요약 출력
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                    최종 요약                              ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

fprintf('【분석 결과】\n');
fprintf('────────────────────────────────────────────\n');
fprintf('총 샘플: %d명 (소화성: %d명, 나머지: %d명)\n\n', ...
    length(y_binary), n_sohwa, n_normal);

fprintf('기존 역량검사 가중치 점수:\n');
fprintf('  소화성: %.2f ± %.2f점\n', mean(base_scores(sohwa_idx)), std(base_scores(sohwa_idx)));
fprintf('  나머지: %.2f ± %.2f점\n', mean(base_scores(normal_idx)), std(base_scores(normal_idx)));
fprintf('  차이: %.2f점\n\n', abs(mean(base_scores(sohwa_idx)) - mean(base_scores(normal_idx))));

if logit_success
    fprintf('소화성 확률 조정 후 점수:\n');
    fprintf('  소화성: %.2f ± %.2f점 (%.2f점 감소)\n', ...
        mean(adjusted_scores(sohwa_idx)), std(adjusted_scores(sohwa_idx)), ...
        mean(base_scores(sohwa_idx)) - mean(adjusted_scores(sohwa_idx)));
    fprintf('  나머지: %.2f ± %.2f점 (%.2f점 감소)\n', ...
        mean(adjusted_scores(normal_idx)), std(adjusted_scores(normal_idx)), ...
        mean(base_scores(normal_idx)) - mean(adjusted_scores(normal_idx)));
    fprintf('  차이: %.2f점\n\n', abs(mean(adjusted_scores(sohwa_idx)) - mean(adjusted_scores(normal_idx))));

    fprintf('Logistic 모델 성능:\n');
    fprintf('  정확도: %.3f\n', accuracy);
    fprintf('  정밀도: %.3f\n', precision);
    fprintf('  재현율: %.3f\n', recall);
    fprintf('  F1 스코어: %.3f\n\n', f1);
end

fprintf('【주요 소화성 특징 역량 (Cohen''s d 상위 5개)】\n');
fprintf('────────────────────────────────────────────\n');
sig_features = ttest_results(ttest_results.p_value < 0.05 & abs(ttest_results.Cohen_d) > 0.5, :);
for i = 1:min(5, height(sig_features))
    fprintf('%d. %s\n', i, sig_features.Competency{i});
    fprintf('   - 소화성: %.1f점 vs 나머지: %.1f점 (차이: %.1f점)\n', ...
        sig_features.Sohwa_Mean(i), sig_features.Normal_Mean(i), sig_features.Mean_Diff(i));
    fprintf('   - Cohen''s d: %.3f, p=%.4f%s\n\n', ...
        sig_features.Cohen_d(i), sig_features.p_value(i), sig_features.Significance{i});
end

fprintf('【Excel 적용 방법】\n');
fprintf('────────────────────────────────────────────\n');
fprintf('1. "Logistic계수" 시트에서 계수 확인\n');
fprintf('2. "Z정규화파라미터" 시트에서 평균/표준편차 확인\n');
fprintf('3. 역량 점수를 Z-score로 변환\n');
fprintf('4. Logistic 값 계산: 절편 + Σ(Z-score × 계수)\n');
fprintf('5. 소화성 확률: 1/(1+EXP(-Logistic값))\n');
fprintf('6. 조정 점수: 기존점수 × (1 - 소화성확률 × 조정강도)\n\n');

fprintf('조정 강도 파라미터 (현재: %.2f):\n', adjustment_strength);
fprintf('  • 0.0 → 조정 없음 (기존 점수 유지)\n');
fprintf('  • 0.5 → 중간 조정 (권장)\n');
fprintf('  • 1.0 → 최대 조정 (소화성 확률 100%% 반영)\n\n');

fprintf('✅ 소화성 인재 분석 + 가중치 통합 완료!\n');
fprintf('📊 결과 저장 위치: %s\n', config.output_dir);
fprintf('📁 Excel 파일: %s\n\n', excel_file);
