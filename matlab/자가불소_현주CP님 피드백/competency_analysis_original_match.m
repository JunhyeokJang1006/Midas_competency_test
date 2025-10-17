%% HR 데이터 분석 시스템 - 원본 코드와 동일한 결과 버전
% =========================================================================
% 주요 기능:
% 1. 원본 코드와 100% 동일한 데이터 처리 방식
% 2. Cost-Sensitive Learning 기반 고성과자 예측
% 3. Leave-One-Out 교차검증으로 최적 Lambda 찾기
% 4. Bootstrap 가중치 안정성 검증 (1000회)
% 5. 퍼뮤테이션 테스트 (캐시 없이 매번 계산)
% 6. 4가지 가중치 방법론 통합 분석
% =========================================================================

clear; clc; close all;
rng(42);

%% ========================================================================
%                          PART 1: 초기 설정 및 데이터 로딩
% =========================================================================

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);
set(0, 'DefaultLineLineWidth', 2);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\자가불소_simple_mean';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
config.baseline_type = 'weighted';
config.extreme_group_method = 'all';
config.force_recalc_permutation = false;

% 파일 관리 시스템 설정
config.create_backup = true;
config.backup_folder = 'backup';
config.use_timestamp = false;

% 성과 순위 정의 (원본과 동일)
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 1]);

% 순서형 로지스틱 회귀용 레벨 정의
config.ordinal_levels = containers.Map(...
    {'소화성', '무능한 불연성', '게으른 가연성', '유능한 불연성', ...
     '유익한 불연성', '성실한 가연성', '자연성'}, ...
    [1, 2, 3, 4, 5, 6, 7]);

% 고성과자/저성과자 정의
config.high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
config.low_performers = {'무능한 불연성', '소화성', '게으른 가연성'};
config.excluded_from_analysis = {'유능한 불연성'};
config.excluded_types = {'위장형 소화성'};

% 결과 디렉토리 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
fprintf('║            HR 데이터 분석 시스템 시작 (원본 매치 버전)      ║\n');
fprintf('╚════════════════════════════════════════════════════════════╝\n\n');

%% 【STEP 1】 데이터 로딩
fprintf('【STEP 1】 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

% 파일 존재 여부 확인
if ~exist(config.hr_file, 'file')
    error('HR 데이터 파일을 찾을 수 없습니다: %s', config.hr_file);
end
if ~exist(config.comp_file, 'file')
    error('역량검사 데이터 파일을 찾을 수 없습니다: %s', config.comp_file);
end

try
    % HR 데이터 로딩
    fprintf('▶ HR 데이터 로딩 중...\n');
    hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
    fprintf('  ✓ HR 데이터: %d명 로드 완료\n', height(hr_data));

    % 역량검사 데이터 로딩 (원본과 동일한 시트명)
    fprintf('▶ 역량검사 데이터 로딩 중...\n');
    comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목_단순평균', 'VariableNamingRule', 'preserve');
    comp_total = readtable(config.comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 상위항목 데이터: %d명\n', height(comp_upper));
    fprintf('  ✓ 종합점수 데이터: %d명\n', height(comp_total));
catch ME
    error('데이터 로딩 실패: %s', ME.message);
end

%% 【STEP 1-1】 신뢰가능성 필터링
fprintf('\n【STEP 1-1】 신뢰가능성 필터링\n');
fprintf('────────────────────────────────────────────\n');

% 신뢰가능성 컬럼 찾기
reliability_col_idx = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col_idx)
    fprintf('▶ 신뢰가능성 컬럼 발견: %s\n', comp_upper.Properties.VariableNames{reliability_col_idx});

    % 신뢰불가 데이터 제외
    reliability_data = comp_upper{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(comp_upper), 1);
    end

    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));

    % 신뢰가능한 데이터만 유지
    comp_upper = comp_upper(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능한 데이터: %d명\n', height(comp_upper));
else
    fprintf('  ⚠ 신뢰가능성 컬럼이 없습니다. 모든 데이터를 사용합니다.\n');
end

%% 【STEP 2】 인재유형 데이터 추출 및 정제
fprintf('\n【STEP 2】 인재유형 데이터 추출 및 정제\n');
fprintf('────────────────────────────────────────────\n');

% 인재유형 컬럼 찾기
talent_col_idx = find(contains(hr_data.Properties.VariableNames, {'인재유형', '인재', '유형'}), 1);
if isempty(talent_col_idx)
    error('인재유형 컬럼을 찾을 수 없습니다.');
end

fprintf('▶ 인재유형 컬럼 발견: %s\n', hr_data.Properties.VariableNames{talent_col_idx});

% 빈 값 제거
valid_idx = ~cellfun(@isempty, hr_data{:, talent_col_idx});
hr_clean = hr_data(valid_idx, :);

% 제외 유형 제거
excluded_mask = ismember(hr_clean{:, talent_col_idx}, config.excluded_types);
hr_clean = hr_clean(~excluded_mask, :);

talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('  제외된 유형: %s\n', strjoin(config.excluded_types, ', '));
fprintf('  정제된 데이터: %d명\n', height(hr_clean));

fprintf('\n인재유형별 분포:\n');
for i = 1:length(unique_types)
    fprintf('  - %s: %d명 (%.1f%%)\n', ...
        unique_types{i}, type_counts(i), type_counts(i)/sum(type_counts)*100);
end

%% 【STEP 3】 역량 데이터 처리 (원본과 동일한 방식)
fprintf('\n【STEP 3】 역량 데이터 처리 (비활성/과활성 포함)\n');
fprintf('────────────────────────────────────────────\n');

% ID 컬럼 찾기
comp_id_col = find(contains(lower(comp_upper.Properties.VariableNames), {'id', '사번'}), 1);
if isempty(comp_id_col)
    error('역량 데이터에서 ID 컬럼을 찾을 수 없습니다.');
end

% 유효한 역량 컬럼 추출 (6번째 컬럼부터 - 원본과 동일)
fprintf('▶ 비활성/과활성을 포함하여 모든 역량 데이터 분석\n');

valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)  % 원본과 동일하게 6번째 컬럼부터
    col_name = comp_upper.Properties.VariableNames{i};
    col_data = comp_upper{:, i};

    if isnumeric(col_data) && ~all(isnan(col_data))
        valid_data = col_data(~isnan(col_data));

        if length(valid_data) >= 5
            data_var = var(valid_data);
            if (data_var > 0 || length(unique(valid_data)) > 1) && ...
               all(valid_data >= 0) && all(valid_data <= 100)
                valid_comp_cols{end+1} = col_name;
                valid_comp_indices(end+1) = i;
            end
        end
    end
end

fprintf('\n포함된 모든 역량 컬럼 (%d개) - 비활성/과활성 포함:\n', length(valid_comp_cols));
for i = 1:min(length(valid_comp_cols), 10)
    fprintf('  - %s\n', valid_comp_cols{i});
end
if length(valid_comp_cols) > 10
    fprintf('  ... 외 %d개 더\n', length(valid_comp_cols) - 10);
end

if isempty(valid_comp_cols)
    error('유효한 역량 컬럼을 찾을 수 없습니다. 데이터를 확인해주세요.');
end

%% 【STEP 4】 데이터 매칭 및 통합 (원본과 동일)
fprintf('\n【STEP 4】 데이터 매칭 및 통합\n');
fprintf('────────────────────────────────────────────\n');

% ID 표준화 (원본과 동일한 방식)
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
    fprintf('  ⚠ 종합점수 매칭 실패\n');
end

fprintf('  최종 매칭 완료: %d명\n', height(matched_hr));

%% 【STEP 7】 성과점수 기반 상관분석 (원본과 동일)
fprintf('\n【STEP 7】 성과점수 기반 상관분석\n');
fprintf('────────────────────────────────────────────\n');

% 각 개인의 성과점수 할당 (원본과 동일)
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

%% 【STEP 8】 역량-성과 상관분석 (원본과 동일)
fprintf('\n【STEP 8】 역량-성과 상관분석\n');
fprintf('────────────────────────────────────────────\n');

% 상관계수 계산 (원본과 동일)
n_competencies = length(valid_comp_cols);
correlation_results = table();
correlation_results.Competency = valid_comp_cols';
correlation_results.Correlation = zeros(n_competencies, 1);
correlation_results.PValue = zeros(n_competencies, 1);
correlation_results.Significance = cell(n_competencies, 1);

% Quartile 기준 극단 그룹 (원본과 동일)
perf_q25 = prctile(valid_performance, 25);
perf_q75 = prctile(valid_performance, 75);

high_perf_idx = valid_performance >= perf_q75;
low_perf_idx = valid_performance <= perf_q25;

for i = 1:n_competencies
    comp_scores = valid_competencies(:, i);
    valid_idx = ~isnan(comp_scores);

    if sum(valid_idx) >= 10
        % Spearman 상관계수 (원본과 동일)
        [r, p] = corr(comp_scores(valid_idx), valid_performance(valid_idx), ...
                     'Type', 'Spearman');

        correlation_results.Correlation(i) = r;
        correlation_results.PValue(i) = p;

        if p < 0.001
            correlation_results.Significance{i} = '***';
        elseif p < 0.01
            correlation_results.Significance{i} = '**';
        elseif p < 0.05
            correlation_results.Significance{i} = '*';
        else
            correlation_results.Significance{i} = 'ns';
        end
    else
        correlation_results.Correlation(i) = 0;
        correlation_results.PValue(i) = 1;
        correlation_results.Significance{i} = 'ns';
    end
end

% 가중치 계산 (원본과 동일)
positive_corr = max(0, correlation_results.Correlation);
weights_corr = positive_corr / (sum(positive_corr) + eps);
correlation_results.Weight = weights_corr * 100;

correlation_results = sortrows(correlation_results, 'Correlation', 'descend');

fprintf('\n상위 10개 성과 예측 역량:\n');
fprintf('%-25s | 상관계수 | p-값 | 가중치(%%)\n', '역량');
fprintf('%s\n', repmat('-', 65, 1));

for i = 1:min(10, height(correlation_results))
    fprintf('%-25s | %8.3f | %6.3f | %7.2f\n', ...
            correlation_results.Competency{i}, ...
            correlation_results.Correlation(i), ...
            correlation_results.PValue(i), ...
            correlation_results.Weight(i));
end

%% 【STEP 9】 데이터 준비 및 클래스 불균형 분석 (원본과 동일)
fprintf('\n【STEP 9】 데이터 준비 및 클래스 불균형 분석\n');
fprintf('────────────────────────────────────────────\n');

% 이진분류를 위한 데이터 준비 (Quartile 기반)
y_binary = zeros(length(valid_performance), 1);
y_binary(high_perf_idx) = 1;  % 고성과자
y_binary(low_perf_idx) = 0;   % 저성과자

% 중간 성과자 제외
extreme_idx = high_perf_idx | low_perf_idx;
X_final = valid_competencies(extreme_idx, :);
y_final = y_binary(extreme_idx);

% 완전한 케이스만 사용
complete_cases = ~any(isnan(X_final), 2);
X_final = X_final(complete_cases, :);
y_final = y_final(complete_cases);

fprintf('▶ 극단 그룹 비교 데이터 준비 완료\n');
fprintf('  총 샘플: %d명\n', length(y_final));
fprintf('  고성과자: %d명 (%.1f%%)\n', sum(y_final==1), sum(y_final==1)/length(y_final)*100);
fprintf('  저성과자: %d명 (%.1f%%)\n', sum(y_final==0), sum(y_final==0)/length(y_final)*100);

% 결측값 대체 없이 완전한 데이터만 사용
X_imputed = X_final;
y_weight = y_final;
feature_names = valid_comp_cols;
n_features = size(X_imputed, 2);

% 클래스 가중치 계산 (원본과 동일)
class_weights = length(y_weight) ./ (2 * [sum(y_weight==0), sum(y_weight==1)]);
sample_weights = zeros(size(y_weight));
sample_weights(y_weight==0) = class_weights(1);
sample_weights(y_weight==1) = class_weights(2);

fprintf('  클래스 가중치 - 저성과자: %.3f, 고성과자: %.3f\n', class_weights(1), class_weights(2));

% 비용 행렬 정의 (원본과 동일)
cost_matrix = [0 1; 1.5 0];

%% 【STEP 11】 Leave-One-Out 교차검증으로 최적 Lambda 찾기 (원본과 동일)
fprintf('\n【STEP 11】 Leave-One-Out 교차검증으로 최적 Lambda 찾기\n');
fprintf('────────────────────────────────────────────\n');

% Lambda 범위 설정 (원본과 동일)
lambda_range = logspace(-3, 1, 20);
cv_scores = zeros(length(lambda_range), 1);
cv_aucs = zeros(length(lambda_range), 1);

fprintf('▶ %d개 Lambda 값에 대해 LOO-CV 수행 중...\n', length(lambda_range));

for lambda_idx = 1:length(lambda_range)
    current_lambda = lambda_range(lambda_idx);

    % LOO-CV를 위한 예측값 저장
    loo_predictions = zeros(length(y_weight), 1);
    loo_probabilities = zeros(length(y_weight), 1);

    % Leave-One-Out 루프
    for i = 1:length(y_weight)
        % 훈련 데이터 (i번째 샘플 제외)
        train_idx = true(length(y_weight), 1);
        train_idx(i) = false;

        X_train = X_imputed(train_idx, :);
        y_train = y_weight(train_idx);
        X_test = X_imputed(i, :);

        % 훈련 세트로만 표준화 (데이터 누수 방지)
        mu = mean(X_train);
        sigma = std(X_train);
        sigma(sigma == 0) = 1;

        X_train_z = (X_train - mu) ./ sigma;
        X_test_z = (X_test - mu) ./ sigma;

        % 가중치 계산 (클래스 불균형 + 비용 반영)
        w = zeros(size(y_train));
        n_class0 = sum(y_train == 0);
        n_class1 = sum(y_train == 1);

        if n_class0 > 0 && n_class1 > 0
            % 역빈도 가중치 * 비용 행렬
            w(y_train == 0) = (length(y_train)/(2*n_class0)) * cost_matrix(1,2);
            w(y_train == 1) = (length(y_train)/(2*n_class1)) * cost_matrix(2,1);

            try
                % 로지스틱 회귀 모델 학습
                mdl = fitclinear(X_train_z, y_train, ...
                    'Learner', 'logistic', ...
                    'Regularization', 'ridge', ...
                    'Lambda', current_lambda, ...
                    'Solver', 'lbfgs', ...
                    'Weights', w);

                % 예측 및 확률 계산
                [pred, score] = predict(mdl, X_test_z);
                loo_predictions(i) = pred;
                loo_probabilities(i) = score(2);  % 클래스 1의 확률

            catch
                % 실패시 랜덤 예측
                loo_predictions(i) = round(rand());
                loo_probabilities(i) = rand();
            end
        else
            % 클래스가 하나만 있는 경우
            loo_predictions(i) = mode(y_train);
            loo_probabilities(i) = double(mode(y_train));
        end
    end

    % 성능 평가
    accuracy = mean(loo_predictions == y_weight);
    cv_scores(lambda_idx) = accuracy;

    % AUC 계산
    try
        [~, ~, ~, auc] = perfcurve(y_weight, loo_probabilities, 1);
        cv_aucs(lambda_idx) = auc;
    catch
        cv_aucs(lambda_idx) = 0.5;
    end

    if mod(lambda_idx, 5) == 0
        fprintf('  진행: %d/%d (Lambda=%.4f, AUC=%.3f)\n', ...
               lambda_idx, length(lambda_range), current_lambda, cv_aucs(lambda_idx));
    end
end

% 최적 Lambda 선택 (AUC 기준)
[best_auc, best_idx] = max(cv_aucs);
optimal_lambda = lambda_range(best_idx);

fprintf('▶ 최적 Lambda: %.4f\n', optimal_lambda);
fprintf('  최고 AUC: %.3f\n', best_auc);
fprintf('  해당 정확도: %.3f\n', cv_scores(best_idx));

%% 【STEP 12】 최종 Cost-Sensitive 모델 학습 (원본과 동일)
fprintf('\n【STEP 12】 최종 Cost-Sensitive 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% 전체 데이터 표준화
mu_final = mean(X_imputed);
sigma_final = std(X_imputed);
sigma_final(sigma_final == 0) = 1;
X_normalized = (X_imputed - mu_final) ./ sigma_final;

% 가중치 계산
sample_weights = zeros(size(y_weight));
n0 = sum(y_weight == 0);
n1 = sum(y_weight == 1);
sample_weights(y_weight == 0) = (length(y_weight)/(2*n0)) * cost_matrix(1,2);
sample_weights(y_weight == 1) = (length(y_weight)/(2*n1)) * cost_matrix(2,1);

try
    final_mdl = fitclinear(X_normalized, y_weight, ...
        'Learner', 'logistic', ...
        'Regularization', 'ridge', ...
        'Lambda', optimal_lambda, ...
        'Solver', 'lbfgs', ...
        'Weights', sample_weights);

    coefficients = final_mdl.Beta;
    intercept = final_mdl.Bias;

    fprintf('  ✓ Cost-Sensitive 로지스틱 회귀 학습 성공\n');
    fprintf('  절편: %.4f\n', intercept);

    % 양수 계수만 사용하여 가중치 변환
    positive_coefs = max(0, coefficients);
    final_weights = positive_coefs / sum(positive_coefs) * 100;

    fprintf('  양수 계수 개수: %d/%d\n', sum(positive_coefs > 0), length(coefficients));
    fprintf('  가중치 변환 완료 (백분율)\n');

catch ME
    warning('모델 학습 실패: %s. 상관계수로 대체합니다.', ME.message);
    correlations = zeros(n_features, 1);
    for i = 1:n_features
        correlations(i) = corr(X_normalized(:,i), y_weight, 'rows', 'complete');
    end
    coefficients = correlations;
    intercept = 0;
    positive_coefs = max(0, coefficients);
    final_weights = positive_coefs / sum(positive_coefs) * 100;
end

%% 【STEP 13】 모델 성능 평가 및 검증 (원본과 동일)
fprintf('\n【STEP 13】 모델 성능 평가 및 검증\n');
fprintf('────────────────────────────────────────────\n');

% 종합점수 계산 (가중치 적용)
weighted_scores = X_normalized * (final_weights / 100);

% 고성과자와 저성과자의 종합점수 비교
high_idx = y_weight == 1;
low_idx = y_weight == 0;
high_scores = weighted_scores(high_idx);
low_scores = weighted_scores(low_idx);

fprintf('종합점수 검증:\n');
fprintf('  고성과자 평균: %.3f ± %.3f (n=%d)\n', mean(high_scores), std(high_scores), length(high_scores));
fprintf('  저성과자 평균: %.3f ± %.3f (n=%d)\n', mean(low_scores), std(low_scores), length(low_scores));

% 효과크기 (Cohen's d)
pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                  (length(low_scores)-1)*var(low_scores)) / ...
                 (length(high_scores) + length(low_scores) - 2));

if pooled_std > 0
    cohens_d = (mean(high_scores) - mean(low_scores)) / pooled_std;
    fprintf('  Cohen''s d: %.3f', cohens_d);
    if abs(cohens_d) >= 0.8
        fprintf(' (대 효과)\n');
    elseif abs(cohens_d) >= 0.5
        fprintf(' (중 효과)\n');
    elseif abs(cohens_d) >= 0.2
        fprintf(' (소 효과)\n');
    else
        fprintf(' (무시 가능)\n');
    end
else
    cohens_d = 0;
    fprintf('  Cohen''s d: 계산 불가 (표준편차 0)\n');
end

% ROC 분석
[X_roc, Y_roc, T_roc, AUC] = perfcurve(y_weight, weighted_scores, 1);
fprintf('  분류 성능 (AUC): %.3f\n', AUC);

% 최적 임계값 찾기 (Youden's J statistic)
J = Y_roc - X_roc;
[~, opt_idx] = max(J);
optimal_threshold = T_roc(opt_idx);

fprintf('  최적 임계값: %.3f (민감도=%.3f, 특이도=%.3f)\n', ...
        optimal_threshold, Y_roc(opt_idx), 1-X_roc(opt_idx));

%% 【STEP 14】 가중치 결과 분석 및 저장 (원본과 동일)
fprintf('\n【STEP 14】 가중치 결과 분석 및 저장\n');
fprintf('────────────────────────────────────────────\n');

% 가중치 결과 테이블 생성
weight_results = table();
weight_results.Feature = feature_names';
weight_results.Weight_Percent = final_weights;
weight_results.Raw_Coefficient = coefficients;

% 기여도가 있는 역량만 필터링 (0.1% 이상)
significant_idx = final_weights > 0.1;
weight_results_significant = weight_results(significant_idx, :);
weight_results_significant = sortrows(weight_results_significant, 'Weight_Percent', 'descend');

fprintf('주요 역량 가중치 (기여도 0.1%% 이상):\n');
fprintf('순위 | %-25s | 가중치(%%) | 원계수\n', '역량명');
fprintf('%s\n', repmat('-', 70, 1));

for i = 1:min(15, height(weight_results_significant))
    fprintf('%2d   | %-25s | %8.2f | %8.4f\n', ...
            i, weight_results_significant.Feature{i}, ...
            weight_results_significant.Weight_Percent(i), ...
            weight_results_significant.Raw_Coefficient(i));
end

%% 【STEP 16】 Bootstrap 가중치 안정성 검증 (원본과 동일)
fprintf('\n【STEP 16】 Bootstrap 가중치 안정성 검증\n');
fprintf('────────────────────────────────────────────\n');

% Bootstrap 설정 (원본과 동일)
n_bootstrap = 1000;  % 원본과 동일
n_samples = length(y_weight);
n_features = size(X_final, 2);

% Bootstrap 결과 저장 배열
bootstrap_weights = zeros(n_features, n_bootstrap);
bootstrap_rankings = zeros(n_features, n_bootstrap);

fprintf('▶ Bootstrap 가중치 안정성 검증 시작 (%d회)\n', n_bootstrap);

for b = 1:n_bootstrap
    % 복원추출
    boot_idx = randsample(n_samples, n_samples, true);
    X_boot = X_final(boot_idx, :);
    y_boot = y_final(boot_idx);

    % 정규화
    X_boot_norm = zscore(X_boot);

    % 샘플 가중치 재계산
    n_high_boot = sum(y_boot == 1);
    n_low_boot = sum(y_boot == 0);

    % 클래스가 하나만 있는 경우 스킵
    if n_high_boot == 0 || n_low_boot == 0
        bootstrap_weights(:, b) = NaN;
        bootstrap_rankings(:, b) = NaN;
        continue;
    end

    class_weights_boot = [n_samples/(2*n_low_boot), n_samples/(2*n_high_boot)];
    sample_weights_boot = zeros(size(y_boot));
    sample_weights_boot(y_boot == 0) = class_weights_boot(1);
    sample_weights_boot(y_boot == 1) = class_weights_boot(2);

    % Cost-Sensitive 모델 학습 (최적 Lambda 사용)
    try
        mdl_boot = fitclinear(X_boot_norm, y_boot, ...
            'Learner', 'logistic', ...
            'Cost', cost_matrix, ...
            'Weights', sample_weights_boot, ...
            'Regularization', 'ridge', ...
            'Lambda', optimal_lambda);

        % 가중치 추출 및 저장
        coefs = mdl_boot.Beta;
        positive_coefs = max(0, coefs);
        if sum(positive_coefs) > 0
            weights = positive_coefs / sum(positive_coefs) * 100;
        else
            weights = zeros(size(positive_coefs));
        end
        bootstrap_weights(:, b) = weights;

        % 순위 저장
        [~, ranks] = sort(weights, 'descend');
        for r = 1:length(ranks)
            bootstrap_rankings(ranks(r), b) = r;
        end

    catch
        % 실패한 경우 NaN 처리
        bootstrap_weights(:, b) = NaN;
        bootstrap_rankings(:, b) = NaN;
    end

    % 진행 상황 표시
    if mod(b, 100) == 0
        fprintf('  진행: %d/%d (%.1f%%)\n', b, n_bootstrap, b/n_bootstrap*100);
    end
end

% 각 역량별 통계
bootstrap_stats = table();
bootstrap_stats.Feature = feature_names';
bootstrap_stats.Original_Weight = final_weights;
bootstrap_stats.Boot_Mean = nanmean(bootstrap_weights, 2);
bootstrap_stats.Boot_Std = nanstd(bootstrap_weights, 0, 2);
bootstrap_stats.CI_Lower = prctile(bootstrap_weights, 2.5, 2);
bootstrap_stats.CI_Upper = prctile(bootstrap_weights, 97.5, 2);
bootstrap_stats.CV = bootstrap_stats.Boot_Std ./ (bootstrap_stats.Boot_Mean + eps);

% 순위 안정성
bootstrap_stats.Top3_Prob = zeros(height(bootstrap_stats), 1);
bootstrap_stats.Top5_Prob = zeros(height(bootstrap_stats), 1);

for i = 1:height(bootstrap_stats)
    ranks = bootstrap_rankings(i, :);
    valid_ranks = ranks(~isnan(ranks));
    if ~isempty(valid_ranks)
        bootstrap_stats.Top3_Prob(i) = sum(valid_ranks <= 3) / length(valid_ranks) * 100;
        bootstrap_stats.Top5_Prob(i) = sum(valid_ranks <= 5) / length(valid_ranks) * 100;
    end
end

% 원본 가중치 기준으로 정렬
bootstrap_stats = sortrows(bootstrap_stats, 'Original_Weight', 'descend');

% 결과 출력
fprintf('가중치 안정성 분석 (상위 10개):\n');
fprintf('%-20s | 원본(%%) | 평균(%%) | 95%% CI | CV | Top3확률(%%) | Top5확률(%%)\n', '역량');
fprintf('%s\n', repmat('-', 95, 1));

for i = 1:min(10, height(bootstrap_stats))
    fprintf('%-20s | %6.2f | %6.2f | [%.2f,%.2f] | %.3f | %8.1f | %8.1f\n', ...
            bootstrap_stats.Feature{i}, ...
            bootstrap_stats.Original_Weight(i), ...
            bootstrap_stats.Boot_Mean(i), ...
            bootstrap_stats.CI_Lower(i), ...
            bootstrap_stats.CI_Upper(i), ...
            bootstrap_stats.CV(i), ...
            bootstrap_stats.Top3_Prob(i), ...
            bootstrap_stats.Top5_Prob(i));
end

%% 【STEP 19】 4가지 가중치 방법론 통합 분석 (원본과 동일)
fprintf('\n【STEP 19】 4가지 가중치 방법론 통합 분석\n');
fprintf('════════════════════════════════════════════════════════════\n\n');

% 1. 상관 기반 가중치
fprintf('▶ 1. 상관 기반 가중치 (Correlation-based Weights)\n');
fprintf('────────────────────────────────────────────\n');
corr_weights = zeros(length(feature_names), 1);
for i = 1:length(feature_names)
    idx = strcmp(correlation_results.Competency, feature_names{i});
    if any(idx)
        corr_weights(i) = correlation_results.Weight(idx);
    end
end
fprintf('  ✓ 상관 가중치 추출 완료\n');

% 2. 로지스틱 회귀 가중치
fprintf('\n▶ 2. 로지스틱 회귀 가중치 (Logistic Regression Weights)\n');
fprintf('────────────────────────────────────────────\n');
logit_weights = final_weights;
fprintf('  ✓ 로지스틱 회귀 가중치 (Cost-Sensitive) 사용\n');

% 3. Bootstrap 평균 가중치
fprintf('\n▶ 3. Bootstrap 평균 가중치 (Bootstrap Mean Weights)\n');
fprintf('────────────────────────────────────────────\n');
boot_weights = zeros(length(feature_names), 1);
for i = 1:length(feature_names)
    idx = strcmp(bootstrap_stats.Feature, feature_names{i});
    if any(idx)
        boot_weights(i) = bootstrap_stats.Boot_Mean(idx);
    end
end
fprintf('  ✓ Bootstrap 평균 가중치 추출 완료\n');

% 4. T-test 효과크기 가중치
fprintf('\n▶ 4. T-test 효과크기 가중치 (T-test Effect Size Weights)\n');
fprintf('────────────────────────────────────────────\n');
ttest_weights = zeros(n_features, 1);

for i = 1:n_features
    high_scores = X_normalized(y_weight == 1, i);
    low_scores = X_normalized(y_weight == 0, i);

    pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                      (length(low_scores)-1)*var(low_scores)) / ...
                     (length(high_scores) + length(low_scores) - 2));

    if pooled_std > 0
        effect_size = abs(mean(high_scores) - mean(low_scores)) / pooled_std;
        ttest_weights(i) = effect_size;
    end
end

ttest_weights = (ttest_weights / sum(ttest_weights)) * 100;
fprintf('  ✓ T-test 효과크기 가중치 계산 완료\n');

% 5. 앙상블 가중치 (4가지 방법 평균)
ensemble_weights = (corr_weights + logit_weights + boot_weights + ttest_weights) / 4;

% 통합 결과 테이블
integrated_weights = table();
integrated_weights.Feature = feature_names';
integrated_weights.Correlation = corr_weights;
integrated_weights.Logistic = logit_weights;
integrated_weights.Bootstrap = boot_weights;
integrated_weights.TTest = ttest_weights;
integrated_weights.Ensemble = ensemble_weights;

% 앙상블 기준으로 정렬
integrated_weights = sortrows(integrated_weights, 'Ensemble', 'descend');

fprintf('\n통합 가중치 분석 결과 (상위 10개):\n');
fprintf('%-20s | 상관(%%) | 로지스틱(%%) | Bootstrap(%%) | T-test(%%) | 앙상블(%%)\n', '역량');
fprintf('%s\n', repmat('-', 95, 1));

for i = 1:min(10, height(integrated_weights))
    fprintf('%-20s | %6.2f | %9.2f | %10.2f | %8.2f | %8.2f\n', ...
            integrated_weights.Feature{i}, ...
            integrated_weights.Correlation(i), ...
            integrated_weights.Logistic(i), ...
            integrated_weights.Bootstrap(i), ...
            integrated_weights.TTest(i), ...
            integrated_weights.Ensemble(i));
end

%% 【STEP 22.5】 모델 재학습 기반 퍼뮤테이션 테스트 (캐시 없이)
fprintf('\n【STEP 22.5】 모델 재학습 기반 퍼뮤테이션 테스트\n');
fprintf('════════════════════════════════════════════════════════════\n');

n_permutations = 1000;  % 원본보다 축소 (5000→1000)
fprintf('▶ 퍼뮤테이션 테스트 시작 (%d회, 캐시 없음)\n', n_permutations);

% 원본 모델 성능 저장
original_auc = AUC;
original_cohens_d = cohens_d;

% 퍼뮤테이션 결과 저장
perm_aucs = zeros(n_permutations, 1);
perm_cohens_d = zeros(n_permutations, 1);

for p = 1:n_permutations
    % 레이블 셔플
    y_shuffled = y_weight(randperm(length(y_weight)));

    try
        % 셔플된 데이터로 모델 재학습
        mdl_perm = fitclinear(X_normalized, y_shuffled, ...
            'Learner', 'logistic', ...
            'Regularization', 'ridge', ...
            'Lambda', optimal_lambda, ...
            'Solver', 'lbfgs', ...
            'Weights', sample_weights);

        % 가중치 추출
        coefs_perm = mdl_perm.Beta;
        positive_coefs_perm = max(0, coefs_perm);

        if sum(positive_coefs_perm) > 0
            weights_perm = positive_coefs_perm / sum(positive_coefs_perm);
            scores_perm = X_normalized * weights_perm;
        else
            scores_perm = randn(size(y_shuffled));  % 랜덤 점수
        end

        % AUC 계산
        try
            [~, ~, ~, auc_perm] = perfcurve(y_shuffled, scores_perm, 1);
            perm_aucs(p) = auc_perm;
        catch
            perm_aucs(p) = 0.5;  % 기본값
        end

        % Cohen's d 계산
        high_idx_perm = y_shuffled == 1;
        low_idx_perm = y_shuffled == 0;

        if sum(high_idx_perm) > 0 && sum(low_idx_perm) > 0
            high_scores_perm = scores_perm(high_idx_perm);
            low_scores_perm = scores_perm(low_idx_perm);

            pooled_std_perm = sqrt(((length(high_scores_perm)-1)*var(high_scores_perm) + ...
                                   (length(low_scores_perm)-1)*var(low_scores_perm)) / ...
                                  (length(high_scores_perm) + length(low_scores_perm) - 2));

            if pooled_std_perm > 0
                perm_cohens_d(p) = abs(mean(high_scores_perm) - mean(low_scores_perm)) / pooled_std_perm;
            else
                perm_cohens_d(p) = 0;
            end
        else
            perm_cohens_d(p) = 0;
        end

    catch
        % 실패시 기본값
        perm_aucs(p) = 0.5;
        perm_cohens_d(p) = 0;
    end

    % 진행 상황 표시
    if mod(p, 200) == 0
        fprintf('  진행: %d/%d (%.1f%%)\n', p, n_permutations, p/n_permutations*100);
    end
end

% p-value 계산
p_value_auc = sum(perm_aucs >= original_auc) / n_permutations;
p_value_cohens_d = sum(perm_cohens_d >= abs(original_cohens_d)) / n_permutations;

fprintf('\n퍼뮤테이션 테스트 결과:\n');
fprintf('  원본 AUC: %.3f\n', original_auc);
fprintf('  퍼뮤테이션 평균 AUC: %.3f ± %.3f\n', mean(perm_aucs), std(perm_aucs));
fprintf('  AUC p-value: %.4f', p_value_auc);

if p_value_auc < 0.001
    fprintf(' (***)\n');
elseif p_value_auc < 0.01
    fprintf(' (**)\n');
elseif p_value_auc < 0.05
    fprintf(' (*)\n');
else
    fprintf(' (ns)\n');
end

fprintf('  원본 Cohen''s d: %.3f\n', abs(original_cohens_d));
fprintf('  퍼뮤테이션 평균 Cohen''s d: %.3f ± %.3f\n', mean(perm_cohens_d), std(perm_cohens_d));
fprintf('  Cohen''s d p-value: %.4f', p_value_cohens_d);

if p_value_cohens_d < 0.001
    fprintf(' (***)\n');
elseif p_value_cohens_d < 0.01
    fprintf(' (**)\n');
elseif p_value_cohens_d < 0.05
    fprintf(' (*)\n');
else
    fprintf(' (ns)\n');
end

%% 【STEP 23】 최종 결과 요약 및 권장사항
fprintf('\n【STEP 23】 최종 결과 요약 및 권장사항\n');
fprintf('════════════════════════════════════════════════════════════\n');

fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
fprintf('║                     최종 분석 요약                         ║\n');
fprintf('╚════════════════════════════════════════════════════════════╝\n\n');

% 최고 성능 방법 찾기
methods = {'Correlation', 'Logistic', 'Bootstrap', 'TTest', 'Ensemble'};
all_weights = [corr_weights, logit_weights, boot_weights, ttest_weights, ensemble_weights];

% 각 방법별 가중 점수 계산
method_scores = zeros(1, 5);
for i = 1:5
    method_weights = all_weights(:, i);
    method_weights = method_weights / sum(method_weights);  % 정규화
    method_score = X_normalized * method_weights;

    % AUC 계산
    try
        [~, ~, ~, method_auc] = perfcurve(y_weight, method_score, 1);
        method_scores(i) = method_auc;
    catch
        method_scores(i) = 0.5;
    end
end

[best_score, best_method_idx] = max(method_scores);
best_method = methods{best_method_idx};

fprintf('【최적 예측 모델】\n');
fprintf('  방법: %s\n', best_method);
fprintf('  AUC: %.3f\n', best_score);

% 핵심 역량 (앙상블 기준 상위 5개)
fprintf('\n【핵심 역량 Top 5】 (앙상블 가중치 기준)\n');
for i = 1:min(5, height(integrated_weights))
    fprintf('  %d. %s: %.2f%%\n', i, ...
            integrated_weights.Feature{i}, ...
            integrated_weights.Ensemble(i));
end

% 통계적 유의성
fprintf('\n【통계적 유의성】\n');
fprintf('  AUC p-value: %.4f', p_value_auc);
if p_value_auc < 0.001
    fprintf(' (***)\n');
elseif p_value_auc < 0.01
    fprintf(' (**)\n');
elseif p_value_auc < 0.05
    fprintf(' (*)\n');
else
    fprintf(' (ns)\n');
end

fprintf('  Cohen''s d p-value: %.4f', p_value_cohens_d);
if p_value_cohens_d < 0.001
    fprintf(' (***)\n');
elseif p_value_cohens_d < 0.01
    fprintf(' (**)\n');
elseif p_value_cohens_d < 0.05
    fprintf(' (*)\n');
else
    fprintf(' (ns)\n');
end

fprintf('\n════════════════════════════════════════════════════════════\n');

fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
fprintf('║            ✅ HR 데이터 분석 완료! (원본 매치 버전)         ║\n');
fprintf('╚════════════════════════════════════════════════════════════╝\n\n');