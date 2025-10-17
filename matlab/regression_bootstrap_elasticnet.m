%% ===============================================================
%  Enhanced Elastic Net Feature Importance Analysis
%  목적: 역량검사 데이터로 성과 예측 및 특성 중요도 분석
%  방법: Elastic Net + Bootstrap + Repeated CV + Enhanced Permutation
%  평가: Accuracy, Precision, Recall, F1-Score, MCC
%  개선: 부트스트랩 리샘플링, 반복 교차검증, 퍼뮤테이션 반복 증가
clear; clc; close all;
main()
%% ===============================================================

function main()
% 메인 실행 함수
clear; clc; close all;
rng(42);  % 재현 가능한 결과를 위한 시드 설정

fprintf('=== Enhanced Elastic Net Feature Importance Analysis ===\n\n');

%% 분석 설정
config = struct();
config.n_bootstrap = 50;        % 부트스트랩 반복 횟수
config.n_cv_repeats = 3;        % 교차검증 반복 횟수
config.n_permutations = 30;     % 퍼뮤테이션 반복 횟수 (15→30)
config.cv_folds = 5;            % 교차검증 폴드 수
config.test_ratio = 0.25;       % 테스트 데이터 비율

fprintf('분석 설정:\n');
fprintf('- 부트스트랩 반복: %d회\n', config.n_bootstrap);
fprintf('- 교차검증 반복: %d회 × %d-fold\n', config.n_cv_repeats, config.cv_folds);
fprintf('- 퍼뮤테이션 반복: %d회\n', config.n_permutations);
fprintf('\n');

%% 1. 데이터 생성 및 전처리
fprintf('1. 데이터 생성 및 전처리...\n');
[X, Y, featureNames] = generateSyntheticData();
[X_train, X_test, Y_train, Y_test] = splitData(X, Y, config.test_ratio);

%% 2. Enhanced Elastic Net 모델 훈련 (Repeated CV)
fprintf('2. Enhanced Elastic Net 모델 훈련 (Repeated CV)...\n');
[best_model, best_coefficients, cv_results] = trainEnhancedElasticNet(X_train, Y_train, config);

%% 3. 부트스트랩 안정성 분석
fprintf('3. 부트스트랩 안정성 분석...\n');
bootstrap_results = performBootstrapAnalysis(X_train, Y_train, config);

%% 4. 모델 성능 평가
fprintf('4. 모델 성능 평가...\n');
performance = evaluateModelPerformance(best_model, X_test, Y_test);
displayPerformance(performance);

%% 5. Enhanced 특성 중요도 분석
fprintf('5. Enhanced 특성 중요도 분석...\n');
importance = analyzeEnhancedFeatureImportance(best_model, best_coefficients, ...
    X_test, Y_test, featureNames, ...
    bootstrap_results, config);

%% 6. 종합 결과 시각화
fprintf('6. 종합 결과 시각화...\n');
visualizeEnhancedResults(importance, cv_results, bootstrap_results);

fprintf('\n=== Enhanced 분석 완료 ===\n');
end

%% ==================== 유틸리티 함수 ====================
function X_std = simpleStandardize(X)
% 간단한 Z-score 표준화 (평균 0, 표준편차 1)
X_std = (X - mean(X)) ./ (std(X) + eps);
end

function [cm, labels] = safeConfusionMatrix(Y_true, Y_pred)
% 호환성을 위한 안전한 혼동행렬 생성
try
    cm = confusionmat(Y_true, Y_pred);
    labels = categories(Y_true);
catch
    categories_true = categories(Y_true);
    categories_pred = categories(Y_pred);
    all_categories = unique([categories_true; categories_pred]);

    [~, idx_true] = ismember(Y_true, all_categories);
    [~, idx_pred] = ismember(Y_pred, all_categories);

    n_classes = length(all_categories);
    cm = zeros(n_classes, n_classes);

    for i = 1:length(idx_true)
        cm(idx_true(i), idx_pred(i)) = cm(idx_true(i), idx_pred(i)) + 1;
    end

    labels = all_categories;
end
end

%% ==================== 데이터 생성 및 전처리 ====================
function [X, Y, featureNames] = generateSyntheticData()
% 합성 데이터 생성 (실제 역량검사 데이터 구조 모방)

n = 100;  % 샘플 수
p = 25;   % 특성 수

% 특성 이름 생성
featureNames = cell(p, 1);
for i = 1:10
    featureNames{i} = sprintf('역량_%d', i);
end
for i = 11:18
    featureNames{i} = sprintf('기술_%d', i-10);
end
for i = 19:25
    featureNames{i} = sprintf('성격_%d', i-18);
end

% 정규분포 특성 생성
X = randn(n, p);

% 실제 중요한 특성들 (처음 8개)
true_weights = [2.5, -1.8, 1.5, -2.0, 1.2, -1.0, 0.8, -1.5, zeros(1, p-8)];

% 비선형 관계 추가
signal = X * true_weights' + ...
    0.5 * sin(X(:,1)) + ...
    0.3 * X(:,3).^2 + ...
    -0.4 * X(:,2) .* X(:,4);

% 노이즈 추가
signal = signal + 0.3 * randn(n, 1);

% 이진 분류 레이블 생성 (약간의 불균형)
probabilities = 1 ./ (1 + exp(-signal));
Y_binary = probabilities > 0.6;  % 약 40% positive class

% Categorical 변환
Y = categorical(Y_binary, [false, true], {'Low_Performance', 'High_Performance'});

% 간단한 표준화
X = simpleStandardize(X);

fprintf('   - 샘플 수: %d, 특성 수: %d\n', n, p);

cats = categories(Y);
counts = countcats(Y);

fprintf('   - 클래스 분포:\n');
for k = 1:numel(cats)
    fprintf('     %s = %d\n', cats{k}, counts(k));
end

end

function [X_train, X_test, Y_train, Y_test] = splitData(X, Y, test_ratio)
% 계층화 데이터 분할
cv_partition = cvpartition(Y, 'HoldOut', test_ratio);

X_train = X(training(cv_partition), :);
X_test = X(test(cv_partition), :);
Y_train = Y(training(cv_partition));
Y_test = Y(test(cv_partition));

fprintf('   - 훈련 데이터: %d개, 테스트 데이터: %d개\n', ...
    length(Y_train), length(Y_test));
end

%% ==================== Enhanced Elastic Net 모델 ====================
function [best_model, best_coefficients, cv_results] = trainEnhancedElasticNet(X_train, Y_train, config)
% Repeated Cross-Validation을 통한 Enhanced Elastic Net 훈련

y_numeric = double(Y_train == 'High_Performance');

% 반복 교차검증 결과 저장
cv_results = struct();
cv_results.scores = [];
cv_results.lambda_values = [];
cv_results.coefficients = [];

fprintf('   반복 교차검증 진행...\n');

all_scores = [];
all_lambdas = [];
all_coeffs = {};

for repeat = 1:config.n_cv_repeats
    fprintf('   - 반복 %d/%d 진행 중...\n', repeat, config.n_cv_repeats);

    try
        [B, FitInfo] = lasso(X_train, y_numeric, ...
            'CV', config.cv_folds, ...
            'Alpha', 0.5, ...
            'Standardize', false, ...
            'NumLambda', 50);

        % 결과 저장
        all_scores = [all_scores; FitInfo.MSE];
        all_lambdas = [all_lambdas; FitInfo.Lambda];
        all_coeffs{repeat} = B;

        % 첫 번째 반복에서 기본 정보 저장
        if repeat == 1
            cv_results.lambda_values = FitInfo.Lambda;
            cv_results.intercepts = FitInfo.Intercept;
        end

    catch ME
        fprintf('   경고: lasso 함수 오류 (반복 %d). 릿지 회귀로 대체.\n', repeat);
        fprintf('   오류: %s\n', ME.message);

        % 릿지 회귀로 대체
        lambda = 0.1;
        B_ridge = (X_train' * X_train + lambda * eye(size(X_train, 2))) \ (X_train' * y_numeric);
        intercept = mean(y_numeric) - mean(X_train) * B_ridge;

        if repeat == 1
            cv_results.lambda_values = lambda;
            cv_results.intercepts = intercept;
            best_coefficients = [intercept; B_ridge];
            best_model = @(X) [ones(size(X,1), 1), X] * best_coefficients;
        end
        break;
    end
end

if ~isempty(all_scores)
    % 평균 CV 점수 계산
    mean_scores = mean(all_scores, 1);
    std_scores = std(all_scores, [], 1);

    % 1-SE 규칙으로 최적 lambda 선택
    [min_mse, min_idx] = min(mean_scores);
    se_threshold = min_mse + std_scores(min_idx);
    lambda_1se_idx = find(mean_scores <= se_threshold, 1, 'first');

    % 최종 모델 구성
    final_coefficients = mean(cat(3, all_coeffs{:}), 3);  % 계수 평균화
    best_coefficients = [cv_results.intercepts(lambda_1se_idx); final_coefficients(:, lambda_1se_idx)];
    best_model = @(X) [ones(size(X,1), 1), X] * best_coefficients;

    cv_results.scores = all_scores;
    cv_results.mean_scores = mean_scores;
    cv_results.std_scores = std_scores;
    cv_results.best_lambda = cv_results.lambda_values(lambda_1se_idx);

    fprintf('   - 최적 lambda: %.4f (평균 %d회 반복)\n', cv_results.best_lambda, config.n_cv_repeats);
end

% 선택된 특성 수
n_selected = sum(abs(best_coefficients(2:end)) > 1e-6);
fprintf('   - 유의미한 특성 수: %d/%d\n', n_selected, length(best_coefficients)-1);
end

%% ==================== 부트스트랩 분석 ====================
function bootstrap_results = performBootstrapAnalysis(X_train, Y_train, config)
% 부트스트랩 리샘플링을 통한 모델 안정성 분석

fprintf('   부트스트랩 분석 진행 (%d회)...\n', config.n_bootstrap);

n_samples = length(Y_train);
n_features = size(X_train, 2);
y_numeric = double(Y_train == 'High_Performance');

% 부트스트랩 결과 저장
bootstrap_coeffs = zeros(n_features + 1, config.n_bootstrap);  % +1 for intercept
bootstrap_selected = zeros(n_features, config.n_bootstrap);
bootstrap_performance = zeros(config.n_bootstrap, 1);

for boot = 1:config.n_bootstrap
    if mod(boot, 10) == 0
        fprintf('   - 부트스트랩 %d/%d 완료\n', boot, config.n_bootstrap);
    end

    % 부트스트랩 샘플링 (복원추출)
    boot_indices = randi(n_samples, n_samples, 1);
    X_boot = X_train(boot_indices, :);
    y_boot = y_numeric(boot_indices);

    % 부트스트랩 샘플로 모델 훈련
    try
        [B, FitInfo] = lasso(X_boot, y_boot, ...
            'CV', 3, ...  % 빠른 CV
            'Alpha', 0.5, ...
            'Standardize', false, ...
            'NumLambda', 30);

        lambda_idx = FitInfo.Index1SE;
        bootstrap_coeffs(:, boot) = [FitInfo.Intercept(lambda_idx); B(:, lambda_idx)];
        bootstrap_selected(:, boot) = abs(B(:, lambda_idx)) > 1e-6;

        % OOB (Out-of-Bag) 성능 평가
        oob_indices = setdiff(1:n_samples, unique(boot_indices));
        if ~isempty(oob_indices)
            model_boot = @(X) [ones(size(X,1), 1), X] * bootstrap_coeffs(:, boot);
            y_pred_oob = model_boot(X_train(oob_indices, :)) > 0.5;
            bootstrap_performance(boot) = mean(y_pred_oob == y_numeric(oob_indices));
        end

    catch
        % 오류 발생시 릿지 회귀
        lambda = 0.1;
        B_ridge = (X_boot' * X_boot + lambda * eye(size(X_boot, 2))) \ (X_boot' * y_boot);
        intercept = mean(y_boot) - mean(X_boot) * B_ridge;
        bootstrap_coeffs(:, boot) = [intercept; B_ridge];
        bootstrap_selected(:, boot) = abs(B_ridge) > 1e-6;
    end
end

% 부트스트랩 결과 요약
bootstrap_results = struct();
bootstrap_results.coefficients = bootstrap_coeffs;
bootstrap_results.selected_features = bootstrap_selected;
bootstrap_results.performance = bootstrap_performance;
bootstrap_results.mean_coeffs = mean(bootstrap_coeffs, 2);
bootstrap_results.std_coeffs = std(bootstrap_coeffs, [], 2);
bootstrap_results.selection_frequency = mean(bootstrap_selected, 2);
bootstrap_results.mean_performance = mean(bootstrap_performance(bootstrap_performance > 0));

fprintf('   - 평균 OOB 성능: %.3f\n', bootstrap_results.mean_performance);
fprintf('   - 안정적 특성 (선택 빈도 > 50%%): %d개\n', ...
    sum(bootstrap_results.selection_frequency > 0.5));
end

%% ==================== 모델 성능 평가 ====================
function performance = evaluateModelPerformance(model, X_test, Y_test)
% 모델 성능 종합 평가

y_scores = model(X_test);
y_pred_binary = y_scores > 0.5;
Y_pred = categorical(y_pred_binary, [false, true], ...
    {'Low_Performance', 'High_Performance'});

[confusion_matrix, ~] = safeConfusionMatrix(Y_test, Y_pred);

performance = struct();
performance.confusion_matrix = confusion_matrix;
performance.accuracy = calculateAccuracy(Y_test, Y_pred);
performance.precision = calculatePrecision(confusion_matrix);
performance.recall = calculateRecall(confusion_matrix);
performance.f1_score = calculateF1Score(performance.precision, performance.recall);
performance.mcc = calculateMCC(confusion_matrix);
performance.y_scores = y_scores;
performance.y_true = Y_test;
performance.y_pred = Y_pred;
end

function accuracy = calculateAccuracy(Y_true, Y_pred)
accuracy = sum(Y_true == Y_pred) / length(Y_true);
end

function precision = calculatePrecision(cm)
if size(cm, 1) == 2
    precision = cm(2,2) / (cm(2,2) + cm(1,2) + eps);
else
    precision = NaN;
end
end

function recall = calculateRecall(cm)
if size(cm, 1) == 2
    recall = cm(2,2) / (cm(2,2) + cm(2,1) + eps);
else
    recall = NaN;
end
end

function f1 = calculateF1Score(precision, recall)
f1 = 2 * precision * recall / (precision + recall + eps);
end

function mcc = calculateMCC(cm)
if isequal(size(cm), [2, 2])
    TP = cm(2,2); TN = cm(1,1); FP = cm(1,2); FN = cm(2,1);
    numerator = (TP * TN) - (FP * FN);
    denominator = sqrt((TP + FP) * (TP + FN) * (TN + FP) * (TN + FN));
    mcc = numerator / max(denominator, eps);
else
    mcc = NaN;
end
end

function displayPerformance(performance)
fprintf('   [모델 성능 결과]\n');
fprintf('   - Accuracy:  %.3f\n', performance.accuracy);
fprintf('   - Precision: %.3f\n', performance.precision);
fprintf('   - Recall:    %.3f\n', performance.recall);
fprintf('   - F1-Score:  %.3f\n', performance.f1_score);
fprintf('   - MCC:       %.3f\n', performance.mcc);
fprintf('\n   혼동 행렬:\n');
disp(performance.confusion_matrix);
end

%% ==================== Enhanced 특성 중요도 분석 ====================
function importance = analyzeEnhancedFeatureImportance(model, coefficients, X_test, Y_test, featureNames, bootstrap_results, config)
% Enhanced 특성 중요도 계산 (부트스트랩 + 강화된 Permutation)

% 1. 계수 기반 중요도
coef_importance = calculateCoefficientImportance(coefficients);

% 2. 강화된 Permutation Importance (반복 증가)
fprintf('   Enhanced Permutation Importance 계산 중 (%d회 반복)...\n', config.n_permutations);
perm_importance = calculateEnhancedPermutationImportance(model, X_test, Y_test, config.n_permutations);

% 3. 부트스트랩 기반 안정성 점수
stability_scores = bootstrap_results.selection_frequency;

% 4. 종합 중요도 점수 (가중 평균)
composite_scores = 0.4 * coef_importance + 0.4 * perm_importance + 0.2 * stability_scores;

% 결과 구조체 생성
importance = struct();
importance.Feature = featureNames;
importance.Coefficient = coefficients(2:end);
importance.Coef_Importance = coef_importance;
importance.Perm_Importance = perm_importance;
importance.Stability_Score = stability_scores;
importance.Composite_Score = composite_scores;
importance.Bootstrap_Mean = bootstrap_results.mean_coeffs(2:end);
importance.Bootstrap_Std = bootstrap_results.std_coeffs(2:end);

% 종합 점수 기준으로 정렬
[~, sort_idx] = sort(composite_scores, 'descend');

importance.Feature = importance.Feature(sort_idx);
importance.Coefficient = importance.Coefficient(sort_idx);
importance.Coef_Importance = importance.Coef_Importance(sort_idx);
importance.Perm_Importance = importance.Perm_Importance(sort_idx);
importance.Stability_Score = importance.Stability_Score(sort_idx);
importance.Composite_Score = importance.Composite_Score(sort_idx);
importance.Bootstrap_Mean = importance.Bootstrap_Mean(sort_idx);
importance.Bootstrap_Std = importance.Bootstrap_Std(sort_idx);

% 상위 특성 표시
fprintf('   [Enhanced 특성 중요도 순위 (Top 10)]\n');
fprintf('   %-12s %8s %8s %8s %8s %8s\n', 'Feature', 'Coef', 'Perm', 'Stab', 'Comp', 'Boot_CI');
fprintf('   %s\n', repmat('-', 1, 70));

top_n = min(10, length(featureNames));
for i = 1:top_n
    ci_lower = importance.Bootstrap_Mean(i) - 1.96 * importance.Bootstrap_Std(i);
    ci_upper = importance.Bootstrap_Mean(i) + 1.96 * importance.Bootstrap_Std(i);

    fprintf('   %-12s %8.3f %8.3f %8.3f %8.3f [%5.2f,%5.2f]\n', ...
        importance.Feature{i}, ...
        importance.Coef_Importance(i), ...
        importance.Perm_Importance(i), ...
        importance.Stability_Score(i), ...
        importance.Composite_Score(i), ...
        ci_lower, ci_upper);
end
end

function coef_importance = calculateCoefficientImportance(coefficients)
abs_coef = abs(coefficients(2:end));
coef_importance = abs_coef / (sum(abs_coef) + eps);
end

function perm_importance = calculateEnhancedPermutationImportance(model, X_test, Y_test, n_permutations)
% 강화된 Permutation Importance 계산 (반복 증가)

n_features = size(X_test, 2);
perm_importance = zeros(n_features, 1);

baseline_scores = model(X_test);
baseline_predictions = categorical(baseline_scores > 0.5, [false, true], ...
    {'Low_Performance', 'High_Performance'});
baseline_accuracy = sum(baseline_predictions == Y_test) / length(Y_test);

for feature_idx = 1:n_features
    if mod(feature_idx, 5) == 0
        fprintf('     특성 %d/%d 처리 중...\n', feature_idx, n_features);
    end

    accuracy_drops = zeros(n_permutations, 1);

    for perm = 1:n_permutations
        X_permuted = X_test;
        perm_indices = randperm(size(X_test, 1));
        X_permuted(:, feature_idx) = X_test(perm_indices, feature_idx);

        permuted_scores = model(X_permuted);
        permuted_predictions = categorical(permuted_scores > 0.5, [false, true], ...
            {'Low_Performance', 'High_Performance'});
        permuted_accuracy = sum(permuted_predictions == Y_test) / length(Y_test);

        accuracy_drops(perm) = baseline_accuracy - permuted_accuracy;
    end

    perm_importance(feature_idx) = mean(accuracy_drops);
end

% 정규화
perm_importance = max(perm_importance, 0);
perm_importance = perm_importance / (sum(perm_importance) + eps);
end

%% ==================== Enhanced 결과 시각화 ====================
function visualizeEnhancedResults(importance, cv_results, bootstrap_results)
% Enhanced 분석 결과 종합 시각화

top_n = min(10, length(importance.Feature));
top_features = importance.Feature(1:top_n);

figure('Name', 'Enhanced Feature Importance Analysis', 'Position', [50, 50, 1600, 800]);

% 1. 다중 중요도 비교
subplot(2, 3, 1);
data_matrix = [importance.Coef_Importance(1:top_n), ...
    importance.Perm_Importance(1:top_n), ...
    importance.Stability_Score(1:top_n), ...
    importance.Composite_Score(1:top_n)];

barh(1:top_n, data_matrix);
set(gca, 'YTick', 1:top_n, 'YTickLabel', top_features);
title('Multiple Importance Scores', 'FontSize', 11, 'FontWeight', 'bold');
xlabel('Normalized Importance Score');
legend({'Coefficient', 'Permutation', 'Stability', 'Composite'}, 'Location', 'best');
grid on;
set(gca, 'FontSize', 9);

% 2. 부트스트랩 계수 분포 (상위 5개 특성)
subplot(2, 3, 2);
top_5_idx = 1:min(5, top_n);
boot_data = bootstrap_results.coefficients(2:end, :);  % 절편 제외

% 원래 순서로 되돌리기 위한 인덱스 찾기
[~, orig_idx] = sort(importance.Composite_Score, 'descend');
reverse_idx = zeros(size(orig_idx));
reverse_idx(orig_idx) = 1:length(orig_idx);

selected_boot_data = boot_data(reverse_idx(top_5_idx), :);

boxplot(selected_boot_data', 'Labels', top_features(top_5_idx));
title('Bootstrap Coefficient Distribution', 'FontSize', 11, 'FontWeight', 'bold');
ylabel('Coefficient Value');
grid on;
set(gca, 'FontSize', 9);
xtickangle(45);

% 3. 특성 선택 안정성
subplot(2, 3, 3);
stability_freq = importance.Stability_Score(1:top_n);
barh(1:top_n, stability_freq, 'FaceColor', [0.8, 0.2, 0.4]);
set(gca, 'YTick', 1:top_n, 'YTickLabel', top_features);
title('Feature Selection Stability', 'FontSize', 11, 'FontWeight', 'bold');
xlabel('Selection Frequency');
xlim([0, 1]);

% 50% 선 추가
hold on;
plot([0.5, 0.5], [0.5, top_n+0.5], 'r--', 'LineWidth', 2);
text(0.52, top_n*0.8, '50% threshold', 'Color', 'red', 'FontWeight', 'bold');
hold off;
grid on;
set(gca, 'FontSize', 9);

% 4. CV 성능 곡선 (available한 경우)
subplot(2, 3, 4);
if isfield(cv_results, 'mean_scores') && ~isempty(cv_results.mean_scores)
    errorbar(1:length(cv_results.mean_scores), cv_results.mean_scores, cv_results.std_scores);
    title('Cross-Validation Performance', 'FontSize', 11, 'FontWeight', 'bold');
    xlabel('Lambda Index');
    ylabel('CV MSE');
    grid on;

    % 최적 lambda 표시
    if isfield(cv_results, 'best_lambda')
        best_idx = find(cv_results.lambda_values == cv_results.best_lambda, 1);
        if ~isempty(best_idx)
            hold on;
            plot(best_idx, cv_results.mean_scores(best_idx), 'ro', 'MarkerSize', 8, 'MarkerFaceColor', 'red');
            text(best_idx, cv_results.mean_scores(best_idx), sprintf(' λ=%.3f', cv_results.best_lambda), ...
                'FontSize', 9, 'Color', 'red', 'FontWeight', 'bold');
            hold off;
        end
    end
else
    text(0.5, 0.5, 'CV 결과 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('Cross-Validation Performance', 'FontSize', 11, 'FontWeight', 'bold');
end
set(gca, 'FontSize', 9);

% 5. 부트스트랩 OOB 성능 히스토그램
subplot(2, 3, 5);
valid_performance = bootstrap_results.performance(bootstrap_results.performance > 0);
if ~isempty(valid_performance)
    histogram(valid_performance, 15, 'FaceColor', [0.3, 0.7, 0.3], 'EdgeColor', 'black');
    title('Bootstrap OOB Performance', 'FontSize', 11, 'FontWeight', 'bold');
    xlabel('Accuracy');
    ylabel('Frequency');

    % 평균 성능 선 표시
    mean_perf = mean(valid_performance);
    hold on;
    line([mean_perf, mean_perf], ylim, 'Color', 'red', 'LineWidth', 2, 'LineStyle', '--');
    text(mean_perf, max(ylim)*0.8, sprintf('Mean=%.3f', mean_perf), ...
        'HorizontalAlignment', 'center', 'Color', 'red', 'FontWeight', 'bold');
    hold off;
else
    text(0.5, 0.5, 'OOB 성능 데이터 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('Bootstrap OOB Performance', 'FontSize', 11, 'FontWeight', 'bold');
end
grid on;
set(gca, 'FontSize', 9);

% 6. 중요도 방법 간 상관관계 매트릭스
subplot(2, 3, 6);
correlation_data = [importance.Coef_Importance, importance.Perm_Importance, ...
    importance.Stability_Score, importance.Composite_Score];

try
    corr_matrix = corr(correlation_data);
    imagesc(corr_matrix);
    colorbar;
    colormap('RdBu');
    caxis([-1, 1]);

    % 상관계수 텍스트 표시
    [rows, cols] = size(corr_matrix);

    for i = 1:rows
        for j = 1:cols
            val = corr_matrix(i,j);
            % 조건에 따라 색상 선택
            if abs(val) > 0.5
                colorSpec = 'white';
            else
                colorSpec = 'black';
            end
            % 텍스트 객체 생성
            text(j, i, sprintf('%.2f', val), ...
                'HorizontalAlignment', 'center', ...
                'FontWeight', 'bold', ...
                'Color', colorSpec);
        end
    end

    set(gca, 'XTick', 1:4, 'XTickLabel', {'Coef', 'Perm', 'Stab', 'Comp'});
    set(gca, 'YTick', 1:4, 'YTickLabel', {'Coef', 'Perm', 'Stab', 'Comp'});
    title('Importance Methods Correlation', 'FontSize', 11, 'FontWeight', 'bold');
catch
    text(0.5, 0.5, '상관관계 계산 오류', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('Importance Methods Correlation', 'FontSize', 11, 'FontWeight', 'bold');
end
set(gca, 'FontSize', 9);

% 전체 제목
sgtitle('Enhanced Elastic Net Feature Importance Analysis Results', ...
    'FontSize', 14, 'FontWeight', 'bold');

% 추가 분석 결과 출력
fprintf('\n=== Enhanced 분석 요약 ===\n');
fprintf('가장 안정적인 특성 (상위 5개):\n');
for i = 1:min(5, length(importance.Feature))
    fprintf('  %d. %s (종합점수: %.3f, 안정성: %.1f%%)\n', ...
        i, importance.Feature{i}, importance.Composite_Score(i), ...
        importance.Stability_Score(i)*100);
end

fprintf('\n높은 안정성 특성 (선택빈도 > 70%%):\n');
high_stability_idx = find(importance.Stability_Score > 0.7);
if ~isempty(high_stability_idx)
    for idx = high_stability_idx'
        feature_pos = find(strcmp(importance.Feature, importance.Feature{idx}));
        fprintf('  - %s (선택빈도: %.1f%%, 순위: %d)\n', ...
            importance.Feature{idx}, importance.Stability_Score(idx)*100, feature_pos);
    end
else
    fprintf('  (없음 - 모든 특성이 70%% 미만)\n');
end

fprintf('\n부트스트랩 OOB 성능 요약:\n');
if ~isempty(valid_performance)
    fprintf('  - 평균: %.3f ± %.3f\n', mean(valid_performance), std(valid_performance));
    fprintf('  - 범위: [%.3f, %.3f]\n', min(valid_performance), max(valid_performance));
    fprintf('  - 신뢰구간 (95%%): [%.3f, %.3f]\n', ...
        prctile(valid_performance, 2.5), prctile(valid_performance, 97.5));
else
    fprintf('  - 데이터 없음\n');
end
end

%% 메인 함수 실행
if ~exist('OCTAVE_VERSION', 'builtin')
    % MATLAB에서만 실행
    main();
else
    fprintf('이 코드는 MATLAB에서 실행하도록 최적화되었습니다.\n');
end