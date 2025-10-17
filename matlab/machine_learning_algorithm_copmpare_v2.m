%% 포괄적인 MATLAB 머신러닝 모델 성능 비교 분석 (확장판: ROC-AUC/PR-AUC/MCC 포함)
% Comprehensive Machine Learning Model Performance Comparison
% 작성일: 2024 → 확장: 2025
% 목적: 다양한 앙상블 및 고급 머신러닝 모델의 체계적 비교 + 고급 지표/시각화

clear; clc; close all;
warning('off', 'all');

%% 1. 설정 및 초기화
fprintf('=====================================\n');
fprintf('   포괄적 머신러닝 모델 비교 분석 (확장판)\n');
fprintf('=====================================\n\n');

% 분석 설정
CONFIG = struct();
CONFIG.problem_type = 'classification';     % 'classification' or 'regression'
CONFIG.n_samples = 2000;                    % 샘플 수
CONFIG.n_features = 25;                     % 특성 수
CONFIG.test_ratio = 0.2;                    % 테스트 세트 비율
CONFIG.cv_folds = 5;                        % 교차검증 폴드 수
CONFIG.random_seed = 42;                    % 랜덤 시드
CONFIG.create_imbalance = true;             % 클래스 불균형 생성 (분류일 때만)
CONFIG.add_outliers = true;                 % 이상값 추가
CONFIG.outlier_ratio = 0.05;                % 이상값 비율
CONFIG.calculate_importance = true;         % 특성 중요도 계산
CONFIG.plot_results = true;                 % 결과 시각화

% 랜덤 시드 설정
rng(CONFIG.random_seed);

%% 2. 복잡한 데이터 생성
fprintf('[ 복잡한 데이터 생성 중... ]\n');

% 2.1 기본 특성 생성
% 선형 관계 특성 (5개)
X_linear = randn(CONFIG.n_samples, 5);

% 비선형 관계 특성 (5개)
X_nonlinear = zeros(CONFIG.n_samples, 5);
X_nonlinear(:,1) = sin(2*pi*rand(CONFIG.n_samples, 1));
X_nonlinear(:,2) = cos(3*rand(CONFIG.n_samples, 1));
X_nonlinear(:,3) = log(abs(randn(CONFIG.n_samples, 1)) + 1);
X_nonlinear(:,4) = randn(CONFIG.n_samples, 1).^2;
X_nonlinear(:,5) = tanh(randn(CONFIG.n_samples, 1));

% 상호작용 특성 (5개)
X_interaction = zeros(CONFIG.n_samples, 5);
X_interaction(:,1) = X_linear(:,1) .* X_linear(:,2);
X_interaction(:,2) = X_linear(:,3) .* X_nonlinear(:,1);
X_interaction(:,3) = X_nonlinear(:,2) .* X_nonlinear(:,3);
X_interaction(:,4) = abs(X_linear(:,1) - X_linear(:,2));
X_interaction(:,5) = (X_linear(:,4) + X_nonlinear(:,4)) / 2;

% 범주형 변수 시뮬레이션 (3개)
X_categorical = zeros(CONFIG.n_samples, 3);
X_categorical(:,1) = randi([1, 4], CONFIG.n_samples, 1) + 0.1*randn(CONFIG.n_samples, 1);
X_categorical(:,2) = randi([1, 3], CONFIG.n_samples, 1) + 0.1*randn(CONFIG.n_samples, 1);
X_categorical(:,3) = randi([1, 5], CONFIG.n_samples, 1) + 0.1*randn(CONFIG.n_samples, 1);

% 노이즈 특성 (나머지)
n_noise = CONFIG.n_features - 18;
X_noise = randn(CONFIG.n_samples, n_noise);

% 모든 특성 결합
X = [X_linear, X_nonlinear, X_interaction, X_categorical, X_noise];

% 2.2 특성 간 상관관계 추가
correlation_strength = 0.7;
for i = 1:3
    idx1 = randi(CONFIG.n_features);
    idx2 = randi(CONFIG.n_features);
    if idx1 ~= idx2
        X(:, idx2) = correlation_strength * X(:, idx1) + (1-correlation_strength) * X(:, idx2);
    end
end

% 2.3 이상값 추가
if CONFIG.add_outliers
    n_outliers = floor(CONFIG.n_samples * CONFIG.outlier_ratio);
    outlier_indices = randperm(CONFIG.n_samples, n_outliers);
    for i = 1:n_outliers
        feature_idx = randi(CONFIG.n_features);
        X(outlier_indices(i), feature_idx) = X(outlier_indices(i), feature_idx) + 5*randn();
    end
    fprintf('  - 이상값 추가: %d개 (%.1f%%)\n', n_outliers, CONFIG.outlier_ratio * 100);
end

% 2.4 타겟 변수 생성
if strcmp(CONFIG.problem_type, 'classification')
    % 이진 분류
    linear_score = X(:,1) - 0.5*X(:,2) + 0.3*X(:,3);
    nonlinear_score = 0.5*sin(X(:,6)) + 0.3*X(:,7).^2;
    interaction_score = 0.4*X(:,11).*X(:,12);
    
    decision_score = linear_score + nonlinear_score + interaction_score;
    probabilities = 1 ./ (1 + exp(-decision_score));
    
    % 노이즈 추가
    probabilities = probabilities + 0.1 * randn(CONFIG.n_samples, 1);
    probabilities = max(0, min(1, probabilities));
    
    Y_binary = probabilities > 0.5;
    Y = categorical(Y_binary, [false, true], {'Class_0', 'Class_1'});
    
    % 클래스 불균형 생성
    if CONFIG.create_imbalance
        minority_indices = find(Y == 'Class_1');
        keep_ratio = 0.3;  % 30%만 유지
        n_keep = floor(length(minority_indices) * keep_ratio);
        remove_indices = minority_indices(n_keep+1:end);
        
        keep_indices = setdiff(1:CONFIG.n_samples, remove_indices);
        X = X(keep_indices, :);
        Y = Y(keep_indices);
        
        fprintf('  - 클래스 불균형 생성: Class_1 비율 = %.1f%%\n', ...
            sum(Y == 'Class_1') / length(Y) * 100);
    end
else
    % 회귀 문제
    true_weights = randn(CONFIG.n_features, 1) * 2;
    Y = X * true_weights;
    Y = Y + 0.5 * X(:,1).^2 + 0.3 * sin(X(:,2));
    Y = Y + randn(size(Y)) * std(Y) * 0.1;  % 10% 노이즈
end

% 특성 이름/타입
feature_names = cell(1, CONFIG.n_features);
feature_types = cell(1, CONFIG.n_features);
for i = 1:5,  feature_names{i} = sprintf('Linear_%d', i);      feature_types{i} = 'linear';      end
for i = 6:10, feature_names{i} = sprintf('Nonlinear_%d', i-5); feature_types{i} = 'nonlinear';   end
for i = 11:15,feature_names{i} = sprintf('Interaction_%d', i-10); feature_types{i} = 'interaction'; end
for i = 16:18,feature_names{i} = sprintf('Categorical_%d', i-15); feature_types{i} = 'categorical'; end
for i = 19:CONFIG.n_features, feature_names{i} = sprintf('Noise_%d', i-18); feature_types{i} = 'noise'; end

% 데이터 정보 출력
fprintf('  - 샘플 수: %d\n', size(X, 1));
fprintf('  - 특성 수: %d\n', size(X, 2));
fprintf('  - 특성 유형: Linear(5), Nonlinear(5), Interaction(5), Categorical(3), Noise(%d)\n', n_noise);

if strcmp(CONFIG.problem_type, 'classification')
    fprintf('  - 클래스 분포:\n');
    tabulate(Y);
else
    fprintf('  - 타겟 범위: [%.2f, %.2f]\n', min(Y), max(Y));
end
fprintf('\n');

%% 3. 데이터 전처리
fprintf('[ 데이터 전처리 중... ]\n');
X_normalized = normalize(X);

if strcmp(CONFIG.problem_type, 'classification')
    cv = cvpartition(Y, 'HoldOut', CONFIG.test_ratio);
else
    cv = cvpartition(length(Y), 'HoldOut', CONFIG.test_ratio);
end
X_train = X_normalized(cv.training, :);
X_test  = X_normalized(cv.test, :);
Y_train = Y(cv.training);
Y_test  = Y(cv.test);

fprintf('  - 훈련 세트: %d samples\n', size(X_train, 1));
fprintf('  - 테스트 세트: %d samples\n\n', size(X_test, 1));

%% 4. 모델 정의 및 훈련
fprintf('[ 모델 훈련 중... ]\n\n');

% 결과 저장
results = table();
models = {};
predictions = {};
scores_all = {}; % ROC/PR 계산용 score 저장

% 양성 클래스 결정 (분류일 때)
if strcmp(CONFIG.problem_type, 'classification')
    [posLabel, ~] = getPositiveClass(Y);
else
    posLabel = [];
end

%% 4.1 Decision Tree
fprintf('  1. Decision Tree 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_tree = fitctree(X_train, Y_train, 'MaxNumSplits', 30, 'MinLeafSize', 10);
    else
        mdl_tree = fitrtree(X_train, Y_train, 'MaxNumSplits', 30, 'MinLeafSize', 10);
    end
    train_time = toc;

    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        [Y_pred_tree, score_tree] = predict(mdl_tree, X_test);
    else
        Y_pred_tree = predict(mdl_tree, X_test);
        score_tree = [];
    end
    pred_time = toc;

    models{end+1} = mdl_tree;
    predictions{end+1} = Y_pred_tree;
    scores_all{end+1} = score_tree;

    if strcmp(CONFIG.problem_type, 'classification')
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_tree, score_tree, posLabel);
    else
        [acc, pre, rec, f1] = evaluateModel(Y_test, Y_pred_tree, 'regression');
        roc_auc = NaN; pr_auc = NaN; mcc = NaN;
    end
    name = uniqueModelName(results, 'Decision_Tree');
    results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.2 Random Forest (Bagged Trees)
fprintf('  2. Random Forest (Bagged Trees) 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_rf = fitcensemble(X_train, Y_train, 'Method', 'Bag', 'NumLearningCycles', 100, ...
                              'Learners', templateTree('MaxNumSplits', 20));
    else
        mdl_rf = fitrensemble(X_train, Y_train, 'Method', 'Bag', 'NumLearningCycles', 100, ...
                              'Learners', templateTree('MaxNumSplits', 20));
    end
    train_time = toc;

    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        [Y_pred_rf, score_rf] = predict(mdl_rf, X_test);
    else
        Y_pred_rf = predict(mdl_rf, X_test);
        score_rf = [];
    end
    pred_time = toc;

    models{end+1} = mdl_rf;
    predictions{end+1} = Y_pred_rf;
    scores_all{end+1} = score_rf;

    if strcmp(CONFIG.problem_type, 'classification')
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_rf, score_rf, posLabel);
    else
        [acc, pre, rec, f1] = evaluateModel(Y_test, Y_pred_rf, 'regression');
        roc_auc = NaN; pr_auc = NaN; mcc = NaN;
    end
    name = uniqueModelName(results, 'Random_Forest');
    results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.3 AdaBoost
fprintf('  3. AdaBoost 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        if length(categories(Y)) == 2
            mdl_ada = fitcensemble(X_train, Y_train, 'Method', 'AdaBoostM1', 'NumLearningCycles', 100, ...
                                   'Learners', templateTree('MaxNumSplits', 10));
        else
            mdl_ada = fitcensemble(X_train, Y_train, 'Method', 'AdaBoostM2', 'NumLearningCycles', 100, ...
                                   'Learners', templateTree('MaxNumSplits', 10));
        end
    else
        mdl_ada = fitrensemble(X_train, Y_train, 'Method', 'LSBoost', 'NumLearningCycles', 100, ...
                               'Learners', templateTree('MaxNumSplits', 10));
    end
    train_time = toc;

    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        [Y_pred_ada, score_ada] = predict(mdl_ada, X_test);
    else
        Y_pred_ada = predict(mdl_ada, X_test);
        score_ada = [];
    end
    pred_time = toc;

    models{end+1} = mdl_ada;
    predictions{end+1} = Y_pred_ada;
    scores_all{end+1} = score_ada;

    if strcmp(CONFIG.problem_type, 'classification')
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_ada, score_ada, posLabel);
    else
        [acc, pre, rec, f1] = evaluateModel(Y_test, Y_pred_ada, 'regression');
        roc_auc = NaN; pr_auc = NaN; mcc = NaN;
    end
    name = uniqueModelName(results, 'AdaBoost');
    results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.4 LogitBoost (이진 분류만)
if strcmp(CONFIG.problem_type, 'classification') && length(categories(Y)) == 2
    fprintf('  4. LogitBoost 훈련 중...\n');
    try
        tic;
        mdl_logit = fitcensemble(X_train, Y_train, 'Method', 'LogitBoost', 'NumLearningCycles', 100, ...
                                 'Learners', templateTree('MaxNumSplits', 10), 'LearnRate', 0.1);
        train_time = toc;
        
        tic;
        [Y_pred_logit, score_logit] = predict(mdl_logit, X_test);
        pred_time = toc;
        
        models{end+1} = mdl_logit;
        predictions{end+1} = Y_pred_logit;
        scores_all{end+1} = score_logit;
        
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_logit, score_logit, posLabel);
        name = uniqueModelName(results, 'LogitBoost');
        results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
    catch ME
        fprintf('     에러: %s\n', ME.message);
    end
end

%% 4.5 GentleBoost
fprintf('  5. GentleBoost 훈련 중...\n');
try
    if strcmp(CONFIG.problem_type, 'classification')
        tic;
        mdl_gentle = fitcensemble(X_train, Y_train, 'Method', 'GentleBoost', 'NumLearningCycles', 100, ...
                                  'Learners', templateTree('MaxNumSplits', 10));
        train_time = toc;

        tic;
        [Y_pred_gentle, score_gentle] = predict(mdl_gentle, X_test);
        pred_time = toc;

        models{end+1} = mdl_gentle;
        predictions{end+1} = Y_pred_gentle;
        scores_all{end+1} = score_gentle;

        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_gentle, score_gentle, posLabel);
        name = uniqueModelName(results, 'GentleBoost');
        results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
    else
        fprintf('     GentleBoost는 분류만 지원합니다.\n');
    end
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.6 RUSBoost (클래스 불균형용)
if strcmp(CONFIG.problem_type, 'classification') && CONFIG.create_imbalance
    fprintf('  6. RUSBoost (불균형 처리) 훈련 중...\n');
    try
        tic;
        mdl_rus = fitcensemble(X_train, Y_train, 'Method', 'RUSBoost', 'NumLearningCycles', 100, ...
                               'Learners', templateTree('MaxNumSplits', 10));
        train_time = toc;

        tic;
        [Y_pred_rus, score_rus] = predict(mdl_rus, X_test);
        pred_time = toc;

        models{end+1} = mdl_rus;
        predictions{end+1} = Y_pred_rus;
        scores_all{end+1} = score_rus;

        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_rus, score_rus, posLabel);
        name = uniqueModelName(results, 'RUSBoost');
        results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
    catch ME
        fprintf('     에러: %s\n', ME.message);
    end
end

%% 4.7 SVM
fprintf('  7. SVM 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        if length(categories(Y)) == 2
            mdl_svm = fitcsvm(X_train, Y_train, 'KernelFunction', 'rbf', 'Standardize', true);
            % 확률 보정(Posterior)으로 점수 안정화
            mdl_svm = fitPosterior(mdl_svm, X_train, Y_train);
        else
            template = templateSVM('KernelFunction', 'rbf', 'Standardize', true);
            mdl_svm = fitcecoc(X_train, Y_train, 'Learners', template);
            % 필요시: mdl_svm = fitPosterior(mdl_svm, X_train, Y_train);
        end
    else
        mdl_svm = fitrsvm(X_train, Y_train, 'KernelFunction', 'rbf', 'Standardize', true);
    end
    train_time = toc;

    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        [Y_pred_svm, score_svm] = predict(mdl_svm, X_test);
    else
        Y_pred_svm = predict(mdl_svm, X_test);
        score_svm = [];
    end
    pred_time = toc;

    models{end+1} = mdl_svm;
    predictions{end+1} = Y_pred_svm;
    scores_all{end+1} = score_svm;

    if strcmp(CONFIG.problem_type, 'classification')
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_svm, score_svm, posLabel);
    else
        [acc, pre, rec, f1] = evaluateModel(Y_test, Y_pred_svm, 'regression');
        roc_auc = NaN; pr_auc = NaN; mcc = NaN;
    end
    name = uniqueModelName(results, 'SVM');
    results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.8 k-NN
fprintf('  8. k-NN 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_knn = fitcknn(X_train, Y_train, 'NumNeighbors', 5, 'Distance', 'euclidean', 'Standardize', true);
    else
        mdl_knn = fitrknn(X_train, Y_train, 'NumNeighbors', 5, 'Distance', 'euclidean', 'Standardize', true);
    end
    train_time = toc;

    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        [Y_pred_knn, score_knn] = predict(mdl_knn, X_test);
    else
        Y_pred_knn = predict(mdl_knn, X_test);
        score_knn = [];
    end
    pred_time = toc;

    models{end+1} = mdl_knn;
    predictions{end+1} = Y_pred_knn;
    scores_all{end+1} = score_knn;

    if strcmp(CONFIG.problem_type, 'classification')
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_knn, score_knn, posLabel);
    else
        [acc, pre, rec, f1] = evaluateModel(Y_test, Y_pred_knn, 'regression');
        roc_auc = NaN; pr_auc = NaN; mcc = NaN;
    end
    name = uniqueModelName(results, 'k-NN');
    results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.9 Naive Bayes (분류만)
if strcmp(CONFIG.problem_type, 'classification')
    fprintf('  9. Naive Bayes 훈련 중...\n');
    try
        tic;
        mdl_nb = fitcnb(X_train, Y_train, 'DistributionNames', 'kernel');
        train_time = toc;

        tic;
        [Y_pred_nb, score_nb] = predict(mdl_nb, X_test);
        pred_time = toc;

        models{end+1} = mdl_nb;
        predictions{end+1} = Y_pred_nb;
        scores_all{end+1} = score_nb;

        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_nb, score_nb, posLabel);
        name = uniqueModelName(results, 'Naive_Bayes');
        results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
    catch ME
        fprintf('     에러: %s\n', ME.message);
    end
end

%% 4.10 Neural Network
fprintf('  10. Neural Network 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_nn = fitcnet(X_train, Y_train, 'LayerSizes', [50, 25, 10], ...
            'Activations', 'relu', 'Standardize', true, 'Verbose', 0);
    else
        mdl_nn = fitrnet(X_train, Y_train, 'LayerSizes', [50, 25, 10], ...
            'Activations', 'relu', 'Standardize', true, 'Verbose', 0);
    end
    train_time = toc;

    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        [Y_pred_nn, score_nn] = predict(mdl_nn, X_test);
    else
        Y_pred_nn = predict(mdl_nn, X_test);
        score_nn = [];
    end
    pred_time = toc;

    models{end+1} = mdl_nn;
    predictions{end+1} = Y_pred_nn;
    scores_all{end+1} = score_nn;

    if strcmp(CONFIG.problem_type, 'classification')
        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_nn, score_nn, posLabel);
    else
        [acc, pre, rec, f1] = evaluateModel(Y_test, Y_pred_nn, 'regression');
        roc_auc = NaN; pr_auc = NaN; mcc = NaN;
    end
    name = uniqueModelName(results, 'Neural_Network');
    results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.11 Discriminant Analysis
fprintf('  11. Discriminant Analysis 훈련 중...\n');
try
    if strcmp(CONFIG.problem_type, 'classification')
        tic;
        mdl_da = fitcdiscr(X_train, Y_train, 'DiscrimType', 'quadratic');
        train_time = toc;

        tic;
        [Y_pred_da, score_da] = predict(mdl_da, X_test);
        pred_time = toc;

        models{end+1} = mdl_da;
        predictions{end+1} = Y_pred_da;
        scores_all{end+1} = score_da;

        [acc, pre, rec, f1, roc_auc, pr_auc, mcc] = ...
            evaluateModelPlus(Y_test, Y_pred_da, score_da, posLabel);
        name = uniqueModelName(results, 'Discriminant_Analysis');
        results = [results; createResultRow(name, acc, pre, rec, f1, roc_auc, pr_auc, mcc, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f, F1: %.3f)\n', acc, f1);
    else
        fprintf('     Discriminant Analysis는 분류만 지원합니다.\n');
    end
catch ME
    if contains(ME.message, '특이 공분산 행렬이 있기 때문에 ''quadratic'' 유형을 사용할 수 없습니다.')
        fprintf('     에러: 하나 이상의 클래스에 특이 공분산 행렬이 있기 때문에 ''quadratic'' 유형을 사용할 수 없습니다.\n');
    else
        fprintf('     에러: %s\n', ME.message);
    end
end

fprintf('\n');

%% 5. 교차검증
fprintf('[ 교차검증 수행 중... ]\n');

cv_scores = table();
if strcmp(CONFIG.problem_type, 'classification')
    cv_partition = cvpartition(Y_train, 'KFold', CONFIG.cv_folds);
else
    cv_partition = cvpartition(length(Y_train), 'KFold', CONFIG.cv_folds);
end

for i = 1:height(results)
    model_name = results.Model{i};
    fprintf('  %s 교차검증 중...\n', model_name);
    
    scores = zeros(CONFIG.cv_folds, 1);
    
    for fold = 1:CONFIG.cv_folds
        train_idx = training(cv_partition, fold);
        val_idx = test(cv_partition, fold);
        
        X_cv_train = X_train(train_idx, :);
        Y_cv_train = Y_train(train_idx);
        X_cv_val = X_train(val_idx, :);
        Y_cv_val = Y_train(val_idx);
        
        try
            % 간단한 모델 재훈련 (빠른 실행을 위해)
            cv_model = trainSimpleModel(model_name, X_cv_train, Y_cv_train, CONFIG.problem_type);
            Y_cv_pred = predict(cv_model, X_cv_val);
            [acc_cv, ~, ~, ~] = evaluateModel(Y_cv_val, Y_cv_pred, CONFIG.problem_type);
            scores(fold) = acc_cv;
        catch
            scores(fold) = 0;
        end
    end
    
    cv_scores = [cv_scores; table({model_name}, mean(scores), std(scores), ...
        'VariableNames', {'Model', 'CV_Mean', 'CV_Std'})];
end
fprintf('\n');

%% 6. 특성 중요도 분석
if CONFIG.calculate_importance
    fprintf('[ 특성 중요도 분석 중... ]\n');
    
    feature_importance = table();
    
    % Random Forest의 특성 중요도
    rf_idx = find(contains(results.Model, 'Random_Forest'), 1, 'first');
    if ~isempty(rf_idx) && rf_idx <= length(models)
        try
            rf_importance = predictorImportance(models{rf_idx});
            feature_importance.Random_Forest = rf_importance';
        catch
            fprintf('  Random Forest 특성 중요도 계산 실패\n');
        end
    end
    
    % ANOVA F-값 (분류) 또는 상관계수 (회귀)
    if strcmp(CONFIG.problem_type, 'classification')
        f_values = zeros(CONFIG.n_features, 1);
        for i = 1:CONFIG.n_features
            try
                [~, tbl] = anova1(X_train(:, i), Y_train, 'off');
                f_values(i) = tbl{2, 5};
            catch
                f_values(i) = 0;
            end
        end
        feature_importance.ANOVA_F = f_values;
    else
        corr_values = abs(corr(X_train, Y_train));
        feature_importance.Correlation = corr_values;
    end
    
    fprintf('  완료!\n\n');
end

%% 7. 결과 요약
fprintf('=====================================\n');
fprintf('           결과 요약\n');
fprintf('=====================================\n\n');

% 테스트 성능
fprintf('[ 테스트 세트 성능 ]\n');
disp(results);

% 교차검증 결과
fprintf('\n[ 교차검증 결과 ]\n');
disp(cv_scores);

% 최고 성능 모델 (F1 기준)
[~, best_idx] = max(results.F1_Score);
fprintf('\n[ 최고 성능 모델 ]\n');
fprintf('  모델: %s\n', results.Model{best_idx});
fprintf('  정확도: %.4f\n', results.Accuracy(best_idx));
fprintf('  F1-Score: %.4f\n', results.F1_Score(best_idx));
if strcmp(CONFIG.problem_type, 'classification')
    fprintf('  ROC-AUC: %.4f\n', results.ROC_AUC(best_idx));
    fprintf('  PR-AUC: %.4f\n', results.PR_AUC(best_idx));
    fprintf('  MCC: %.4f\n', results.MCC(best_idx));
end
fprintf('  훈련 시간: %.4f초\n', results.Train_Time(best_idx));
fprintf('  예측 시간: %.4f초\n\n', results.Pred_Time(best_idx));

%% 8. 시각화
if CONFIG.plot_results
    fprintf('[ 결과 시각화 중... ]\n');
    
    % Figure 1: 모델 성능 비교 (확장: ROC/PR/MCC 포함) → 2x5 레이아웃
    figure('Name', '모델 성능 비교', 'Position', [50, 50, 1700, 900]);
    
    subplot(2, 5, 1);
    bar(categorical(results.Model), results.Accuracy);
    title('모델별 Accuracy', 'FontSize', 12); ylabel('Accuracy'); xtickangle(45); grid on;
    
    subplot(2, 5, 2);
    bar(categorical(results.Model), results.Precision);
    title('모델별 Precision', 'FontSize', 12); ylabel('Precision'); xtickangle(45); grid on;
    
    subplot(2, 5, 3);
    bar(categorical(results.Model), results.Recall);
    title('모델별 Recall', 'FontSize', 12); ylabel('Recall'); xtickangle(45); grid on;
    
    subplot(2, 5, 4);
    bar(categorical(results.Model), results.F1_Score);
    title('모델별 F1-Score', 'FontSize', 12); ylabel('F1'); xtickangle(45); grid on;

    subplot(2, 5, 5);
    bar(categorical(results.Model), results.MCC);
    title('모델별 MCC', 'FontSize', 12); ylabel('MCC'); xtickangle(45); grid on;

    subplot(2, 5, 6);
    bar(categorical(results.Model), results.ROC_AUC);
    title('모델별 ROC-AUC', 'FontSize', 12); ylabel('AUC'); xtickangle(45); grid on;

    subplot(2, 5, 7);
    bar(categorical(results.Model), results.PR_AUC);
    title('모델별 PR-AUC', 'FontSize', 12); ylabel('AUC'); xtickangle(45); grid on;

    subplot(2, 5, 8);
    bar(categorical(results.Model), results.Train_Time);
    title('모델별 훈련 시간', 'FontSize', 12); ylabel('Seconds'); xtickangle(45); grid on;

    subplot(2, 5, 9);
    bar(categorical(results.Model), results.Pred_Time * 1000);
    title('모델별 예측 시간', 'FontSize', 12); ylabel('Milliseconds'); xtickangle(45); grid on;

    subplot(2, 5, 10);
    metrics = [results.Accuracy, results.Precision, results.Recall, results.F1_Score, results.MCC];
    plot(metrics', 'o-', 'LineWidth', 1.5);
    title('종합 성능 비교', 'FontSize', 12);
    legend(strrep(results.Model, '_', ' '), 'Location', 'eastoutside', 'FontSize', 8);
    xticks(1:5); xticklabels({'Acc','Prec','Rec','F1','MCC'}); grid on;

    % Figure 2: Confusion Matrices (분류 문제만)
    if strcmp(CONFIG.problem_type, 'classification') && ~isempty(predictions)
        figure('Name', 'Confusion Matrices', 'Position', [50, 50, 1500, 900]);
        n_models = length(predictions);
        cols = ceil(sqrt(n_models));
        rows = ceil(n_models / cols);
        for i = 1:n_models
            subplot(rows, cols, i);
            cm = confusionmat(Y_test, predictions{i});
            imagesc(cm);
            colorbar; colormap(parula);
            title(strrep(results.Model{i}, '_', ' '));
            xlabel('Predicted'); ylabel('Actual');
            [r, c] = size(cm);
            for j = 1:r
                for k = 1:c
                    text(k, j, num2str(cm(j,k)), 'HorizontalAlignment', 'center', ...
                        'Color', 'white', 'FontWeight', 'bold');
                end
            end
        end
    end
    
    % Figure 3: 특성 중요도
    if CONFIG.calculate_importance && exist('feature_importance','var') && ~isempty(feature_importance)
        figure('Name', '특성 중요도 분석', 'Position', [50, 50, 1200, 700]);
        importance_cols = feature_importance.Properties.VariableNames;
        n_importance = length(importance_cols);
        for i = 1:n_importance
            subplot(1, n_importance, i);
            importance_values = feature_importance.(importance_cols{i});
            [sorted_values, sorted_idx] = sort(importance_values, 'descend');
            top_k = min(15, length(sorted_values));
            colors = zeros(top_k, 3);
            for j = 1:top_k
                type = feature_types{sorted_idx(j)};
                switch type
                    case 'linear',      colors(j,:) = [0.2, 0.6, 1];
                    case 'nonlinear',   colors(j,:) = [1, 0.4, 0.4];
                    case 'interaction', colors(j,:) = [0.4, 0.8, 0.4];
                    case 'categorical', colors(j,:) = [1, 0.8, 0.2];
                    case 'noise',       colors(j,:) = [0.7, 0.7, 0.7];
                end
            end
            barh(1:top_k, sorted_values(1:top_k), 'FaceColor', 'flat', 'CData', colors);
            set(gca, 'YTick', 1:top_k, 'YTickLabel', feature_names(sorted_idx(1:top_k)));
            xlabel('Importance'); title(strrep(importance_cols{i}, '_', ' ')); grid on;
        end
        legend({'Linear','Nonlinear','Interaction','Categorical','Noise'}, ...
               'Location', 'southoutside', 'Orientation', 'horizontal');
    end
    
    % Figure 4: 성능 대시보드
    figure('Name', '모델 성능 대시보드', 'Position', [50, 50, 1400, 800]);
    overall_score = 0.35*results.F1_Score + 0.25*results.Accuracy + 0.2*nanfill(results.PR_AUC) + 0.2*nanfill(results.ROC_AUC);
    efficiency_score = 1 ./ (1 + results.Train_Time + 10 * results.Pred_Time);
    efficiency_score = efficiency_score / max(efficiency_score);

    subplot(2, 3, 1);
    [sorted_score, sort_idx] = sort(overall_score, 'descend');
    barh(1:length(sorted_score), sorted_score);
    set(gca, 'YTick', 1:length(sorted_score), 'YTickLabel', results.Model(sort_idx));
    xlabel('종합 성능 점수'); title('모델 종합 순위'); grid on;

    subplot(2, 3, 2);
    scatter(results.F1_Score, efficiency_score, 100, 'filled');
    text(results.F1_Score, efficiency_score, results.Model, 'VerticalAlignment','bottom', ...
        'HorizontalAlignment','center', 'FontSize', 8);
    xlabel('F1-Score'); ylabel('효율성 점수'); title('성능 vs 효율성'); grid on;

    subplot(2, 3, 3);
    scatter(results.Train_Time, results.Pred_Time * 1000, 100, results.F1_Score, 'filled');
    colorbar; xlabel('훈련 시간 (초)'); ylabel('예측 시간 (ms)'); title('시간 효율성 (색상: F1)'); grid on;

    subplot(2, 3, [4, 5, 6]);
    top5_idx = sort_idx(1:min(5, length(sort_idx)));
    top5_metrics = [results.Accuracy(top5_idx), results.Precision(top5_idx), ...
                    results.Recall(top5_idx), results.F1_Score(top5_idx), ...
                    nanfill(results.MCC(top5_idx))];
    h = bar(categorical(results.Model(top5_idx)), top5_metrics);
    legend({'Acc','Prec','Rec','F1','MCC'}, 'Location', 'best');
    title('Top 5 모델 상세 성능'); ylabel('Score'); grid on;
    
    fprintf('  시각화 완료!\n\n');
end

fprintf('=====================================\n');
fprintf('        분석 완료!\n');
fprintf('=====================================\n');


function v = ensureColumn(v, p)
    % v를 p×1 열벡터로 강제. 크기 불일치/NaN 시 안전 보정.
    if isempty(v), v = zeros(p,1); return; end
    v = v(:);
    if numel(v) ~= p
        warning('ensureColumn:LengthMismatch', 'Vector length %d ~= p (%d). Filling zeros.', numel(v), p);
        v = zeros(p,1);
    end
    v(isnan(v)) = 0;
end

function M = shapToMatrix(expl)
    % shapley 결과 expl.ShapleyValues를 nShap×p 숫자 행렬로 변환
    vals = expl.ShapleyValues;

    % table → array
    if istable(vals)
        try
            vals = table2array(vals);
        catch
            C = vals{:,:}; % 테이블 내용
            if iscell(C)
                vals = cell2mat(C);
            else
                vals = C;
            end
        end
    end

    % cell → 3D concat (클래스별), 이후 클래스 평균
    if iscell(vals)
        vals = cat(3, vals{:});     % nShap×p×C
    end

    if ndims(vals) == 3
        vals = mean(vals, 3);       % nShap×p
    end

    % 혹시 p×nShap 형태면 전치
    if size(vals,2) ~= size(expl.Data,2)
        % expl.Data는 nShap×p 행렬
        if size(vals,1) == size(expl.Data,2) && size(vals,2) == size(expl.Data,1)
            vals = vals.';
        end
    end

    M = vals; % nShap×p
end

function v = normalizePositive(v)
    v = v(:);
    v(v < 0) = 0;
    s = sum(v);
    if s > 0
        v = v / s;
    else
        v = ones(size(v)) / numel(v);
    end
end

function thr = findBestThreshold(score, Ybin)
    % score: 연속 스코어 (큰 값일수록 양성)
    % Ybin : categorical 이진 레이블
    if ~iscategorical(Ybin), Ybin = categorical(Ybin); end
    classes = categories(Ybin);
    posLabel = 'Class_1'; if ~ismember(posLabel, classes), posLabel = classes{end}; end
    negLabel = setdiff(classes, posLabel, 'stable');
    if isempty(negLabel), negLabel = {classes{1}}; end

    % 후보 임계값 (분위수 기반)
    qs = unique(quantile(score, linspace(0.05,0.95,41)));
    if numel(qs) < 2, qs = linspace(min(score), max(score), 21); end

    bestF1 = -Inf; thr = median(score);
    for t = qs(:)'
        Yh = categorical(score >= t, [false true], {negLabel{1} posLabel});
        [~,~,~,f1] = evaluateModel(Ybin, Yh, 'classification');
        if f1 > bestF1
            bestF1 = f1; thr = t;
        end
    end
end

function y = sigmoid(x)
    y = 1 ./ (1 + exp(-x));
end



%% 보조 함수들
function x = nanfill(x, val)
    if nargin < 2, val = 0; end
    x(isnan(x)) = val;
end

function [accuracy, precision, recall, f1] = evaluateModel(Y_true, Y_pred, problem_type)
    % 기본 평가 함수 (원본 유지: 회귀는 R2 기반)
    if strcmp(problem_type, 'classification')
        if iscategorical(Y_true)
            accuracy = sum(Y_true == Y_pred) / length(Y_true);
            classes = categories(Y_true);
            n_classes = length(classes);
            precision = zeros(n_classes, 1);
            recall = zeros(n_classes, 1);
            f1 = zeros(n_classes, 1);
            for i = 1:n_classes
                tp = sum((Y_true == classes{i}) & (Y_pred == classes{i}));
                fp = sum((Y_true ~= classes{i}) & (Y_pred == classes{i}));
                fn = sum((Y_true == classes{i}) & (Y_pred ~= classes{i}));
                precision(i) = tp/(tp+fp+eps);
                recall(i)    = tp/(tp+fn+eps);
                f1(i)        = 2*precision(i)*recall(i)/(precision(i)+recall(i)+eps);
            end
            precision = mean(precision); recall = mean(recall); f1 = mean(f1);
        else
            accuracy = sum(Y_true == Y_pred) / length(Y_true);
            precision = accuracy; recall = accuracy; f1 = accuracy;
        end
    else
        mse = mean((Y_true - Y_pred).^2);
        rmse = sqrt(mse);
        mae = mean(abs(Y_true - Y_pred));
        ss_res = sum((Y_true - Y_pred).^2);
        ss_tot = sum((Y_true - mean(Y_true)).^2);
        r2 = 1 - (ss_res / (ss_tot + eps));
        accuracy = max(0, r2);  % 음수 방지
        precision = 1 / (1 + rmse);
        recall = 1 / (1 + mae);
        f1 = 2 * precision * recall / (precision + recall + eps);
    end
end

function result_row = createResultRow(model_name, accuracy, precision, recall, f1, ...
                                      roc_auc, pr_auc, mcc, train_time, pred_time)
    % 결과 테이블 행 생성 (확장)
    result_row = table({model_name}, accuracy, precision, recall, f1, ...
        roc_auc, pr_auc, mcc, train_time, pred_time, ...
        'VariableNames', {'Model','Accuracy','Precision','Recall','F1_Score', ...
                          'ROC_AUC','PR_AUC','MCC','Train_Time','Pred_Time'});
end

function model = trainSimpleModel(model_name, X, Y, problem_type)
    % 교차검증용 간단한 모델 훈련
    base = erase(model_name, {'_v2','_v3','_v4','_v5'});
    switch base
        case 'Decision_Tree'
            model = strcmp(problem_type,'classification') ...
                .* fitctree(X, Y, 'MaxNumSplits', 20) + ...
                strcmp(problem_type,'regression') ...
                .* fitrtree(X, Y, 'MaxNumSplits', 20);
            if ~iscell(model), model = model(~cellfun('isempty',{model})); end %#ok<NASGU>
            if strcmp(problem_type,'classification'), model = fitctree(X, Y, 'MaxNumSplits', 20); else, model = fitrtree(X, Y, 'MaxNumSplits', 20); end
        case 'Random_Forest'
            if strcmp(problem_type, 'classification')
                model = fitcensemble(X, Y, 'Method', 'Bag', 'NumLearningCycles', 50);
            else
                model = fitrensemble(X, Y, 'Method', 'Bag', 'NumLearningCycles', 50);
            end
        case 'k-NN'
            if strcmp(problem_type, 'classification')
                model = fitcknn(X, Y, 'NumNeighbors', 5);
            else
                model = fitrknn(X, Y, 'NumNeighbors', 5);
            end
        otherwise
            if strcmp(problem_type, 'classification')
                model = fitctree(X, Y);
            else
                model = fitrtree(X, Y);
            end
    end
end

function [accuracy, precision, recall, f1, roc_auc, pr_auc, mcc] = ...
         evaluateModelPlus(Y_true, Y_pred, score, posLabel)
    % 분류 성능 확장 지표 계산: ROC-AUC, PR-AUC, MCC 포함
    [accuracy, precision, recall, f1] = evaluateModel(Y_true, Y_pred, 'classification');

    % MCC (이진 분류 기준)
    cm = confusionmat(Y_true, Y_pred);
    if isequal(size(cm), [2,2])
        TP = cm(2,2); TN = cm(1,1); FP = cm(1,2); FN = cm(2,1);
        denom = sqrt((TP+FP)*(TP+FN)*(TN+FP)*(TN+FN));
        mcc = ((TP*TN) - (FP*FN)) / max(denom, eps);
    else
        mcc = NaN; % 다중분류는 확장 버전 별도 필요
    end

    % ROC-AUC / PR-AUC
    roc_auc = NaN; pr_auc = NaN;
    if ~isempty(score)
        classes = categories(Y_true);
        % score의 열이 클래스 순서를 따른다고 가정 (fit* 대부분 동일)
        posIdx = find(strcmp(classes, posLabel), 1);
        if isempty(posIdx)
            if size(score,2) == 2
                posIdx = 2; % 관례적으로 2열을 양성
                posLabel = classes{posIdx};
            else
                [~, posIdx] = max(mean(score,1));
                posLabel = classes{posIdx};
            end
        end
        scoresPos = score(:, posIdx);
        try
            [~,~,~,roc_auc] = perfcurve(Y_true, scoresPos, posLabel);
            [~,~,~,pr_auc]  = perfcurve(Y_true, scoresPos, posLabel, 'xCrit','reca','yCrit','prec');
        catch
            % perfcurve 실패 시 NaN 유지
        end
    end
end




function name = uniqueModelName(results, base)
    % results 테이블이 비었거나 'Model' 변수가 아직 없을 때 안전 처리
    if isempty(results) || ~ismember('Model', results.Properties.VariableNames)
        name = base;
        return;
    end
    name = base;
    k = 2;
    while any(strcmp(results.Model, name))
        name = sprintf('%s_v%d', base, k);
        k = k + 1;
    end
end



%%


function [posLabel, posIdx] = getPositiveClass(Y)
    % 양성 클래스를 안정적으로 선택
    classes = categories(Y);
    if any(strcmp(classes, 'Class_1'))
        posLabel = 'Class_1';
        posIdx = find(strcmp(classes, 'Class_1'));
        return;
    end
    counts = countcats(Y);
    [~, posIdx] = min(counts);
    posLabel = classes{posIdx};
end


