%% 포괄적인 MATLAB 머신러닝 모델 성능 비교 분석
% Comprehensive Machine Learning Model Performance Comparison
% 작성일: 2024
% 목적: 다양한 앙상블 및 고급 머신러닝 모델의 체계적 비교

clear; clc; close all;
warning('off', 'all');


results = cell2table(cell(0,10), 'VariableNames', ...
    {'Model','Accuracy','Precision','Recall','F1_Score', ...
     'ROC_AUC','PR_AUC','MCC','Train_Time','Pred_Time'});
%% 1. 설정 및 초기화
fprintf('=====================================\n');
fprintf('   포괄적 머신러닝 모델 비교 분석\n');
fprintf('=====================================\n\n');

% 분석 설정
CONFIG = struct();
CONFIG.problem_type = 'classification';     % 'classification' or 'regression'
CONFIG.n_samples = 2000;                   % 샘플 수
CONFIG.n_features = 25;                    % 특성 수
CONFIG.test_ratio = 0.2;                   % 테스트 세트 비율
CONFIG.cv_folds = 5;                        % 교차검증 폴드 수
CONFIG.random_seed = 42;                   % 랜덤 시드
CONFIG.create_imbalance = true;            % 클래스 불균형 생성
CONFIG.add_outliers = true;                % 이상값 추가
CONFIG.outlier_ratio = 0.05;               % 이상값 비율
CONFIG.calculate_importance = true;        % 특성 중요도 계산
CONFIG.plot_results = true;                % 결과 시각화

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
    % 복잡한 결정 경계
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
        keep_ratio = 0.3;  % 30% 만 유지
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

% 특성 이름 생성
feature_names = cell(1, CONFIG.n_features);
feature_types = {};
for i = 1:5
    feature_names{i} = sprintf('Linear_%d', i);
    feature_types{i} = 'linear';
end
for i = 6:10
    feature_names{i} = sprintf('Nonlinear_%d', i-5);
    feature_types{i} = 'nonlinear';
end
for i = 11:15
    feature_names{i} = sprintf('Interaction_%d', i-10);
    feature_types{i} = 'interaction';
end
for i = 16:18
    feature_names{i} = sprintf('Categorical_%d', i-15);
    feature_types{i} = 'categorical';
end
for i = 19:CONFIG.n_features
    feature_names{i} = sprintf('Noise_%d', i-18);
    feature_types{i} = 'noise';
end

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

% 특성 정규화
X_normalized = normalize(X);

% 데이터 분할
cv = cvpartition(Y, 'HoldOut', CONFIG.test_ratio);
X_train = X_normalized(cv.training, :);
Y_train = Y(cv.training);
X_test = X_normalized(cv.test, :);
Y_test = Y(cv.test);

fprintf('  - 훈련 세트: %d samples\n', size(X_train, 1));
fprintf('  - 테스트 세트: %d samples\n', size(X_test, 1));
fprintf('\n');

%% 4. 모델 정의 및 훈련
fprintf('[ 모델 훈련 중... ]\n\n');

% 결과 저장을 위한 구조체
results = table();
models = {};
predictions = {};

%% 4.1 Decision Tree
fprintf('  1. Decision Tree 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_tree = fitctree(X_train, Y_train, ...
            'MaxNumSplits', 30, ...
            'MinLeafSize', 10);
    else
        mdl_tree = fitrtree(X_train, Y_train, ...
            'MaxNumSplits', 30, ...
            'MinLeafSize', 10);
    end
    train_time = toc;
    
    tic;
    Y_pred_tree = predict(mdl_tree, X_test);
    pred_time = toc;
    
    models{end+1} = mdl_tree;
    predictions{end+1} = Y_pred_tree;
    
    [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_tree, CONFIG.problem_type);
    results = [results; createResultRow('Decision_Tree', accuracy, precision, recall, f1, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f)\n', accuracy);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.2 Random Forest (Bagged Trees)
fprintf('  2. Random Forest (Bagged Trees) 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_rf = fitcensemble(X_train, Y_train, ...
            'Method', 'Bag', ...
            'NumLearningCycles', 100, ...
            'Learners', templateTree('MaxNumSplits', 20));
    else
        mdl_rf = fitrensemble(X_train, Y_train, ...
            'Method', 'Bag', ...
            'NumLearningCycles', 100, ...
            'Learners', templateTree('MaxNumSplits', 20));
    end
    train_time = toc;
    
    tic;
    Y_pred_rf = predict(mdl_rf, X_test);
    pred_time = toc;
    
    models{end+1} = mdl_rf;
    predictions{end+1} = Y_pred_rf;
    
    [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_rf, CONFIG.problem_type);
    results = [results; createResultRow('Random_Forest', accuracy, precision, recall, f1, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f)\n', accuracy);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.3 AdaBoost
fprintf('  3. AdaBoost 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        if length(categories(Y)) == 2
            mdl_ada = fitcensemble(X_train, Y_train, ...
                'Method', 'AdaBoostM1', ...
                'NumLearningCycles', 100, ...
                'Learners', templateTree('MaxNumSplits', 10));
        else
            mdl_ada = fitcensemble(X_train, Y_train, ...
                'Method', 'AdaBoostM2', ...
                'NumLearningCycles', 100, ...
                'Learners', templateTree('MaxNumSplits', 10));
        end
    else
        mdl_ada = fitrensemble(X_train, Y_train, ...
            'Method', 'LSBoost', ...
            'NumLearningCycles', 100, ...
            'Learners', templateTree('MaxNumSplits', 10));
    end
    train_time = toc;
    
    tic;
    Y_pred_ada = predict(mdl_ada, X_test);
    pred_time = toc;
    
    models{end+1} = mdl_ada;
    predictions{end+1} = Y_pred_ada;
    
    [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_ada, CONFIG.problem_type);
    results = [results; createResultRow('AdaBoost', accuracy, precision, recall, f1, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f)\n', accuracy);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.4 LogitBoost (이진 분류만)
if strcmp(CONFIG.problem_type, 'classification') && length(categories(Y)) == 2
    fprintf('  4. LogitBoost 훈련 중...\n');
    try
        tic;
        mdl_logit = fitcensemble(X_train, Y_train, ...
            'Method', 'LogitBoost', ...
            'NumLearningCycles', 100, ...
            'Learners', templateTree('MaxNumSplits', 10), ...
            'LearnRate', 0.1);
        train_time = toc;
        
        tic;
        Y_pred_logit = predict(mdl_logit, X_test);
        pred_time = toc;
        
        models{end+1} = mdl_logit;
        predictions{end+1} = Y_pred_logit;
        
        [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_logit, CONFIG.problem_type);
        results = [results; createResultRow('LogitBoost', accuracy, precision, recall, f1, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f)\n', accuracy);
    catch ME
        fprintf('     에러: %s\n', ME.message);
    end
end

%% 4.5 GentleBoost
fprintf('  5. GentleBoost 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_gentle = fitcensemble(X_train, Y_train, ...
            'Method', 'GentleBoost', ...
            'NumLearningCycles', 100, ...
            'Learners', templateTree('MaxNumSplits', 10));
    else
        % GentleBoost는 분류만 지원
        fprintf('     GentleBoost는 분류만 지원합니다.\n');
        mdl_gentle = [];
    end
    
    if ~isempty(mdl_gentle)
        train_time = toc;
        
        tic;
        Y_pred_gentle = predict(mdl_gentle, X_test);
        pred_time = toc;
        
        models{end+1} = mdl_gentle;
        predictions{end+1} = Y_pred_gentle;
        
        [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_gentle, CONFIG.problem_type);
        results = [results; createResultRow('GentleBoost', accuracy, precision, recall, f1, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f)\n', accuracy);
    end
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.6 RUSBoost (클래스 불균형용)
if strcmp(CONFIG.problem_type, 'classification') && CONFIG.create_imbalance
    fprintf('  6. RUSBoost (불균형 처리) 훈련 중...\n');
    try
        tic;
        mdl_rus = fitcensemble(X_train, Y_train, ...
            'Method', 'RUSBoost', ...
            'NumLearningCycles', 100, ...
            'Learners', templateTree('MaxNumSplits', 10));
        train_time = toc;
        
        tic;
        Y_pred_rus = predict(mdl_rus, X_test);
        pred_time = toc;
        
        models{end+1} = mdl_rus;
        predictions{end+1} = Y_pred_rus;
        
        [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_rus, CONFIG.problem_type);
        results = [results; createResultRow('RUSBoost', accuracy, precision, recall, f1, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f)\n', accuracy);
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
            mdl_svm = fitcsvm(X_train, Y_train, ...
                'KernelFunction', 'rbf', ...
                'Standardize', true);
        else
            template = templateSVM('KernelFunction', 'rbf', 'Standardize', true);
            mdl_svm = fitcecoc(X_train, Y_train, 'Learners', template);
        end
    else
        mdl_svm = fitrsvm(X_train, Y_train, ...
            'KernelFunction', 'rbf', ...
            'Standardize', true);
    end
    train_time = toc;
    
    tic;
    Y_pred_svm = predict(mdl_svm, X_test);
    pred_time = toc;
    
    models{end+1} = mdl_svm;
    predictions{end+1} = Y_pred_svm;
    
    [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_svm, CONFIG.problem_type);
    results = [results; createResultRow('SVM', accuracy, precision, recall, f1, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f)\n', accuracy);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.8 k-NN
fprintf('  8. k-NN 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_knn = fitcknn(X_train, Y_train, ...
            'NumNeighbors', 5, ...
            'Distance', 'euclidean', ...
            'Standardize', true);
    else
        mdl_knn = fitrknn(X_train, Y_train, ...
            'NumNeighbors', 5, ...
            'Distance', 'euclidean', ...
            'Standardize', true);
    end
    train_time = toc;
    
    tic;
    Y_pred_knn = predict(mdl_knn, X_test);
    pred_time = toc;
    
    models{end+1} = mdl_knn;
    predictions{end+1} = Y_pred_knn;
    
    [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_knn, CONFIG.problem_type);
    results = [results; createResultRow('k-NN', accuracy, precision, recall, f1, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f)\n', accuracy);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.9 Naive Bayes (분류만)
if strcmp(CONFIG.problem_type, 'classification')
    fprintf('  9. Naive Bayes 훈련 중...\n');
    try
        tic;
        mdl_nb = fitcnb(X_train, Y_train, ...
            'DistributionNames', 'kernel');  % 커널 밀도 추정 사용
        train_time = toc;
        
        tic;
        Y_pred_nb = predict(mdl_nb, X_test);
        pred_time = toc;
        
        models{end+1} = mdl_nb;
        predictions{end+1} = Y_pred_nb;
        
        [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_nb, CONFIG.problem_type);
        results = [results; createResultRow('Naive_Bayes', accuracy, precision, recall, f1, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f)\n', accuracy);
    catch ME
        fprintf('     에러: %s\n', ME.message);
    end
end

%% 4.10 Neural Network
fprintf('  10. Neural Network 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_nn = fitcnet(X_train, Y_train, ...
            'LayerSizes', [50, 25, 10], ...
            'Activations', 'relu', ...
            'Standardize', true, ...
            'Verbose', 0);
    else
        mdl_nn = fitrnet(X_train, Y_train, ...
            'LayerSizes', [50, 25, 10], ...
            'Activations', 'relu', ...
            'Standardize', true, ...
            'Verbose', 0);
    end
    train_time = toc;
    
    tic;
    Y_pred_nn = predict(mdl_nn, X_test);
    pred_time = toc;
    
    models{end+1} = mdl_nn;
    predictions{end+1} = Y_pred_nn;
    
    [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_nn, CONFIG.problem_type);
    results = [results; createResultRow('Neural_Network', accuracy, precision, recall, f1, train_time, pred_time)];
    fprintf('     완료! (정확도: %.3f)\n', accuracy);
catch ME
    fprintf('     에러: %s\n', ME.message);
end

%% 4.11 Discriminant Analysis
fprintf('  11. Discriminant Analysis 훈련 중...\n');
try
    tic;
    if strcmp(CONFIG.problem_type, 'classification')
        mdl_da = fitcdiscr(X_train, Y_train, ...
            'DiscrimType', 'quadratic');
        train_time = toc;
        
        tic;
        Y_pred_da = predict(mdl_da, X_test);
        pred_time = toc;
        
        models{end+1} = mdl_da;
        predictions{end+1} = Y_pred_da;
        
        [accuracy, precision, recall, f1] = evaluateModel(Y_test, Y_pred_da, CONFIG.problem_type);
        results = [results; createResultRow('Discriminant_Analysis', accuracy, precision, recall, f1, train_time, pred_time)];
        fprintf('     완료! (정확도: %.3f)\n', accuracy);
    else
        fprintf('     Discriminant Analysis는 분류만 지원합니다.\n');
    end
catch ME
    fprintf('     에러: %s\n', ME.message);
end

fprintf('\n');

%% 5. 교차검증
fprintf('[ 교차검증 수행 중... ]\n');

cv_scores = table();
cv_partition = cvpartition(Y_train, 'KFold', CONFIG.cv_folds);

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
            [acc, ~, ~, ~] = evaluateModel(Y_cv_val, Y_cv_pred, CONFIG.problem_type);
            scores(fold) = acc;
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
    rf_idx = find(strcmp(results.Model, 'Random_Forest'));
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
            [~, tbl] = anova1(X_train(:, i), Y_train, 'off');
            f_values(i) = tbl{2, 5};
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

% 최고 성능 모델
[~, best_idx] = max(results.F1_Score);
fprintf('\n[ 최고 성능 모델 ]\n');
fprintf('  모델: %s\n', results.Model{best_idx});
fprintf('  정확도: %.4f\n', results.Accuracy(best_idx));
fprintf('  F1-Score: %.4f\n', results.F1_Score(best_idx));
fprintf('  훈련 시간: %.4f초\n', results.Train_Time(best_idx));
fprintf('  예측 시간: %.4f초\n\n', results.Pred_Time(best_idx));

%% 8. 시각화
if CONFIG.plot_results
    fprintf('[ 결과 시각화 중... ]\n');
    
    % Figure 1: 모델 성능 비교
    figure('Name', '모델 성능 비교', 'Position', [50, 50, 1500, 900]);
    
    % 정확도 비교
    subplot(2, 4, 1);
    bar(categorical(results.Model), results.Accuracy);
    title('모델별 정확도', 'FontSize', 12);
    ylabel('Accuracy');
    xtickangle(45);
    grid on;
    
    % Precision 비교
    subplot(2, 4, 2);
    bar(categorical(results.Model), results.Precision);
    title('모델별 Precision', 'FontSize', 12);
    ylabel('Precision');
    xtickangle(45);
    grid on;
    
    % Recall 비교
    subplot(2, 4, 3);
    bar(categorical(results.Model), results.Recall);
    title('모델별 Recall', 'FontSize', 12);
    ylabel('Recall');
    xtickangle(45);
    grid on;
    
    % F1-Score 비교
    subplot(2, 4, 4);
    bar(categorical(results.Model), results.F1_Score);
    title('모델별 F1-Score', 'FontSize', 12);
    ylabel('F1-Score');
    xtickangle(45);
    grid on;
    
    % 훈련 시간 비교
    subplot(2, 4, 5);
    bar(categorical(results.Model), results.Train_Time);
    title('모델별 훈련 시간', 'FontSize', 12);
    ylabel('Time (seconds)');
    xtickangle(45);
    grid on;
    
    % 예측 시간 비교
    subplot(2, 4, 6);
    bar(categorical(results.Model), results.Pred_Time * 1000);
    title('모델별 예측 시간', 'FontSize', 12);
    ylabel('Time (milliseconds)');
    xtickangle(45);
    grid on;
    
    % 교차검증 성능
    subplot(2, 4, 7);
    bar(categorical(cv_scores.Model), cv_scores.CV_Mean);
    hold on;
    errorbar(1:height(cv_scores), cv_scores.CV_Mean, cv_scores.CV_Std, 'k', 'LineStyle', 'none');
    title('교차검증 평균 정확도', 'FontSize', 12);
    ylabel('CV Accuracy');
    xtickangle(45);
    grid on;
    
    % 성능 레이더 차트
    subplot(2, 4, 8);
    metrics = [results.Accuracy, results.Precision, results.Recall, results.F1_Score];
    plot(metrics', 'o-', 'LineWidth', 1.5);
    title('성능 지표 종합 비교', 'FontSize', 12);
    legend(strrep(results.Model, '_', ' '), 'Location', 'eastoutside', 'FontSize', 8);
    xticks(1:4);
    xticklabels({'Accuracy', 'Precision', 'Recall', 'F1-Score'});
    grid on;
    
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
            colorbar;
            colormap(parula);
            title(strrep(results.Model{i}, '_', ' '));
            xlabel('Predicted');
            ylabel('Actual');
            
            % 혼동 행렬 값 표시
            [r, c] = size(cm);
            for j = 1:r
                for k = 1:c
                    text(k, j, num2str(cm(j,k)), ...
                        'HorizontalAlignment', 'center', ...
                        'Color', 'white', ...
                        'FontWeight', 'bold');
                end
            end
        end
    end
    
    % Figure 3: 특성 중요도
    if CONFIG.calculate_importance && ~isempty(feature_importance)
        figure('Name', '특성 중요도 분석', 'Position', [50, 50, 1200, 700]);
        
        importance_cols = feature_importance.Properties.VariableNames;
        n_importance = length(importance_cols);
        
        for i = 1:n_importance
            subplot(1, n_importance, i);
            
            importance_values = feature_importance.(importance_cols{i});
            [sorted_values, sorted_idx] = sort(importance_values, 'descend');
            
            % 상위 15개 특성 표시
            top_k = min(15, length(sorted_values));
            
            % 특성 타입별 색상 구분
            colors = zeros(top_k, 3);
            for j = 1:top_k
                type = feature_types{sorted_idx(j)};
                switch type
                    case 'linear'
                        colors(j, :) = [0.2, 0.6, 1];  % 파란색
                    case 'nonlinear'
                        colors(j, :) = [1, 0.4, 0.4];  % 빨간색
                    case 'interaction'
                        colors(j, :) = [0.4, 0.8, 0.4];  % 초록색
                    case 'categorical'
                        colors(j, :) = [1, 0.8, 0.2];  % 노란색
                    case 'noise'
                        colors(j, :) = [0.7, 0.7, 0.7];  % 회색
                end
            end
            
            barh(1:top_k, sorted_values(1:top_k), 'FaceColor', 'flat', 'CData', colors);
            
            set(gca, 'YTick', 1:top_k);
            set(gca, 'YTickLabel', feature_names(sorted_idx(1:top_k)));
            xlabel('Importance');
            title(strrep(importance_cols{i}, '_', ' '));
            grid on;
        end
        
        % 범례 추가
        legend({'Linear', 'Nonlinear', 'Interaction', 'Categorical', 'Noise'}, ...
            'Location', 'southoutside', 'Orientation', 'horizontal');
    end
    
    % Figure 4: 성능 대시보드
    figure('Name', '모델 성능 대시보드', 'Position', [50, 50, 1400, 800]);
    
    % 종합 점수 계산 (정규화된 메트릭의 가중 평균)
    overall_score = 0.4 * results.F1_Score + ...
                   0.3 * results.Accuracy + ...
                   0.2 * results.Precision + ...
                   0.1 * results.Recall;
    
    % 효율성 점수 (속도 고려)
    efficiency_score = 1 ./ (1 + results.Train_Time + 10 * results.Pred_Time);
    efficiency_score = efficiency_score / max(efficiency_score);
    
    % 종합 점수 플롯
    subplot(2, 3, 1);
    [sorted_score, sort_idx] = sort(overall_score, 'descend');
    barh(1:length(sorted_score), sorted_score);
    set(gca, 'YTick', 1:length(sorted_score));
    set(gca, 'YTickLabel', results.Model(sort_idx));
    xlabel('종합 성능 점수');
    title('모델 종합 순위');
    grid on;
    
    % 효율성 vs 성능
    subplot(2, 3, 2);
    scatter(results.F1_Score, efficiency_score, 100, 'filled');
    text(results.F1_Score, efficiency_score, results.Model, ...
        'VerticalAlignment', 'bottom', 'HorizontalAlignment', 'center', 'FontSize', 8);
    xlabel('F1-Score');
    ylabel('효율성 점수');
    title('성능 vs 효율성');
    grid on;
    
    % 훈련 시간 vs 예측 시간
    subplot(2, 3, 3);
    scatter(results.Train_Time, results.Pred_Time * 1000, 100, results.F1_Score, 'filled');
    colorbar;
    xlabel('훈련 시간 (초)');
    ylabel('예측 시간 (밀리초)');
    title('시간 효율성 (색상: F1-Score)');
    grid on;
    
    % Top 5 모델 상세 비교
    subplot(2, 3, [4, 5, 6]);
    top5_idx = sort_idx(1:min(5, length(sort_idx)));
    top5_metrics = [results.Accuracy(top5_idx), results.Precision(top5_idx), ...
                   results.Recall(top5_idx), results.F1_Score(top5_idx)];
    
    h = bar(categorical(results.Model(top5_idx)), top5_metrics);
    legend({'Accuracy', 'Precision', 'Recall', 'F1-Score'}, 'Location', 'best');
    title('Top 5 모델 상세 성능');
    ylabel('Score');
    grid on;
    
    fprintf('  시각화 완료!\n\n');
end

fprintf('=====================================\n');
fprintf('        분석 완료!\n');
fprintf('=====================================\n');

%% 보조 함수들

function [accuracy, precision, recall, f1] = evaluateModel(Y_true, Y_pred, problem_type)
    % 모델 성능 평가 함수
    
    if strcmp(problem_type, 'classification')
        % 분류 문제
        if iscategorical(Y_true)
            accuracy = sum(Y_true == Y_pred) / length(Y_true);
            
            % 각 클래스별 성능 계산
            classes = categories(Y_true);
            n_classes = length(classes);
            
            precision = zeros(n_classes, 1);
            recall = zeros(n_classes, 1);
            f1 = zeros(n_classes, 1);
            
            for i = 1:n_classes
                tp = sum((Y_true == classes{i}) & (Y_pred == classes{i}));
                fp = sum((Y_true ~= classes{i}) & (Y_pred == classes{i}));
                fn = sum((Y_true == classes{i}) & (Y_pred ~= classes{i}));
                
                if (tp + fp) > 0
                    precision(i) = tp / (tp + fp);
                else
                    precision(i) = 0;
                end
                
                if (tp + fn) > 0
                    recall(i) = tp / (tp + fn);
                else
                    recall(i) = 0;
                end
                
                if (precision(i) + recall(i)) > 0
                    f1(i) = 2 * precision(i) * recall(i) / (precision(i) + recall(i));
                else
                    f1(i) = 0;
                end
            end
            
            % 매크로 평균
            precision = mean(precision);
            recall = mean(recall);
            f1 = mean(f1);
        else
            accuracy = sum(Y_true == Y_pred) / length(Y_true);
            precision = accuracy;
            recall = accuracy;
            f1 = accuracy;
        end
        
    else
        % 회귀 문제
        mse = mean((Y_true - Y_pred).^2);
        rmse = sqrt(mse);
        mae = mean(abs(Y_true - Y_pred));
        
        % R-squared
        ss_res = sum((Y_true - Y_pred).^2);
        ss_tot = sum((Y_true - mean(Y_true)).^2);
        r2 = 1 - (ss_res / ss_tot);
        
        % 회귀에서는 R2를 accuracy로 사용
        accuracy = max(0, r2);  % 음수 방지
        precision = 1 / (1 + rmse);  % RMSE의 역수를 정규화
        recall = 1 / (1 + mae);      % MAE의 역수를 정규화
        f1 = 2 * precision * recall / (precision + recall + eps);
    end
end

function result_row = createResultRow(model_name, accuracy, precision, recall, f1, train_time, pred_time)
    % 결과 테이블 행 생성
    result_row = table({model_name}, accuracy, precision, recall, f1, train_time, pred_time, ...
        'VariableNames', {'Model', 'Accuracy', 'Precision', 'Recall', 'F1_Score', 'Train_Time', 'Pred_Time'});
end

function model = trainSimpleModel(model_name, X, Y, problem_type)
    % 교차검증용 간단한 모델 훈련
    
    switch model_name
        case 'Decision_Tree'
            if strcmp(problem_type, 'classification')
                model = fitctree(X, Y, 'MaxNumSplits', 20);
            else
                model = fitrtree(X, Y, 'MaxNumSplits', 20);
            end
            
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
            % 기본값: Decision Tree
            if strcmp(problem_type, 'classification')
                model = fitctree(X, Y);
            else
                model = fitrtree(X, Y);
            end
    end
end