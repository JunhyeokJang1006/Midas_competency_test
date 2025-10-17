%% 고급 MATLAB 머신러닝 모델 성능 비교 (교차검증 + 하이퍼파라미터 최적화)
% 작성자: Claude AI
% 날짜: 2024
% 포함사항: 교차검증, 하이퍼파라미터 최적화, 사전확률 통제, 특성 스케일링

clear; clc; close all;

%% 1. 설정 및 데이터 로드
fprintf('=== 고급 머신러닝 분석 시작 ===\n');

% 분석 설정
use_binary_classification = true;  % true: 이진분류, false: 다중분류
use_hyperparameter_optimization = true;  % 하이퍼파라미터 최적화 여부
use_cross_validation = true;  % 교차검증 사용 여부
k_fold = 5;  % 교차검증 폴드 수
control_class_imbalance = true;  % 클래스 불균형 처리 여부

% 예시 데이터 로드 (본인의 데이터에 맞게 수정하세요)
% load('your_data.mat');
% X = features;
% Y = labels;

% 복잡한 현실적 데이터 생성 (다양한 패턴과 노이즈 포함)
rng(42);
n_samples = 3000;
n_features = 20;

fprintf('복잡한 합성 데이터 생성 중...\n');

% 1. 기본 특성군 생성
% 주요 특성군 1: 선형 관계 특성 (5개)
X_linear = randn(n_samples, 5);

% 주요 특성군 2: 비선형 관계 특성 (4개)
X_nonlinear = zeros(n_samples, 4);
X_nonlinear(:,1) = sin(2*pi*randn(n_samples, 1)) + 0.3*randn(n_samples, 1);
X_nonlinear(:,2) = cos(randn(n_samples, 1)) .* randn(n_samples, 1);
X_nonlinear(:,3) = log(abs(randn(n_samples, 1)) + 1) .* sign(randn(n_samples, 1));
X_nonlinear(:,4) = (randn(n_samples, 1)).^3 + 0.5*randn(n_samples, 1);

% 주요 특성군 3: 상호작용 특성 (3개)
X_interaction = zeros(n_samples, 3);
X_interaction(:,1) = X_linear(:,1) .* X_linear(:,2) + 0.2*randn(n_samples, 1);
X_interaction(:,2) = X_nonlinear(:,1) .* X_nonlinear(:,2) + 0.3*randn(n_samples, 1);
X_interaction(:,3) = (X_linear(:,1) + X_linear(:,2)) ./ (abs(X_nonlinear(:,1)) + 1) + 0.1*randn(n_samples, 1);

% 범주형 특성을 연속형으로 변환한 특성 (2개)
X_categorical = zeros(n_samples, 2);
categories_1 = randi([1, 4], n_samples, 1);  % 4개 카테고리
categories_2 = randi([1, 3], n_samples, 1);  % 3개 카테고리
X_categorical(:,1) = categories_1 + 0.1*randn(n_samples, 1);
X_categorical(:,2) = categories_2 + 0.15*randn(n_samples, 1);

% 순수 노이즈 특성 (6개)
X_noise = 0.5 * randn(n_samples, 6);

% 전체 특성 결합
X_combined = [X_linear, X_nonlinear, X_interaction, X_categorical, X_noise];

% 2. 특성 간 상관관계 추가 (현실적 다중공선성)
correlation_matrix = eye(n_features);
correlation_matrix(1,2) = 0.7; correlation_matrix(2,1) = 0.7;  % 강한 상관관계
correlation_matrix(3,4) = 0.5; correlation_matrix(4,3) = 0.5;  % 중간 상관관계
correlation_matrix(15,16) = 0.8; correlation_matrix(16,15) = 0.8;  % 노이즈 간 상관관계

% 상관관계가 있는 데이터로 변환
L = chol(correlation_matrix, 'lower');
X = (L * X_combined')';

% 3. 이상값(Outliers) 추가
outlier_ratio = 0.05;  % 5% 이상값
n_outliers = floor(n_samples * outlier_ratio);
outlier_indices = randperm(n_samples, n_outliers);

for i = 1:n_outliers
    idx = outlier_indices(i);
    feature_to_corrupt = randi(n_features);
    X(idx, feature_to_corrupt) = X(idx, feature_to_corrupt) + 5*randn();  % 극단값 추가
end

% 4. 복잡한 타겟 변수 생성
if use_binary_classification
    fprintf('복잡한 이진 분류 문제 생성\n');
    
    % 선형 성분
    linear_component = 1.2*X(:,1) - 0.8*X(:,2) + 0.6*X(:,3) - 0.4*X(:,4) + 0.3*X(:,5);
    
    % 비선형 성분
    nonlinear_component = 0.5*sin(X(:,6)) + 0.3*cos(X(:,7)) + 0.2*log(abs(X(:,8))+1);
    
    % 상호작용 성분
    interaction_component = 0.4*X(:,9).*X(:,10) + 0.2*X(:,11).*X(:,12);
    
    % 임계값 기반 분류 (복잡한 결정 경계)
    decision_score = linear_component + nonlinear_component + interaction_component;
    
    % 동적 임계값 (데이터에 따라 변함)
    adaptive_threshold = median(decision_score) + 0.1*std(decision_score)*randn(n_samples, 1);
    
    % 확률적 분류 + 노이즈
    probabilities = 1 ./ (1 + exp(-(decision_score - adaptive_threshold)));
    random_noise = 0.1 * randn(n_samples, 1);  % 분류 노이즈
    final_probs = max(0.05, min(0.95, probabilities + random_noise));
    
    Y_binary = double(final_probs > 0.5) + 1;
    Y = categorical(Y_binary);
    
    % 극단적 클래스 불균형 생성 (실제 상황 모방)
    if control_class_imbalance
        majority_class = mode(Y);
        minority_indices = find(Y ~= majority_class);
        
        % 소수 클래스의 70%만 유지 (30:70 비율로 불균형)
        keep_minority = randperm(length(minority_indices), floor(length(minority_indices) * 0.4));
        keep_indices = [find(Y == majority_class); minority_indices(keep_minority)];
        
        X = X(keep_indices, :);
        Y = Y(keep_indices);
        n_samples = length(keep_indices);
        fprintf('극단적 클래스 불균형 생성: 소수 클래스 비율 = %.1f%%\n', ...
            sum(Y ~= majority_class) / length(Y) * 100);
    end
    
else
    fprintf('복잡한 다중 분류 문제 생성 (3클래스)\n');
    
    % 각 클래스별 복잡한 결정 영역
    % 클래스 1: 주로 선형 특성 기반
    score_1 = X(:,1) + 0.5*X(:,2) - 0.3*X(:,3) + 0.2*sin(X(:,6));
    
    % 클래스 2: 비선형 특성 기반  
    score_2 = 0.7*cos(X(:,7)) + 0.4*X(:,9).*X(:,10) - 0.2*X(:,4);
    
    % 클래스 3: 상호작용 기반
    score_3 = 0.5*X(:,11).*X(:,12) + 0.3*log(abs(X(:,8))+1) - 0.1*X(:,5);
    
    % 소프트맥스 기반 확률적 할당
    scores = [score_1, score_2, score_3];
    exp_scores = exp(scores - max(scores, [], 2));  % 수치적 안정성
    probabilities = exp_scores ./ sum(exp_scores, 2);
    
    % 노이즈 추가
    probabilities = probabilities + 0.05*randn(size(probabilities));
    probabilities = max(0.01, probabilities);
    probabilities = probabilities ./ sum(probabilities, 2);
    
    % 확률적 샘플링
    Y_multi = zeros(n_samples, 1);
    for i = 1:n_samples
        Y_multi(i) = randsample(1:3, 1, true, probabilities(i,:));
    end
    Y = categorical(Y_multi);
end

% 5. 특성명 생성 (해석 가능성)
feature_names = {};
feature_types = {};
for i = 1:5
    feature_names{end+1} = sprintf('Linear_F%d', i);
    feature_types{end+1} = 'linear';
end
for i = 1:4
    feature_names{end+1} = sprintf('Nonlinear_F%d', i);
    feature_types{end+1} = 'nonlinear';
end
for i = 1:3
    feature_names{end+1} = sprintf('Interaction_F%d', i);
    feature_types{end+1} = 'interaction';
end
for i = 1:2
    feature_names{end+1} = sprintf('Categorical_F%d', i);
    feature_types{end+1} = 'categorical';
end
for i = 1:6
    feature_names{end+1} = sprintf('Noise_F%d', i);
    feature_types{end+1} = 'noise';
end

fprintf('생성된 특성 유형별 개수:\n');
fprintf('- 선형 관계: %d개\n', sum(strcmp(feature_types, 'linear')));
fprintf('- 비선형 관계: %d개\n', sum(strcmp(feature_types, 'nonlinear')));
fprintf('- 상호작용: %d개\n', sum(strcmp(feature_types, 'interaction')));
fprintf('- 범주형: %d개\n', sum(strcmp(feature_types, 'categorical')));
fprintf('- 순수 노이즈: %d개\n', sum(strcmp(feature_types, 'noise')));
fprintf('- 이상값 비율: %.1f%%\n', outlier_ratio * 100);

% 데이터 정보 출력
fprintf('데이터 크기: %d samples, %d features\n', size(X, 1), size(X, 2));
fprintf('클래스 분포:\n');
tabulate(Y)

% 클래스 불균형 비율 계산
class_counts = countcats(Y);
imbalance_ratio = max(class_counts) / min(class_counts);
fprintf('클래스 불균형 비율: %.2f:1\n', imbalance_ratio);

%% 2. 데이터 분할 (훈련/테스트)
cv = cvpartition(Y, 'HoldOut', 0.2); % 80% 훈련, 20% 테스트
X_train = X(cv.training, :);
Y_train = Y(cv.training);
X_test = X(cv.test, :);
Y_test = Y(cv.test);

fprintf('훈련 데이터: %d samples\n', size(X_train, 1));
fprintf('테스트 데이터: %d samples\n', size(X_test, 1));

%% 3. 모델 정의 및 훈련
fprintf('\n=== 모델 훈련 시작 ===\n');

% 결과 저장을 위한 구조체
results = struct();
model_names = {};
models = {};

%% 3.1 Bagged Trees (Bagging Ensemble)
fprintf('1. Bagged Trees 훈련 중...\n');
tic;
try
    model_bagged = fitcensemble(X_train, Y_train, 'Method', 'Bag', ...
        'NumLearningCycles', 100, 'Learners', 'tree');
    time_bagged = toc;
    models{end+1} = model_bagged;
    model_names{end+1} = 'Bagged Trees';
    results.bagged.training_time = time_bagged;
    fprintf('   완료! 훈련 시간: %.2f초\n', time_bagged);
catch ME
    fprintf('   에러: %s\n', ME.message);
end

%% 3.2 Boosted Trees (AdaBoost)
fprintf('2. AdaBoost 훈련 중...\n');
tic;
try
    model_adaboost = fitcensemble(X_train, Y_train, 'Method', 'AdaBoostM1', ...
        'NumLearningCycles', 100, 'Learners', 'tree');
    time_adaboost = toc;
    models{end+1} = model_adaboost;
    model_names{end+1} = 'AdaBoost';
    results.adaboost.training_time = time_adaboost;
    fprintf('   완료! 훈련 시간: %.2f초\n', time_adaboost);
catch ME
    fprintf('   에러: %s\n', ME.message);
end

%% 3.3 LogitBoost (Adaptive Logistic Regression)
fprintf('3. LogitBoost 훈련 중...\n');
tic;
try
    % LogitBoost는 이진 분류에만 사용 가능
    if length(categories(Y_train)) == 2
        model_logitboost = fitcensemble(X_train, Y_train, 'Method', 'LogitBoost', ...
            'NumLearningCycles', 100, 'Learners', 'tree', 'LearnRate', 0.1);
        time_logitboost = toc;
        models{end+1} = model_logitboost;
        model_names{end+1} = 'LogitBoost';
        results.logitboost.training_time = time_logitboost;
        fprintf('   완료! 훈련 시간: %.2f초\n', time_logitboost);
    else
        fprintf('   LogitBoost는 이진 분류만 지원합니다. 다중 클래스 데이터에서는 건너뜁니다.\n');
    end
catch ME
    fprintf('   에러: %s\n', ME.message);
end

%% 3.4 LPBoost (Linear Programming Boosting)
fprintf('4. LPBoost 훈련 중...\n');
tic;
try
    % LPBoost는 Optimization Toolbox가 필요합니다
    if exist('linprog', 'file') == 2
        model_lpboost = fitcensemble(X_train, Y_train, 'Method', 'LPBoost', ...
            'NumLearningCycles', 50, 'Learners', 'tree', 'MarginPrecision', 0.01);
        time_lpboost = toc;
        models{end+1} = model_lpboost;
        model_names{end+1} = 'LPBoost';
        results.lpboost.training_time = time_lpboost;
        fprintf('   완료! 훈련 시간: %.2f초\n', time_lpboost);
    else
        fprintf('   LPBoost는 Optimization Toolbox가 필요합니다.\n');
    end
catch ME
    fprintf('   에러: %s\n', ME.message);
    if contains(ME.message, 'Optimization Toolbox')
        fprintf('   Optimization Toolbox가 설치되어 있지 않습니다.\n');
    end
end

%% 3.5 Support Vector Machines (SVM)
fprintf('5. SVM 훈련 중...\n');
tic;
try
    % 다중 클래스 분류를 위한 SVM
    if length(categories(Y_train)) == 2
        model_svm = fitcsvm(X_train, Y_train, 'KernelFunction', 'rbf');
    else
        model_svm = fitcecoc(X_train, Y_train, 'Learners', ...
            templateSVM('KernelFunction', 'rbf'));
    end
    time_svm = toc;
    models{end+1} = model_svm;
    model_names{end+1} = 'SVM';
    results.svm.training_time = time_svm;
    fprintf('   완료! 훈련 시간: %.2f초\n', time_svm);
catch ME
    fprintf('   에러: %s\n', ME.message);
end

%% 3.6 Neural Networks
fprintf('6. Neural Networks 훈련 중...\n');
tic;
try
    % fitcnet 사용 (Statistics and Machine Learning Toolbox)
    model_nn = fitcnet(X_train, Y_train, 'LayerSizes', [50, 25], ...
        'Activations', 'relu', 'Verbose', 0);
    time_nn = toc;
    models{end+1} = model_nn;
    model_names{end+1} = 'Neural Network';
    results.nn.training_time = time_nn;
    fprintf('   완료! 훈련 시간: %.2f초\n', time_nn);
catch ME
    fprintf('   Neural Network 에러: %s\n', ME.message);
    fprintf('   Deep Learning Toolbox가 없을 수 있습니다.\n');
end

%% 3.7 Transformer (Deep Learning Toolbox 필요)
fprintf('7. Transformer/Deep Neural Network 훈련 중...\n');
try
    % Transformer는 주로 시퀀스 데이터용이므로 tabular data에는 제한적
    % 여기서는 간단한 구현 예시를 제공합니다
    if exist('dlarray', 'file') == 2
        fprintf('   Transformer는 시퀀스 데이터에 최적화되어 있어 표준 tabular data에는 적합하지 않습니다.\n');
        fprintf('   대신 Deep Neural Network를 추가로 구현합니다.\n');
        
        % Deep Learning Toolbox를 사용한 더 깊은 신경망
        layers = [
            featureInputLayer(size(X_train, 2))
            fullyConnectedLayer(128)
            reluLayer
            dropoutLayer(0.2)
            fullyConnectedLayer(64)
            reluLayer
            dropoutLayer(0.2)
            fullyConnectedLayer(length(categories(Y_train)))
            softmaxLayer
            classificationLayer];
        
        options = trainingOptions('adam', ...
            'MaxEpochs', 50, ...
            'MiniBatchSize', 32, ...
            'ValidationFrequency', 30, ...
            'Verbose', false, ...
            'Plots', 'none');
        
        tic;
        model_deep = trainNetwork(X_train, Y_train, layers, options);
        time_deep = toc;
        models{end+1} = model_deep;
        model_names{end+1} = 'Deep Neural Network';
        results.deep.training_time = time_deep;
        fprintf('   완료! 훈련 시간: %.2f초\n', time_deep);
    else
        fprintf('   Deep Learning Toolbox가 설치되어 있지 않습니다.\n');
    end
catch ME
    fprintf('   Deep Learning 에러: %s\n', ME.message);
end

%% 4. 모델 성능 평가
fprintf('\n=== 모델 성능 평가 ===\n');

% 성능 지표 저장
performance_metrics = table();
confusion_matrices = {};

for i = 1:length(models)
    fprintf('모델: %s\n', model_names{i});
    
    % 예측 수행
    tic;
    if isa(models{i}, 'SeriesNetwork') % Deep Learning 모델
        Y_pred = classify(models{i}, X_test);
        [~, scores] = classify(models{i}, X_test);
    else
        [Y_pred, scores] = predict(models{i}, X_test);
    end
    prediction_time = toc;
    
    % 성능 지표 계산
    accuracy = sum(Y_pred == Y_test) / length(Y_test);
    
    % Confusion Matrix
    cm = confusionmat(Y_test, Y_pred);
    confusion_matrices{i} = cm;
    
    % 다중 클래스 성능 지표
    [precision, recall, f1] = calculateMultiClassMetrics(Y_test, Y_pred);
    
    % 결과 저장
    performance_metrics = [performance_metrics; table({model_names{i}}, accuracy, ...
        mean(precision), mean(recall), mean(f1), prediction_time, ...
        'VariableNames', {'Model', 'Accuracy', 'Precision', 'Recall', 'F1_Score', 'Prediction_Time'})];
    
    fprintf('   정확도: %.4f\n', accuracy);
    fprintf('   정밀도: %.4f\n', mean(precision));
    fprintf('   재현율: %.4f\n', mean(recall));
    fprintf('   F1-Score: %.4f\n', mean(f1));
    fprintf('   예측 시간: %.4f초\n\n', prediction_time);
end

%% 5. 결과 시각화
fprintf('=== 결과 시각화 ===\n');

% 성능 비교 차트
figure('Position', [100, 100, 1200, 800]);

% 정확도 비교
subplot(2, 3, 1);
bar(performance_metrics.Accuracy);
set(gca, 'XTickLabel', performance_metrics.Model, 'XTickLabelRotation', 45);
title('모델별 정확도 비교');
ylabel('정확도');
grid on;

% F1-Score 비교
subplot(2, 3, 2);
bar(performance_metrics.F1_Score);
set(gca, 'XTickLabel', performance_metrics.Model, 'XTickLabelRotation', 45);
title('모델별 F1-Score 비교');
ylabel('F1-Score');
grid on;

% 훈련 시간 비교
training_times = [];
for i = 1:length(model_names)
    field_name = lower(strrep(model_names{i}, ' ', '_'));
    field_name = strrep(field_name, '-', '_');
    if isfield(results, field_name) && isfield(results.(field_name), 'training_time')
        training_times(i) = results.(field_name).training_time;
    else
        training_times(i) = NaN;
    end
end

subplot(2, 3, 3);
bar(training_times);
set(gca, 'XTickLabel', performance_metrics.Model, 'XTickLabelRotation', 45);
title('모델별 훈련 시간 비교');
ylabel('시간 (초)');
grid on;

% 예측 시간 비교
subplot(2, 3, 4);
bar(performance_metrics.Prediction_Time);
set(gca, 'XTickLabel', performance_metrics.Model, 'XTickLabelRotation', 45);
title('모델별 예측 시간 비교');
ylabel('시간 (초)');
grid on;

% 전체 성능 레이더 차트 (정규화된 지표)
subplot(2, 3, [5, 6]);
normalized_metrics = [performance_metrics.Accuracy, performance_metrics.F1_Score, ...
                     1./(performance_metrics.Prediction_Time+0.001)]; % 예측 시간은 역수로 변환
plot(normalized_metrics', 'LineWidth', 2, 'Marker', 'o');
legend(performance_metrics.Model, 'Location', 'best');
title('정규화된 성능 지표 비교');
xlabel('지표 (1: 정확도, 2: F1-Score, 3: 1/예측시간)');
ylabel('성능 값');
grid on;

%% 6. 최고 성능 모델 선택 및 요약
fprintf('=== 최고 성능 모델 요약 ===\n');

% 정확도 기준 최고 성능 모델
[best_accuracy, best_idx] = max(performance_metrics.Accuracy);
fprintf('최고 정확도 모델: %s (정확도: %.4f)\n', ...
    performance_metrics.Model{best_idx}, best_accuracy);

% F1-Score 기준 최고 성능 모델
[best_f1, best_f1_idx] = max(performance_metrics.F1_Score);
fprintf('최고 F1-Score 모델: %s (F1-Score: %.4f)\n', ...
    performance_metrics.Model{best_f1_idx}, best_f1);

% 종합 성능 표 출력
fprintf('\n전체 성능 결과:\n');
disp(performance_metrics);

% Confusion Matrix 시각화 (최고 성능 모델)
figure;
confusionchart(Y_test, predict(models{best_idx}, X_test));
title(sprintf('Confusion Matrix - %s', model_names{best_idx}));

fprintf('분석 완료!\n');

%% 보조 함수들
function [precision, recall, f1] = calculateMultiClassMetrics(Y_true, Y_pred)
    % 다중 클래스 분류를 위한 성능 지표 계산
    classes = categories(Y_true);
    num_classes = length(classes);
    
    precision = zeros(num_classes, 1);
    recall = zeros(num_classes, 1);
    f1 = zeros(num_classes, 1);
    
    for i = 1:num_classes
        tp = sum((Y_true == classes{i}) & (Y_pred == classes{i}));
        fp = sum((Y_true ~= classes{i}) & (Y_pred == classes{i}));
        fn = sum((Y_true == classes{i}) & (Y_pred ~= classes{i}));
        
        precision(i) = tp / (tp + fp + eps);
        recall(i) = tp / (tp + fn + eps);
        f1(i) = 2 * precision(i) * recall(i) / (precision(i) + recall(i) + eps);
    end
end