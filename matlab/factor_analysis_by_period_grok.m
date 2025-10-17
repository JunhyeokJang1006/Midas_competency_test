%% ===============================================================
%  역량검사 → 성과 예측 검증: LogitBoost / AdaBoost / Bagging / NN
%  - 성능: Acc/Prec/Rec/F1/MCC/ROC-AUC/PR-AUC
%  - 해석: 트리계열 predictorImportance + 전 모델 SHAP
%  - 작성: 2025 (Cleve Moler 스타일, 현대 문법)
%  - 개선: 불균형 데이터 처리 강화 (RUSBoost 선택적 지원), 에러 처리 추가, 코드 가독성 향상, SHAP 계산 효율화
%% ===============================================================

clear; clc; close all;
rng(42);  % 재현성을 위한 시드 설정

%% 0) 데이터 어댑터: 실데이터(X,Y) 꽂기
%  - itemTbl: [ID, Q1..Qp] 또는 [ID, Subscale1..K]
%  - perfTbl: [ID, Promoted(0/1) 또는 등급/레이블, Contribution_Average ...]
%  아래는 예시: 실데이터가 있다면 이 블록만 교체하세요.
use_synthetic = true;   % 실데이터 준비되면 false로 두고 아래 주석 참고

if ~use_synthetic
    % 예) 실데이터 예시
    % T = innerjoin(itemTbl, perfTbl, 'Keys','ID');
    % featureVars = T.Properties.VariableNames(contains(T.Properties.VariableNames, {'Q','Subscale'}));
    % X = T{:, featureVars};
    % Y = categorical(T.Promoted > 0, [false true], {'Class_0','Class_1'});
else
    % === 합성 예시 데이터 (이진 분류, 불균형) ===
    n = 1200; p = 30;
    X = randn(n, p);
    w = [3*randn(10,1); zeros(p-10,1)];
    s = X(:,1:10)*w(1:10) + 0.5*sin(X(:,3)) + 0.3*X(:,5).^2 + 0.5*randn(n,1);
    pr = 1./(1+exp(-s));
    Ybin = pr + 0.05*randn(n,1) > 0.65;  % 양성 적은 불균형 생성
    Y = categorical(Ybin, [false true], {'Class_0','Class_1'});
    featureVars = "x" + (1:p);
end

% 전처리: 표준화 (정규화)
X = normalize(X);

% 학습/평가 분할 (계층화 Holdout)
cvHO = cvpartition(Y, 'HoldOut', 0.2);
Xtr = X(training(cvHO), :); Ytr = Y(training(cvHO));
Xte = X(test(cvHO), :);     Yte = Y(test(cvHO));

% K-fold 교차검증 구성 (튜닝 시 사용)
cv = cvpartition(Ytr, 'KFold', 5);
opt = struct('CVPartition', cv, 'ShowPlots', false, 'Verbose', 0);

%% 1) 네 가지 모델 훈련 (튜닝 포함)
t = templateTree('Reproducible', true);  % 재현성을 위한 트리 템플릿

fprintf("훈련/튜닝 중...\n");

% 1) LogitBoost
mdl_logit = fitcensemble(Xtr, Ytr, 'Method', 'LogitBoost', 'Learners', t, ...
                         'OptimizeHyperparameters', 'auto', ...
                         'HyperparameterOptimizationOptions', opt);

% 2) AdaBoost (이진 분류의 경우 AdaBoostM1, 불균형 시 RUSBoost 선택적 지원)
adaMethod = 'AdaBoostM1';  % 불균형이 심하면 'RUSBoost'로 변경 가능
mdl_ada = fitcensemble(Xtr, Ytr, 'Method', adaMethod, 'Learners', t, ...
                       'OptimizeHyperparameters', 'auto', ...
                       'HyperparameterOptimizationOptions', opt);

% 3) Bagging (Random Forest 계열)
mdl_bag = fitcensemble(Xtr, Ytr, 'Method', 'Bag', 'Learners', t, ...
                       'OptimizeHyperparameters', 'auto', ...
                       'HyperparameterOptimizationOptions', opt);

% 4) Neural Network
mdl_nn = fitcnet(Xtr, Ytr, 'Standardize', true, ...
                 'OptimizeHyperparameters', 'auto', ...
                 'HyperparameterOptimizationOptions', opt);

models = {mdl_logit, mdl_ada, mdl_bag, mdl_nn};
names  = {'LogitBoost', 'AdaBoost', 'Bagging', 'NeuralNet'};

%% 2) 테스트 세트 성능 평가
fprintf("테스트 평가 중...\n");
nModels = numel(names);
varNames = {'Model', 'Accuracy', 'Precision', 'Recall', 'F1', 'MCC', 'ROC_AUC', 'PR_AUC', 'Pred_Time'};
varTypes = {'string', 'double', 'double', 'double', 'double', 'double', 'double', 'double', 'double'};
R = table('Size', [nModels numel(varNames)], ...
          'VariableTypes', varTypes, ...
          'VariableNames', varNames);

% 모델명 채우기
R.Model = string(names(:));

scores_cache = cell(nModels, 1);
yhat_cache   = cell(nModels, 1);

for i = 1:nModels
    M = models{i};
    tic;
    try
        [yhat, score] = predict(M, Xte);
    catch ME
        warning('%s 예측 실패: %s', names{i}, ME.message);
        yhat = categorical(NaN(size(Yte, 1), 1));
        score = NaN(size(Yte, 1), numel(M.ClassNames));
        elapsed = NaN;
    end
    elapsed = toc;

    yhat_cache{i}   = yhat;
    scores_cache{i} = score;

    [acc, pre, rec, f1] = evaluateModel(Yte, yhat, 'classification');
    mcc = computeMCC(Yte, yhat);
    [auc, ap] = computeAUCs(Yte, score, M);
    R{i, {'Accuracy', 'Precision', 'Recall', 'F1', 'MCC', 'ROC_AUC', 'PR_AUC', 'Pred_Time'}} = ...
        [acc, pre, rec, f1, mcc, auc, ap, elapsed];
end

disp(R);

%% 3) 특성 중요도 (트리 앙상블 전용: predictorImportance)
hasNames = exist('featureVars', 'var') && numel(featureVars) == size(X, 2);
if ~hasNames
    featureVars = "x" + (1:size(X, 2));
end
featureVars = string(featureVars(:));

impTbl = table(featureVars, 'VariableNames', {'Feature'});
firstImpCol = [];

for i = 1:3  % 1:LogitBoost, 2:AdaBoost, 3:Bagging (트리 기반 모델만)
    colName = matlab.lang.makeValidName(names{i});
    if isa(models{i}, 'ClassificationEnsemble')
        try
            imp = predictorImportance(models{i});
            impTbl.(colName) = imp(:);
            if isempty(firstImpCol)
                firstImpCol = colName;
            end
        catch ME
            warning('%s importance 실패: %s', names{i}, ME.message);
            impTbl.(colName) = NaN(height(impTbl), 1);
        end
    else
        impTbl.(colName) = NaN(height(impTbl), 1);
    end
end

% 시각화 (가장 먼저 성공한 모델 기준)
if ~isempty(firstImpCol)
    fprintf('\n[트리계열 Feature Importance 상위 15]\n');
    [~, ix] = sort(impTbl.(firstImpCol), 'descend', 'MissingPlacement', 'last');
    topk = ix(1:min(15, numel(ix)));
    figure('Name', 'Tree Ensembles - Predictor Importance', 'Position', [100 100 900 500]);
    barh(categorical(impTbl.Feature(topk)), impTbl.(firstImpCol)(topk));
    grid on;
    title(sprintf('%s Predictor Importance (Top-15)', firstImpCol));
end

%% 4) SHAP: 전 모델 공통 (NN 포함)
% 비용을 줄이려면 일부 샘플만 사용
nShap = min(200, size(Xtr, 1));
Xshap = Xtr(1:nShap, :);

fprintf('\nSHAP 계산 중(전 모델)...\n');
for i = 1:nModels
    try
        expl = shapley(models{i}, Xshap);  % 전 모델 지원
        sv = mean(abs(expl.ShapleyValues), 1);  % 전역 중요도
        [svs, idx] = sort(sv, 'descend');
        k = min(15, numel(idx));
        figure('Name', ['SHAP - ' names{i}], 'Position', [50 + 250*i 50 650 500]);
        barh(svs(1:k));
        set(gca, 'YTickLabel', featureVars(idx(1:k)));
        title(['SHAP Global Importance (Top-' num2str(k) ') - ' names{i}]);
        grid on;
    catch ME
        warning('%s: SHAP 계산 실패: %s', names{i}, ME.message);
    end
end

%% 5) 혼동행렬 대시보드
figure('Name', 'Confusion Matrices', 'Position', [50 50 1200 800]);
tiledlayout(2, 2, 'Padding', 'compact', 'TileSpacing', 'compact');
for i = 1:nModels
    nexttile;
    cm = confusionmat(Yte, yhat_cache{i});
    imagesc(cm); axis image; colorbar; title(names{i});
    xlabel('Predicted'); ylabel('Actual'); colormap(parula);
    for r = 1:size(cm, 1)
        for c = 1:size(cm, 2)
            text(c, r, num2str(cm(r, c)), 'Color', 'w', 'FontWeight', 'bold', ...
                 'HorizontalAlignment', 'center');
        end
    end
end

%% 6) 성능 비교 플롯
figure('Name', 'Model Performance', 'Position', [60 60 1200 550]);
subplot(1, 2, 1);
bar(categorical(R.Model), [R.Accuracy R.F1 R.MCC]);
legend({'Acc', 'F1', 'MCC'}, 'Location', 'best');
title('정확도/ F1/ MCC'); grid on; xtickangle(20);
subplot(1, 2, 2);
bar(categorical(R.Model), [R.ROC_AUC R.PR_AUC R.Pred_Time]);
legend({'ROC-AUC', 'PR-AUC', 'Pred Time(s)'}, 'Location', 'best');
title('AUC/시간'); grid on; xtickangle(20);

fprintf('\n완료! 최고 모델(기본 F1 기준): %s\n', R.Model{argmax(R.F1)});

%% ==================== 보조 함수 ====================
function [accuracy, precision, recall, f1] = evaluateModel(Y_true, Y_pred, problem_type)
    if strcmp(problem_type, 'classification')
        accuracy = mean(Y_true == Y_pred);
        classes = categories(Y_true); k = numel(classes);
        P = zeros(k, 1); R = zeros(k, 1); F = zeros(k, 1);
        for j = 1:k
            tp = sum(Y_true == classes{j} & Y_pred == classes{j});
            fp = sum(Y_true ~= classes{j} & Y_pred == classes{j});
            fn = sum(Y_true == classes{j} & Y_pred ~= classes{j});
            P(j) = tp / (tp + fp + eps);
            R(j) = tp / (tp + fn + eps);
            F(j) = 2 * P(j) * R(j) / (P(j) + R(j) + eps);
        end
        precision = mean(P); recall = mean(R); f1 = mean(F);
    else
        error('Only classification is implemented here.');
    end
end

function mcc = computeMCC(Y_true, Y_pred)
    cm = confusionmat(Y_true, Y_pred);
    if isequal(size(cm), [2, 2])
        TP = cm(2, 2); TN = cm(1, 1); FP = cm(1, 2); FN = cm(2, 1);
        denom = sqrt((TP + FP) * (TP + FN) * (TN + FP) * (TN + FN));
        mcc = ((TP * TN) - (FP * FN)) / max(denom, eps);
    else
        mcc = NaN;  % 다중분류 확장은 별도
    end
end

function [roc_auc, pr_auc] = computeAUCs(Y_true, score, M)
    % score의 양성 열 찾기
    classes = categories(Y_true);
    if size(score, 2) == 2
        posIdx = 2;
    else
        % 평균 최대값 열
        [~, posIdx] = max(mean(score, 1));
    end
    try
        [~, ~, ~, roc_auc] = perfcurve(Y_true, score(:, posIdx), M.ClassNames(posIdx));
        [~, ~, ~, pr_auc] = perfcurve(Y_true, score(:, posIdx), M.ClassNames(posIdx), 'xCrit', 'reca', 'yCrit', 'prec');
    catch
        roc_auc = NaN; pr_auc = NaN;
    end
end

function i = argmax(x)
    [~, i] = max(x);
end