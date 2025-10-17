%% ================================================================
% 역량검사 → 성과 예측: LogitBoost / AdaBoost / Bagging / NN
% - 성능 평가: Accuracy, Precision, Recall, F1, MCC, ROC‑AUC, PR‑AUC
% - 해석: predictorImportance (Tree Ensembles) + SHAP (전체 모델)
% 작성: 2025, Cleve Moler 스타일, 최신 기능 활용
%% ================================================================

clear; clc; close all;
rng(42);

%% 0) 데이터 준비 (실데이터 교체 가능)
use_synthetic = true;

if ~use_synthetic
    % 실제 데이터 예시:
    % T = innerjoin(itemTbl, perfTbl, 'Keys', 'ID');
    % featureVars = T.Properties.VariableNames;
    % featureVars = featureVars(contains(featureVars, 'Q') | contains(featureVars, 'Subscale'));
    % X = T{:, featureVars};
    % Y = categorical(T.Promoted > 0, [false true], {'Class_0','Class_1'});
else
    n = 1200; p = 30;
    X = randn(n,p);
    w = [3*randn(10,1); zeros(p-10,1)];
    s = X(:,1:10)*w(1:10) + 0.5*sin(X(:,3)) + 0.3*X(:,5).^2 + 0.5*randn(n,1);
    pr = 1./(1+exp(-s));
    Ybin = pr + 0.05*randn(n,1) > 0.65;
    Y = categorical(Ybin, [false true], {'Class_0','Class_1'});
    featureVars = "x" + (1:p);
end

X = normalize(X);

cvHO = cvpartition(Y,'HoldOut',0.2);
Xtr = X(training(cvHO),:); Ytr = Y(training(cvHO));
Xte = X(test(cvHO),:);     Yte = Y(test(cvHO));

cv5 = cvpartition(Ytr,'KFold',5);
opt = struct('CVPartition',cv5, 'ShowPlots',false, 'Verbose',0);

%% 1) 모델 학습 및 하이퍼파라미터 최적화
t = templateTree(Reproducible=true);
fprintf("모델 학습 및 튜닝...\n");

methods = {'LogitBoost', 'AdaBoostM1', 'Bag'};  % Bag은 Random Forest 계열:contentReference[oaicite:1]{index=1}
models = cell(size(methods));
names  = methods;

for i = 1:numel(methods)
    mdl = fitcensemble(Xtr, Ytr, ...
        Method=methods{i}, Learners=t, ...
        OptimizeHyperparameters="auto", ...
        HyperparameterOptimizationOptions=opt);
    models{i} = mdl;
end

% Neural Net (fitcnet)
mdl_nn = fitcnet(Xtr, Ytr, Standardize=true, ...
    OptimizeHyperparameters="auto", ...
    HyperparameterOptimizationOptions=opt);
models{end+1} = mdl_nn;
names{end+1} = 'NeuralNet';

%% 2) 테스트 성능 측정
fprintf("테스트 세트 성능 평가...\n");

metrics = ["Accuracy","Precision","Recall","F1","MCC","ROC_AUC","PR_AUC","PredTime"];
R = array2table(nan(numel(models), numel(metrics)), ...
    'VariableNames', metrics);
R.Properties.RowNames = names;

yhat_all = cell(size(models));
score_all = cell(size(models));

for i = 1:numel(models)
    mdl = models{i};
    tic;
    [yhat, score] = predict(mdl, Xte);
    t_elapsed = toc;
    
    yhat_all{i} = yhat;
    score_all{i} = score;
    
    [acc, pre, rec, f1] = evaluateModel(Yte, yhat);
    mcc = computeMCC(Yte, yhat);
    [roc_auc, pr_auc] = computeAUCs(Yte, score, mdl);
    
    R{i, :} = [acc, pre, rec, f1, mcc, roc_auc, pr_auc, t_elapsed];
end

disp(R);

%% 3) 트리 기반 모델 특성 중요도
hasNames = numel(featureVars)==size(X,2);
if ~hasNames
    featureVars = "x" + (1:size(X,2));
end

impTbl = table(featureVars(:), 'VariableNames', {'Feature'});
for i = 1:min(3, numel(methods))
    mdl = models{i};
    if isa(mdl, 'ClassificationEnsemble')
        try
            imp = predictorImportance(mdl);  % tree 기반 중요도:contentReference[oaicite:2]{index=2}
            impTbl.(names{i}) = imp(:);
        catch
            impTbl.(names{i}) = nan(size(impTbl,1),1);
        end
    end
end

% 시각화
firstTreeModel = names{ find(ismember(names, methods), 1) };
[~, ix] = sort(impTbl.(firstTreeModel), 'descend');
topk = ix(1:min(15, numel(ix)));
figure('Name','Tree Model Feature Importance','Position',[100 100 800 400]);
barh(categorical(impTbl.Feature(topk)), impTbl.(firstTreeModel)(topk));
title(['Feature Importance — ' firstTreeModel]); xlabel('Importance'); grid on;

%% 4) SHAP 전 모델 적용
nShap = min(200, size(Xtr,1));
Xshap = Xtr(1:nShap,:);

fprintf("SHAP 계산 중...\n");
for i = 1:numel(models)
    mdl = models{i};
    try
        explainer = shapley(mdl, array2table(Xtr,'VariableNames',cellstr(featureVars)), ...
            NumObservationsToSample=100, Method="tree");  % 최신 옵션:contentReference[oaicite:3]{index=3}
        explainer = fit(explainer, array2table(Xshap,'VariableNames',cellstr(featureVars)), UseParallel=true);
        
        mas = explainer.MeanAbsoluteShapley;  % 전역 중요도
        [~, ix2] = sort(mas.Value, 'descend');
        k = min(15, height(mas));
        
        figure('Name',['SHAP — ' names{i}],'Position',[50+300*i 50 600 400]);
        barh(categorical(mas.Predictor(ix2(1:k))), mas.Value(ix2(1:k)));
        title(['SHAP Global Importance — ' names{i}]); xlabel('Mean |SHAP|'); grid on;
    catch ME
        warning("SHAP 실패 (%s): %s", names{i}, ME.message);
    end
end

%% 5) 혼동 행렬 배열
figure('Name','Confusion Matrices','Position',[50 50 1200 800]);
tiledlayout(2,2,'Padding','compact','TileSpacing','compact');
for i = 1:numel(models)
    nexttile;
    cm = confusionmat(Yte, yhat_all{i});
    imagesc(cm); axis image; colorbar;
    title(names{i}); xlabel('Predicted'); ylabel('Actual');
    textColors = 'w';
    for r = 1:size(cm,1)
        for c = 1:size(cm,2)
            text(c, r, num2str(cm(r,c)), 'Color', textColors, 'FontWeight','bold', 'HorizontalAlignment','center');
        end
    end
end

%% 6) 성능 비교 플롯
figure('Name','Model Performance Comparison','Position',[60 60 1200 500]);
subplot(1,2,1);
bar(categorical(names), [R.Accuracy R.F1 R.MCC]);
legend('Accuracy','F1','MCC','Location','best'); title('Accuracy / F1 / MCC'); grid on; xtickangle(20);

subplot(1,2,2);
bar(categorical(names), [R.ROC_AUC R.PR_AUC R.PredTime]);
legend('ROC‑AUC','PR‑AUC','PredTime','Location','best'); title('AUC / Prediction Time'); grid on; xtickangle(20);

[~, bestIdx] = max(R.F1);
fprintf("완료! 최고 모델 (F1 기준): %s\n", names{bestIdx});

%%  보조 함수
function [acc, pre, rec, f1] = evaluateModel(Y_true, Y_pred)
    acc = mean(Y_true == Y_pred);
    C = confusionmat(Y_true, Y_pred);
    P = C(2,2) / (C(2,2) + C(1,2) + eps);
    R = C(2,2) / (C(2,2) + C(2,1) + eps);
    pre = P; rec = R;
    f1 = 2*P*R/(P+R+eps);
end

function mcc = computeMCC(Y_true, Y_pred)
    C = confusionmat(Y_true, Y_pred);
    if numel(C)==4
        TP = C(2,2); TN = C(1,1); FP = C(1,2); FN = C(2,1);
        mcc = (TP*TN - FP*FN)/sqrt((TP+FP)*(TP+FN)*(TN+FP)*(TN+FN) + eps);
    else
        mcc = NaN;
    end
end

function [rocAUC, prAUC] = computeAUCs(Y_true, scores, M)
    posClass = M.ClassNames{2};
    try
        [~,~,~,rocAUC] = perfcurve(Y_true, scores(:,2), posClass);
        [~,~,~,prAUC]  = perfcurve(Y_true, scores(:,2), posClass, 'xCrit','reca', 'yCrit','prec');
    catch
        rocAUC = NaN; prAUC = NaN;
    end
end



%% 데이터 (X: 역검 하위요인, Y: 성과 종합점수)레
% X: n×p 행렬, Y: n×1 벡터
nBoot = 1000;                      % 부트스트랩 반복 수
lambda = 1;                        % 릿지 규제 계수(튜닝 가능)

B_boot = zeros(size(X,2), nBoot);

parfor b = 1:nBoot
    idx = randsample(size(X,1), size(X,1), true);   % 부트스트랩 샘플
    Xb = X(idx,:); Yb = Y(idx);

    % 릿지 회귀 계수 추정
    % ridge는 상수항 포함, 두 번째 입력 Y, 세 번째 lambda
    B = ridge(Yb, Xb, lambda, 0);   % 0=상수항 분리
    B_boot(:,b) = B(2:end);         % 첫 번째는 절편 → 제외
end

% 평균 기여도, 신뢰구간
B_mean = mean(B_boot,2);
B_CI   = prctile(B_boot,[2.5 97.5],2);

% 결과 표
result = table(string(featureVars(:)), B_mean, B_CI(:,1), B_CI(:,2), ...
    'VariableNames',{'Feature','MeanCoef','CI_lower','CI_upper'})

% 시각화
figure;
errorbar(1:numel(featureVars), B_mean, ...
    B_mean-B_CI(:,1), B_CI(:,2)-B_mean, 'o');
xticks(1:numel(featureVars)); xticklabels(featureVars); xtickangle(45);
ylabel('Ridge Coefficient (with 95% CI)');
title('하위 요인 기여도 (Ridge + Bootstrap)');
grid on;
