%% ==============================================
%  Small-n(≤100) 파이프라인
%  - 로지스틱 회귀(정규화 옵션) + 부트스트랩 CI
%  - (선택) 다중선형회귀 버전도 함께 제공
%  - 성능: ROC-AUC/PR-AUC의 부트스트랩 추정(.632-ish OOB)
%% ==============================================
clear; clc; close all; rng(7);

%% 가상 데이터 (이진 분류)
n = 100; p = 15;
X = randn(n,p); X = normalize(X);
beta = [2*randn(5,1); zeros(p-5,1)];
lin  = X*beta + 0.5*randn(n,1);
pr   = 1./(1+exp(-lin));
Ybin = pr > 0.6;
Y    = categorical(Ybin,[false true],{'Class_0','Class_1'});
feat = "x"+(1:p);

%% --- 1) 기본 로지스틱 회귀 (fitglm) + 부트스트랩 CI(계수)
glm = fitglm(X, Y, 'Distribution','binomial', 'Link','logit', 'LikelihoodPenalty','jeffreys-prior');            % 기본 로지스틱
Bhat = table(glm.Coefficients.Estimate, glm.Coefficients.SE, ...
    'VariableNames',{'Estimate','SE'}, 'RowNames', glm.CoefficientNames);
disp('fitglm 계수 요약'); disp(Bhat);

% 계수의 부트스트랩 신뢰구간 (percentile)
B = 1000;                                       % 부트스트랩 반복
bootfun = @(Xb,Yb) fitglm(Xb,Yb,'Distribution','binomial', 'LikelihoodPenalty','jeffreys-prior').Coefficients.Estimate;
ci_glm = bootci(B, bootfun, X, Y);              % 행: 신뢰구간 하/상, 열: 계수
CI_tbl = table(glm.CoefficientNames, ci_glm(1,:)', ci_glm(2,:)', ...
    'VariableNames',{'Term','CI_low','CI_high'});
disp('fitglm 계수 95% 부트스트랩 CI'); disp(CI_tbl);

%% --- 2) 정규화 로지스틱 (lassoglm) + 교차검증
%  작은 n에서는 과적합 방지에 유리
[B_lasso, FitInfo] = lassoglm(X, double(Y=='Class_1'), 'binomial', 'CV', 5);
idx = FitInfo.IndexMinDeviance;                 % CV로 고른 최적 λ
coef = [FitInfo.Intercept(idx); B_lasso(:,idx)];
nz   = find(coef(2:end)~=0);
fprintf('lassoglm 선택 변수 개수: %d/%d\n', numel(nz), p);

%% --- 3) ROC/PR AUC 부트스트랩 (OOB 평가)
% 각 부트스트랩에서 학습: resample, 평가: OOB(배깅과 유사한 .632 느낌)
B = 500;
rocA = nan(B,1); prA = nan(B,1);
for b = 1:B
    idx = randsample(n, n, true);               % 부트스트랩 인덱스
    oob = setdiff(1:n, unique(idx));            % out-of-bag
    if isempty(oob), oob = randsample(n, round(n*0.3), false); end

    % 규제 로지스틱(안정성): λ는 고정 사용
    [Btmp, FitTmp] = lassoglm(X(idx,:), double(Y(idx)=='Class_1'), 'binomial', ...
                              'Lambda', FitInfo.Lambda(idx));
    coefB = [FitTmp.Intercept; Btmp];
    s = sigmoid([ones(numel(oob),1) X(oob,:)]*coefB); % 1 클래스 확률

    % ROC/PR AUC
    try
        [~,~,~,rocA(b)] = perfcurve(Y(oob), s, 'Class_1');
        [~,~,~,prA(b)]  = perfcurve(Y(oob), s, 'Class_1', 'xCrit','reca','yCrit','prec');
    catch
        rocA(b) = NaN; prA(b) = NaN;
    end
end
fprintf('ROC-AUC (부트): median=%.3f, IQR=%.3f\n', median(rocA,'omitnan'), iqr(rocA,'omitnan'));
fprintf('PR-AUC  (부트): median=%.3f, IQR=%.3f\n', median(prA ,'omitnan'), iqr(prA ,'omitnan'));

figure('Name','Bootstrap AUC Distributions','Position',[80 80 900 380]);
tiledlayout(1,2,'Padding','compact','TileSpacing','compact');
nexttile; histogram(rocA,20); title('ROC-AUC (bootstrap OOB)'); grid on;
nexttile; histogram(prA ,20); title('PR-AUC  (bootstrap OOB)'); grid on;

%% --- 4) 다중선형회귀(연속형 Y) 버전: 부트스트랩 CI
% 예시용 연속형 타깃
Yreg = X*([beta;]+0) + 0.5*randn(n,1);
lm = fitglm(X, Yreg, 'Distribution','normal');          % 다중선형회귀(일반화선형)
ci_lm = bootci(B, @(Xb,Yb) fitglm(Xb,Yb,'Distribution','normal').Coefficients.Estimate, X, Yreg);
fprintf('선형회귀 계수 95%% CI(첫 6개 항)\n'); disp(ci_lm(:,1:6));

%% ---- 보조 ----
function y = sigmoid(a), y = 1./(1+exp(-a)); end



%% ==============================================
%  Small-n(≈100) 추천 경로: Firth(Logistic) + Bootstrap
%% ==============================================
clear; clc; close all; rng(7);

% 예시 데이터 (이진)
n = 100; p = 15;
X = normalize(randn(n,p));
beta = [2*randn(5,1); zeros(p-5,1)];
eta  = X*beta + 0.5*randn(n,1);
pr   = 1./(1+exp(-eta));
Y    = categorical(pr>0.6, [false true], {'Class_0','Class_1'});
feat = "x"+(1:p);

%% 1) Firth(Jeffreys-prior) 로지스틱 회귀
glmF = fitglm(X, Y, 'Distribution','binomial', 'Link','logit', ...
              'LikelihoodPenalty','jeffreys-prior');     % R2024a+  :contentReference[oaicite:1]{index=1}
disp(glmF)

% 계수 부트스트랩 CI
B = 1000;
bootfun = @(Xb,Yb) fitglm(Xb,Yb,'Distribution','binomial','Link','logit', ...
                          'LikelihoodPenalty','jeffreys-prior').Coefficients.Estimate; % :contentReference[oaicite:2]{index=2}
ci = bootci(B, bootfun, X, Y);  % 95% percentile CI  :contentReference[oaicite:3]{index=3}
coefTbl = table(glmF.CoefficientNames, glmF.Coefficients.Estimate, ci(1,:)', ci(2,:)', ...
    'VariableNames',{'Term','Estimate','CI_low','CI_high'});
disp(coefTbl)

% 인샘플 추정치로 대략 AUC
pihat = predict(glmF, X);
[~,~,~,rocA_in] = perfcurve(Y, pihat, 'Class_1');
[~,~,~,prA_in ] = perfcurve(Y, pihat, 'Class_1','xCrit','reca','yCrit','prec'); % :contentReference[oaicite:4]{index=4}
fprintf('In-sample ROC=%.3f, PR=%.3f\n', rocA_in, prA_in);

%% 2) OOB 부트스트랩 성능 분포(권장)
B = 300; rocB = nan(B,1); prB = nan(B,1);
for b = 1:B
    idx = randsample(n,n,true); oob = setdiff(1:n,unique(idx));
    if isempty(oob), oob = randsample(n, round(0.3*n)); end
    M = fitglm(X(idx,:), Y(idx), 'Distribution','binomial','Link','logit', ...
               'LikelihoodPenalty','jeffreys-prior');   % Firth  :contentReference[oaicite:5]{index=5}
    s = predict(M, X(oob,:));
    [~,~,~,rocB(b)] = perfcurve(Y(oob), s, 'Class_1');
    [~,~,~,prB(b) ] = perfcurve(Y(oob), s, 'Class_1','xCrit','reca','yCrit','prec'); % :contentReference[oaicite:6]{index=6}
end
fprintf('Bootstrap OOB ROC median=%.3f (IQR=%.3f)\n', median(rocB,'omitnan'), iqr(rocB,'omitnan'));
fprintf('Bootstrap OOB  PR median=%.3f (IQR=%.3f)\n',  median(prB,'omitnan'),  iqr(prB,'omitnan'));

%% 3) (대안) 변수많음 → lasso 로지스틱
[B_las,Fit] = lassoglm(X, double(Y=='Class_1'),'binomial','CV',5); % :contentReference[oaicite:7]{index=7}
idx = Fit.IndexMinDeviance;
sel = find([Fit.Intercept(idx); B_las(:,idx)](2:end)~=0);
fprintf('LASSO 선택 변수: %d/%d\n', numel(sel), p);
%% --- 4) Firth 로지스틱 회귀(연속형 Y) 버전: 부트스트랩 CI
% 예시용 연속형 타깃
YregFirth = X*([beta;]+0) + 0.5*randn(n,1);
glmF_reg = fitglm(X, YregFirth, 'Distribution','normal');          % Firth 다중선형회귀
ci_lmF = bootci(B, @(Xb,Yb) fitglm(Xb,Yb,'Distribution','normal').Coefficients.Estimate, X, YregFirth);
fprintf('Firth 선형회귀 계수 95%% CI(첫 6개 항)\n'); disp(ci_lmF(:,1:6));

%% --- 5) Lasso 회귀 (lasso) + 교차검증
[B_lasso_full, FitInfo_full] = lasso(X, double(Y=='Class_1'), 'CV', 5);
idx_lasso = FitInfo_full.Index1SE;                 % CV로 고른 최적 λ
coef_lasso = [FitInfo_full.Intercept(idx_lasso); B_lasso_full(:,idx_lasso)];
nz_lasso = find(coef_lasso(2:end)~=0);
fprintf('Lasso 선택 변수 개수: %d/%d\n', numel(nz_lasso), p);

%%
% Ridge 회귀: (A) fitrlinear + K-fold CV로 최적 Lambda 선택 (권장)
% Cleve Moler 스타일로 간결하게 작성

% 데이터: X (n-by-p), Y (분류 레이블), p = size(X,2)
y = double(Y=='Class_1');

lambda = 10.^(-3:0.1:3);                % 후보 λ
K = 10;                                  % K-fold

% 교차검증된 Ridge 회귀 (절편 포함)
CVMdl = fitrlinear(X,y,'Learner','leastsquares', ...
    'Regularization','ridge','Lambda',lambda, ...
    'KFold',K,'Intercept',true);

mse_cv = kfoldLoss(CVMdl);               % 1-by-L (각 λ의 CV-MSE)
[~,idx] = min(mse_cv);
bestLambda = CVMdl.Trained{1}.Lambda(idx);

% 최적 λ로 최종 모델 적합
Mdl = fitrlinear(X,y,'Learner','leastsquares', ...
    'Regularization','ridge','Lambda',bestLambda, ...
    'Intercept',true);

coef_ridge = [Mdl.Bias; Mdl.Beta];       % [절편; 계수]
nz_ridge = find(coef_ridge(2:end)~=0);   % Ridge는 보통 0이 잘 안 나옴
fprintf('Ridge 선택 변수(실제로는 거의 모두 비영): %d/%d, best λ=%.4g\n', ...
    numel(nz_ridge), size(X,2), bestLambda);

% ---------------------------------------------------------------
% Ridge 회귀: (B) ridge + 수동 K-fold CV (원 함수 고수)
% ridge는 출력이 B 하나뿐이고 FitInfo가 없습니다. scaled=0이면 첫 행이 절편입니다.

lambda = 10.^(-3:0.1:3);
K = 10;
cvp = cvpartition(numel(y),'KFold',K);

mse_cv = zeros(numel(lambda),1);
for j = 1:numel(lambda)
    err = zeros(K,1);
    for k = 1:K
        itr = training(cvp,k); ite = test(cvp,k);
        b = ridge(y(itr), X(itr,:), lambda(j), 0);   % (p+1)-by-1, 첫 행 절편
        yhat = b(1) + X(ite,:)*b(2:end);
        err(k) = mean((y(ite) - yhat).^2);
    end
    mse_cv(j) = mean(err);
end
[~,idx] = min(mse_cv);

B = ridge(y, X, lambda, 0);              % (p+1)-by-L
coef_ridge = B(:,idx);
bestLambda = lambda(idx);
nz_ridge = find(coef_ridge(2:end)~=0);
fprintf('Ridge(수동 CV) 변수(거의 모두 비영): %d/%d, best λ=%.4g\n', ...
    numel(nz_ridge), size(X,2), bestLambda);
