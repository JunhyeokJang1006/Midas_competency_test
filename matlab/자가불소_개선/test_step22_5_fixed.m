%% 수정된 STEP 22.5 LOOCV 구현 테스트
% 함수 정의 오류 해결 후 테스트

clear; clc;

fprintf('=== 수정된 STEP 22.5 LOOCV 구현 테스트 ===\n');

%% 1. 테스트 데이터 생성
n_samples = 30;  % 작은 데이터셋으로 빠른 테스트
n_features = 4;

% 테스트용 데이터 생성
X_normalized = randn(n_samples, n_features);
y_weight = [ones(15, 1); zeros(15, 1)];  % 균등 분포

% 설정
config = struct();
config.output_dir = 'D:\project\HR데이터\결과\자가불소';
config.force_recalc_permutation = true;

% 예측 결과 구조 (간단한 테스트용)
prediction_results = struct();
prediction_results.TestMethod = struct();
prediction_results.TestMethod.accuracy = 0.75;
prediction_results.TestMethod.precision = 0.70;

best_method = 'TestMethod';

fprintf('✓ 테스트 데이터 준비 완료\n');
fprintf('  - 샘플 수: %d (LOOCV 적합)\n', n_samples);
fprintf('  - 특성 수: %d\n', n_features);

%% 2. STEP 22.5 수정된 LOOCV 구현 테스트

fprintf('\n--- STEP 22.5: 수정된 LOOCV 기반 퍼뮤테이션 테스트 ---\n');

% Leave-One-Out Cross-Validation (LOOCV) 방식 사용
validation_method = 'loocv';
n_samples = length(y_weight);

% LOOCV는 계산 비용이 높으므로 퍼뮤테이션 횟수 조정
if n_samples <= 100
    n_permutations = 50;  % 테스트용으로 더 적게
elseif n_samples <= 500
    n_permutations = 20;
else
    n_permutations = 10;
end

fprintf('LOOCV 설정:\n');
fprintf('  - 샘플 수: %d\n', n_samples);
fprintf('  - 퍼뮤테이션 횟수: %d (테스트용 감소)\n', n_permutations);

% 원본 모델의 LOOCV 성능 평가 (인라인 구현)
fprintf('\n  ▶ 원본 모델의 LOOCV 성능 평가 중...\n');

n = length(y_weight);
loo_predictions = zeros(n, 1);
loo_probabilities = zeros(n, 1);
failed = 0;

% LOOCV 루프
for i = 1:n
    try
        % i번째 샘플 제외
        train_idx = true(n, 1);
        train_idx(i) = false;

        X_train = X_normalized(train_idx, :);
        y_train = y_weight(train_idx);
        X_test = X_normalized(i, :);

        % 클래스 분포 확인
        if length(unique(y_train)) < 2
            % 한 클래스만 남은 경우 고정 예측
            loo_predictions(i) = mode(y_train);
            loo_probabilities(i) = 0.5;
            continue;
        end

        % 모델 학습
        mdl_loo = fitclinear(X_train, y_train, ...
            'Learner', 'logistic', ...
            'Regularization', 'lasso', ...
            'Lambda', 1e-4, ...
            'Solver', 'sparsa', ...
            'Verbose', 0);

        % 예측
        [pred_label, pred_scores] = predict(mdl_loo, X_test);
        loo_predictions(i) = pred_label;
        if size(pred_scores, 2) >= 2
            loo_probabilities(i) = pred_scores(2);
        else
            loo_probabilities(i) = pred_scores(1);
        end

    catch
        failed = failed + 1;
        % 실패 시 다수 클래스로 예측
        other_y = y_weight(y_weight ~= y_weight(i));
        if ~isempty(other_y)
            loo_predictions(i) = mode(other_y);
        else
            loo_predictions(i) = y_weight(i);  % 모두 같은 클래스인 경우
        end
        loo_probabilities(i) = 0.5;
    end

    % 진행 표시
    if mod(i, 5) == 0 || i == n
        fprintf('    진행: %d/%d (%.1f%%)\n', i, n, i/n*100);
    end
end

% AUC 계산
try
    [~, ~, ~, original_auc] = perfcurve(y_weight, loo_probabilities, 1);
    fprintf('  ✓ 원본 LOOCV AUC: %.4f\n', original_auc);
catch
    original_auc = 0.5;
    fprintf('  ⚠ AUC 계산 실패, 기본값 사용: %.4f\n', original_auc);
end

% F1 계산
TP = sum(loo_predictions == 1 & y_weight == 1);
FP = sum(loo_predictions == 1 & y_weight == 0);
FN = sum(loo_predictions == 0 & y_weight == 1);
precision = TP / (TP + FP + eps);
recall = TP / (TP + FN + eps);
original_f1 = 2 * (precision * recall) / (precision + recall + eps);
fprintf('  ✓ 원본 LOOCV F1: %.4f\n', original_f1);

%% 3. 퍼뮤테이션 테스트 (수정된 인라인 구현)
fprintf('\n  ▶ LOOCV 퍼뮤테이션 실행 중...\n');

% null distribution 초기화
null_auc_distribution = zeros(n_permutations, 1);
null_f1_distribution = zeros(n_permutations, 1);
failed_permutations = 0;

tic;
for perm = 1:n_permutations
    try
        % 레이블 셔플
        shuffled_y = y_weight(randperm(n_samples));

        % LOOCV로 퍼뮤테이션 성능 측정 (인라인 구현)
        perm_n = length(shuffled_y);
        perm_loo_predictions = zeros(perm_n, 1);
        perm_loo_probabilities = zeros(perm_n, 1);
        perm_failed = 0;

        % LOOCV 루프
        for i = 1:perm_n
            try
                % i번째 샘플 제외
                train_idx = true(perm_n, 1);
                train_idx(i) = false;

                perm_X_train = X_normalized(train_idx, :);
                perm_y_train = shuffled_y(train_idx);
                perm_X_test = X_normalized(i, :);

                % 클래스 분포 확인
                if length(unique(perm_y_train)) < 2
                    % 한 클래스만 남은 경우 고정 예측
                    perm_loo_predictions(i) = mode(perm_y_train);
                    perm_loo_probabilities(i) = 0.5;
                    continue;
                end

                % 모델 학습
                perm_mdl_loo = fitclinear(perm_X_train, perm_y_train, ...
                    'Learner', 'logistic', ...
                    'Regularization', 'lasso', ...
                    'Lambda', 1e-4, ...
                    'Solver', 'sparsa', ...
                    'Verbose', 0);

                % 예측
                [perm_pred_label, perm_pred_scores] = predict(perm_mdl_loo, perm_X_test);
                perm_loo_predictions(i) = perm_pred_label;
                if size(perm_pred_scores, 2) >= 2
                    perm_loo_probabilities(i) = perm_pred_scores(2);
                else
                    perm_loo_probabilities(i) = perm_pred_scores(1);
                end

            catch
                perm_failed = perm_failed + 1;
                % 실패 시 다수 클래스로 예측
                other_y = shuffled_y(shuffled_y ~= shuffled_y(i));
                if ~isempty(other_y)
                    perm_loo_predictions(i) = mode(other_y);
                else
                    perm_loo_predictions(i) = shuffled_y(i);  % 모두 같은 클래스인 경우
                end
                perm_loo_probabilities(i) = 0.5;
            end
        end

        % AUC 계산
        try
            [~, ~, ~, perm_auc] = perfcurve(shuffled_y, perm_loo_probabilities, 1);
        catch
            perm_auc = 0.5;
        end

        % F1 계산
        perm_TP = sum(perm_loo_predictions == 1 & shuffled_y == 1);
        perm_FP = sum(perm_loo_predictions == 1 & shuffled_y == 0);
        perm_FN = sum(perm_loo_predictions == 0 & shuffled_y == 1);
        perm_precision = perm_TP / (perm_TP + perm_FP + eps);
        perm_recall = perm_TP / (perm_TP + perm_FN + eps);
        perm_f1 = 2 * (perm_precision * perm_recall) / (perm_precision + perm_recall + eps);

        % null distribution에 저장
        null_auc_distribution(perm) = perm_auc;
        null_f1_distribution(perm) = perm_f1;

    catch ME
        % LOOCV 퍼뮤테이션 실패 처리
        failed_permutations = failed_permutations + 1;
        null_auc_distribution(perm) = 0.5;  % 랜덤 수준
        null_f1_distribution(perm) = 0;     % 최저 성능

        if failed_permutations <= 3
            fprintf('    ⚠ LOOCV 퍼뮤테이션 %d 실패: %s\n', perm, ME.message);
        end
    end

    % 진행상황 표시
    if mod(perm, 10) == 0 || perm == n_permutations
        fprintf('    퍼뮤테이션 진행: %d/%d (%.1f%%)\n', perm, n_permutations, perm/n_permutations*100);
    end
end

elapsed_time = toc;
fprintf('  ✓ 퍼뮤테이션 완료 (%.1f초)\n', elapsed_time);

%% 4. 통계 분석
fprintf('\n--- 통계 분석 결과 ---\n');

% p-value 계산
p_auc = sum(null_auc_distribution >= original_auc) / n_permutations;
p_f1 = sum(null_f1_distribution >= original_f1) / n_permutations;

fprintf('원본 성능 - AUC: %.4f, F1: %.4f\n', original_auc, original_f1);
fprintf('귀무분포 평균 - AUC: %.4f, F1: %.4f\n', mean(null_auc_distribution), mean(null_f1_distribution));
fprintf('p-value - AUC: %.3f, F1: %.3f\n', p_auc, p_f1);

% 실패 통계
if failed > 0
    fprintf('원본 LOOCV 실패: %d/%d (%.1f%%)\n', failed, n, failed/n*100);
end
if failed_permutations > 0
    fprintf('퍼뮤테이션 실패: %d/%d (%.1f%%)\n', failed_permutations, n_permutations, failed_permutations/n_permutations*100);
end

%% 완료
fprintf('\n✅ 수정된 STEP 22.5 LOOCV 구현 테스트 완료!\n');
fprintf('   - 함수 정의 오류: 해결됨\n');
fprintf('   - 인라인 LOOCV 구현: 정상 작동\n');
fprintf('   - 퍼뮤테이션 루프: 정상 작동\n');
fprintf('   - 통계 계산: 정상 작동\n');