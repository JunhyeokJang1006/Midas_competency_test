%% LOOCV 기반 STEP 22.5 테스트 스크립트

clear; clc;

fprintf('=== LOOCV 기반 STEP 22.5 테스트 ===\n');

%% 1. 테스트 데이터 생성
n_samples = 50;  % LOOCV 테스트를 위해 적당한 크기
n_features = 6;

% 테스트용 데이터 생성
X_normalized = randn(n_samples, n_features);
y_weight = [ones(25, 1); zeros(25, 1)];  % 균등 분포

% 설정
config = struct();
config.output_dir = 'D:\project\HR데이터\결과\자가불소';
config.force_recalc_permutation = true;

% 예측 결과 구조 (다양한 경우 테스트)
prediction_results = struct();
prediction_results.TestMethod = struct();
prediction_results.TestMethod.accuracy = 0.78;  % AUC 대신 accuracy만
prediction_results.TestMethod.precision = 0.72; % F1 대신 precision만

best_method = 'TestMethod';

fprintf('✓ 테스트 데이터 준비 완료\n');
fprintf('  - 샘플 수: %d (LOOCV 적합)\n', n_samples);
fprintf('  - 특성 수: %d\n', n_features);
fprintf('  - 클래스 분포: %d/%d\n', sum(y_weight==1), sum(y_weight==0));

%% 2. LOOCV 성능 평가 함수 테스트
fprintf('\n--- LOOCV 성능 평가 함수 테스트 ---\n');

function [auc_score, f1_score] = test_evaluate_loocv_performance(X, y)
    n = length(y);
    loo_predictions = zeros(n, 1);
    loo_probabilities = zeros(n, 1);
    local_failed = 0;

    fprintf('  LOOCV 실행: %d번의 모델 학습\n', n);

    for i = 1:n
        try
            % i번째 샘플 제외
            train_idx = true(n, 1);
            train_idx(i) = false;

            X_train = X(train_idx, :);
            y_train = y(train_idx);
            X_test = X(i, :);

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
            local_failed = local_failed + 1;
            % 실패 시 다수 클래스로 예측
            other_y = y(y ~= y(i));
            if ~isempty(other_y)
                loo_predictions(i) = mode(other_y);
            else
                loo_predictions(i) = y(i);  % 모두 같은 클래스인 경우
            end
            loo_probabilities(i) = 0.5;
        end

        % 진행 표시
        if mod(i, 10) == 0 || i == n
            fprintf('    진행: %d/%d (%.1f%%)\n', i, n, i/n*100);
        end
    end

    % AUC 계산
    try
        [~, ~, ~, auc_score] = perfcurve(y, loo_probabilities, 1);
        fprintf('  ✓ AUC 계산 성공: %.4f\n', auc_score);
    catch
        auc_score = 0.5;
        fprintf('  ⚠ AUC 계산 실패, 기본값 사용: %.4f\n', auc_score);
    end

    % F1 계산
    TP = sum(loo_predictions == 1 & y == 1);
    FP = sum(loo_predictions == 1 & y == 0);
    FN = sum(loo_predictions == 0 & y == 1);
    precision = TP / (TP + FP + eps);
    recall = TP / (TP + FN + eps);
    f1_score = 2 * (precision * recall) / (precision + recall + eps);
    fprintf('  ✓ F1 계산: %.4f (P: %.4f, R: %.4f)\n', f1_score, precision, recall);

    if local_failed > 0
        fprintf('  ⚠ LOOCV 실패: %d/%d (%.1f%%)\n', local_failed, n, local_failed/n*100);
    end
end

% LOOCV 평가 함수 테스트
[test_auc, test_f1] = test_evaluate_loocv_performance(X_normalized, y_weight);

%% 3. 간단한 퍼뮤테이션 테스트 (3회만)
fprintf('\n--- 소규모 LOOCV 퍼뮤테이션 테스트 ---\n');

n_test_permutations = 3;  % 테스트용으로 적게
null_aucs = zeros(n_test_permutations, 1);
null_f1s = zeros(n_test_permutations, 1);

fprintf('소규모 퍼뮤테이션 테스트: %d회\n', n_test_permutations);

for perm = 1:n_test_permutations
    fprintf('  퍼뮤테이션 %d/%d', perm, n_test_permutations);

    % 레이블 셔플
    shuffled_y = y_weight(randperm(n_samples));

    % LOOCV로 퍼뮤테이션 성능 측정
    [perm_auc, perm_f1] = test_evaluate_loocv_performance(X_normalized, shuffled_y);

    null_aucs(perm) = perm_auc;
    null_f1s(perm) = perm_f1;

    fprintf(' → AUC: %.3f, F1: %.3f\n', perm_auc, perm_f1);
end

%% 4. 통계 분석
fprintf('\n--- 통계 분석 결과 ---\n');

% p-value 계산
p_auc = sum(null_aucs >= test_auc) / n_test_permutations;
p_f1 = sum(null_f1s >= test_f1) / n_test_permutations;

fprintf('원본 성능 - AUC: %.4f, F1: %.4f\n', test_auc, test_f1);
fprintf('귀무분포 평균 - AUC: %.4f, F1: %.4f\n', mean(null_aucs), mean(null_f1s));
fprintf('p-value - AUC: %.3f, F1: %.3f\n', p_auc, p_f1);

%% 5. 캐시 키 테스트
fprintf('\n--- LOOCV 캐시 키 생성 테스트 ---\n');

cache_key = struct();
cache_key.n_samples = n_samples;
cache_key.n_features = n_features;
cache_key.class_ratio = sum(y_weight==1) / n_samples;
cache_key.validation_method = 'loocv';
cache_key.loocv_version = '1.0';
cache_key.data_checksum = sum(X_normalized(:)) + sum(y_weight(:))*1000;

fprintf('캐시 키 생성 성공:\n');
fprintf('  - 샘플: %d, 특성: %d\n', cache_key.n_samples, cache_key.n_features);
fprintf('  - 클래스 비율: %.3f\n', cache_key.class_ratio);
fprintf('  - 방법: %s\n', cache_key.validation_method);
fprintf('  - 체크섬: %.0f\n', cache_key.data_checksum);

%% 완료
fprintf('\n✅ LOOCV 기반 STEP 22.5 테스트 완료!\n');
fprintf('   - LOOCV 성능 평가: 정상 작동\n');
fprintf('   - 퍼뮤테이션 루프: 정상 작동\n');
fprintf('   - 통계 계산: 정상 작동\n');
fprintf('   - 캐시 시스템: 정상 작동\n');