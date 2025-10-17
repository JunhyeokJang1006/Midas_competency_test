%% STEP 22.5 디버깅 테스트 스크립트
% 모델 재학습 퍼뮤테이션 테스트 디버깅용

clear; clc;

fprintf('=== STEP 22.5 디버깅 테스트 ===\n');

%% 1. 기본 설정
% 가상 데이터 생성 (테스트용)
n_samples = 100;
n_features = 5;

X_normalized = randn(n_samples, n_features);
y_weight = randi([0, 1], n_samples, 1);

% config 구조체 생성
config = struct();
config.output_dir = 'D:\project\HR데이터\결과\자가불소';
config.force_recalc_permutation = true;  % 강제 재계산

% 예측 결과 가상 생성
prediction_results = struct();
prediction_results.Logistic = struct();
prediction_results.Logistic.auc = 0.75;
prediction_results.Logistic.f1_score = 0.68;

best_method = 'Logistic';

fprintf('✓ 테스트 데이터 생성 완료\n');
fprintf('  - 샘플 수: %d\n', n_samples);
fprintf('  - 특성 수: %d\n', n_features);
fprintf('  - 클래스 분포: %d/%d\n', sum(y_weight==1), sum(y_weight==0));

%% 2. 검증 방법 테스트
validation_methods = {'holdout', 'crossval'};

for vm = 1:length(validation_methods)
    fprintf('\n--- %s 방법 테스트 ---\n', validation_methods{vm});

    % 설정
    validation_method = validation_methods{vm};
    if strcmp(validation_method, 'holdout')
        train_ratio = 0.7;
        fprintf('  훈련/테스트 비율: %.1f/%.1f\n', train_ratio, 1-train_ratio);
    elseif strcmp(validation_method, 'crossval')
        k_fold = 3;  % 빠른 테스트를 위해 3-fold
        fprintf('  K-Fold: %d\n', k_fold);
    end

    %% 3. 단일 퍼뮤테이션 테스트
    fprintf('  단일 퍼뮤테이션 테스트 중...\n');

    try
        % 레이블 셔플
        shuffled_y = y_weight(randperm(length(y_weight)));

        if strcmp(validation_method, 'holdout')
            % Holdout 검증
            n_train = round(length(shuffled_y) * train_ratio);
            train_idx = randperm(length(shuffled_y), n_train);
            test_idx = setdiff(1:length(shuffled_y), train_idx);

            X_train = X_normalized(train_idx, :);
            y_train = shuffled_y(train_idx);
            X_test = X_normalized(test_idx, :);
            y_test = shuffled_y(test_idx);

            % 클래스 분포 확인
            if length(unique(y_train)) < 2
                error('훈련 데이터에 클래스가 부족함');
            end

            % 모델 학습
            mdl_test = fitclinear(X_train, y_train, ...
                'Learner', 'logistic', ...
                'Regularization', 'lasso', ...
                'Lambda', 1e-4, ...
                'Solver', 'sparsa', ...
                'Verbose', 0);

            % 예측
            [pred_labels, pred_scores] = predict(mdl_test, X_test);

            % AUC 계산
            try
                [~, ~, ~, auc_val] = perfcurve(y_test, pred_scores(:,2), 1);
                fprintf('    AUC: %.4f\n', auc_val);
            catch
                fprintf('    AUC 계산 실패\n');
            end

            % F1 계산
            TP = sum(pred_labels == 1 & y_test == 1);
            FP = sum(pred_labels == 1 & y_test == 0);
            FN = sum(pred_labels == 0 & y_test == 1);

            precision = TP / (TP + FP + eps);
            recall = TP / (TP + FN + eps);
            f1_val = 2 * (precision * recall) / (precision + recall + eps);

            fprintf('    F1: %.4f (P: %.4f, R: %.4f)\n', f1_val, precision, recall);

        elseif strcmp(validation_method, 'crossval')
            % Cross-validation
            cv_partition = cvpartition(shuffled_y, 'KFold', k_fold);
            auc_scores = zeros(k_fold, 1);
            f1_scores = zeros(k_fold, 1);

            for fold = 1:k_fold
                train_idx = cv_partition.training(fold);
                test_idx = cv_partition.test(fold);

                X_train = X_normalized(train_idx, :);
                y_train = shuffled_y(train_idx);
                X_test = X_normalized(test_idx, :);
                y_test = shuffled_y(test_idx);

                % 클래스 분포 확인
                if length(unique(y_train)) < 2
                    error('훈련 데이터에 클래스가 부족함 (fold %d)', fold);
                end

                % 모델 학습
                mdl_fold = fitclinear(X_train, y_train, ...
                    'Learner', 'logistic', ...
                    'Regularization', 'lasso', ...
                    'Lambda', 1e-4, ...
                    'Solver', 'sparsa', ...
                    'Verbose', 0);

                % 예측
                [pred_labels, pred_scores] = predict(mdl_fold, X_test);

                % AUC 계산
                try
                    [~, ~, ~, auc_scores(fold)] = perfcurve(y_test, pred_scores(:,2), 1);
                catch
                    auc_scores(fold) = 0.5;
                end

                % F1 계산
                TP = sum(pred_labels == 1 & y_test == 1);
                FP = sum(pred_labels == 1 & y_test == 0);
                FN = sum(pred_labels == 0 & y_test == 1);

                precision = TP / (TP + FP + eps);
                recall = TP / (TP + FN + eps);
                f1_scores(fold) = 2 * (precision * recall) / (precision + recall + eps);

                fprintf('    Fold %d - AUC: %.4f, F1: %.4f\n', fold, auc_scores(fold), f1_scores(fold));
            end

            fprintf('    평균 - AUC: %.4f, F1: %.4f\n', mean(auc_scores), mean(f1_scores));
        end

        fprintf('  ✓ %s 방법 테스트 성공\n', validation_method);

    catch ME
        fprintf('  ✗ %s 방법 테스트 실패: %s\n', validation_method, ME.message);
    end
end

%% 4. 박스플롯 테스트
fprintf('\n--- 박스플롯 테스트 ---\n');

try
    % 가상 분포 생성
    auc_dist = 0.5 + 0.1 * randn(100, 1);
    f1_dist = 0.3 + 0.15 * randn(100, 1);

    % 음수 제거
    auc_dist = max(auc_dist, 0);
    f1_dist = max(f1_dist, 0);

    figure('Position', [100, 100, 800, 400]);

    % 데이터 준비
    auc_data = auc_dist(:);
    f1_data = f1_dist(:);

    data_for_box = [auc_data; f1_data];
    group_labels = [repmat({'AUC'}, length(auc_data), 1); ...
                   repmat({'F1'}, length(f1_data), 1)];

    % 박스플롯 생성
    h_box = boxplot(data_for_box, group_labels);

    % 박스 색상 설정
    box_colors = [0.5 0.5 0.9; 0.9 0.5 0.5];
    h_patch = findobj(gca, 'Tag', 'Box');
    for j = 1:length(h_patch)
        if j <= size(box_colors, 1)
            patch(get(h_patch(j), 'XData'), get(h_patch(j), 'YData'), ...
                  box_colors(j, :), 'FaceAlpha', 0.7);
        end
    end

    title('박스플롯 테스트');
    ylabel('성능 점수');

    fprintf('  ✓ 박스플롯 생성 성공\n');

catch ME
    fprintf('  ✗ 박스플롯 테스트 실패: %s\n', ME.message);
end

%% 5. 완료 메시지
fprintf('\n=== 디버깅 테스트 완료 ===\n');
fprintf('모든 주요 기능이 정상적으로 작동하는지 확인하세요.\n');