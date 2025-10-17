%% STEP 22.6: 모든 가중치 방법 퍼뮤테이션 테스트 비교 (수정된 버전)
fprintf('\n\n【STEP 22.6】 모든 가중치 방법 퍼뮤테이션 테스트 비교 (수정됨)\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 필요한 변수들이 정의되어 있는지 확인
if ~exist('prediction_results', 'var')
    fprintf('⚠ prediction_results 변수가 정의되지 않았습니다. STEP 22.6을 건너뜁니다.\n');
    return;
end

if ~exist('n_permutations', 'var')
    % 기본값 설정
    if exist('n_samples', 'var')
        if n_samples <= 100
            n_permutations = 1000;   % 작은 데이터용으로 축소
        elseif n_samples <= 500
            n_permutations = 500;
        else
            n_permutations = 300;
        end
    else
        n_permutations = 500;  % 기본값
    end
    fprintf('⚠ n_permutations가 정의되지 않아 기본값 %d로 설정합니다.\n', n_permutations);
end

% 5가지 방법 정의
all_methods = {'Correlation', 'Logistic', 'Bootstrap', 'Ttest', 'Ensemble'};
n_all_methods = length(all_methods);

% 사용 가능한 방법들 확인
available_methods = {};
for i = 1:n_all_methods
    method_name = all_methods{i};
    if isfield(prediction_results, method_name)
        if isfield(prediction_results.(method_name), 'AUC') && ...
           isfield(prediction_results.(method_name), 'f1_score')
            available_methods{end+1} = method_name;
        else
            fprintf('⚠ %s 방법에 AUC 또는 f1_score가 없습니다.\n', method_name);
        end
    else
        fprintf('⚠ %s 방법의 데이터가 prediction_results에 없습니다.\n', method_name);
    end
end

n_available_methods = length(available_methods);

if n_available_methods == 0
    fprintf('⚠ 사용 가능한 방법이 없습니다. STEP 22.6을 건너뜁니다.\n');
    return;
end

fprintf('\n📊 사용 가능한 방법: %d개 (%s)\n', n_available_methods, strjoin(available_methods, ', '));

% 예상 시간 계산 및 사용자 알림
estimated_time_per_method = n_permutations * 0.05; % 더 보수적인 추정
total_estimated_time = estimated_time_per_method * n_available_methods;
fprintf('  - 퍼뮤테이션 수: %d회 (방법당)\n', n_permutations);
fprintf('  - 예상 소요시간: %.1f분 (총 %.1f시간)\n', total_estimated_time/60, total_estimated_time/3600);

% 사용자 확인
response = input('\n계속 진행하시겠습니까? (y/n): ', 's');
if ~strcmpi(response, 'y')
    fprintf('⚠ 사용자가 취소했습니다. STEP 22.6을 건너뜁니다.\n');
    fprintf('\n✅ STEP 22.6 건너뛰기 완료\n');
    return;
end

% 모든 방법의 퍼뮤테이션 결과 저장소
all_permutation_results = struct();

% 병렬 처리 준비
try
    if isempty(gcp('nocreate'))
        fprintf('\n  ▶ 병렬 풀 시작 중...\n');
        parpool('local');
    end
    use_parallel = true;
    fprintf('  ✓ 병렬 처리 사용 (워커 수: %d)\n', gcp().NumWorkers);
catch
    use_parallel = false;
    fprintf('  ⚠ 병렬 처리 사용 불가, 순차 처리로 진행\n');
end

% 필요한 변수들 확인 및 기본값 설정
if ~exist('X_normalized', 'var') || ~exist('y_weight', 'var')
    fprintf('⚠ X_normalized 또는 y_weight 변수가 없습니다. STEP 22.6을 건너뜁니다.\n');
    return;
end

if ~exist('config', 'var')
    config = struct();
    config.output_dir = pwd;  % 현재 디렉토리를 기본값으로 설정
    config.force_recalc_permutation = false;
end

n_samples = size(X_normalized, 1);

% 각 방법별 퍼뮤테이션 테스트 수행 (사용 가능한 방법들만)
for method_idx = 1:n_available_methods
    current_method = available_methods{method_idx};

    fprintf('\n【방법 %d/%d: %s 퍼뮤테이션 테스트】\n', method_idx, n_available_methods, current_method);
    fprintf('────────────────────────────────────────────\n');

    % 캐시 파일 경로
    method_cache_file = fullfile(config.output_dir, sprintf('model_permutation_cache_%s.mat', current_method));

    % 캐시 키 생성 (방법명 포함)
    cache_key = struct();
    cache_key.method = current_method;
    cache_key.n_samples = n_samples;
    cache_key.n_features = size(X_normalized, 2);
    cache_key.class_ratio = sum(y_weight==1) / n_samples;
    cache_key.version = '2.1';  % 수정된 버전
    cache_key.data_checksum = sum(X_normalized(:)) + sum(y_weight(:))*1000;

    % 캐시 확인
    use_cached = false;
    if exist(method_cache_file, 'file') && ~config.force_recalc_permutation
        try
            load(method_cache_file, 'method_perm_cache');
            if isequal(method_perm_cache.cache_key, cache_key)
                fprintf('  ✓ 캐시 유효: 기존 결과 사용 (%s)\n', current_method);
                all_permutation_results.(current_method) = method_perm_cache.results;
                use_cached = true;
            else
                fprintf('  ⚠ 캐시 무효: 새로 계산 필요\n');
            end
        catch
            fprintf('  ⚠ 캐시 로드 실패: 새로 계산\n');
        end
    end

    if ~use_cached
        % 퍼뮤테이션 테스트 실행
        fprintf('  ▶ %s 방법 퍼뮤테이션 실행 중...\n', current_method);

        % 관찰된 성능 가져오기
        observed_auc = prediction_results.(current_method).AUC;
        observed_f1 = prediction_results.(current_method).f1_score;

        fprintf('    관찰값 - AUC: %.4f, F1: %.4f\n', observed_auc, observed_f1);

        % null distribution 초기화
        null_auc_distribution = zeros(n_permutations, 1);
        null_f1_distribution = zeros(n_permutations, 1);
        failed_permutations = 0;

        tic;

        % 진행률 표시를 위한 설정
        fprintf('    진행률: 0%% ');
        update_interval = max(1, ceil(n_permutations / 20));

        for perm = 1:n_permutations
            try
                % 레이블 셔플
                shuffled_y = y_weight(randperm(n_samples));
                perm_n = length(shuffled_y);

                % 방법별 퍼뮤테이션 처리
                switch current_method
                    case 'Correlation'
                        % 상관계수 기반 가중치
                        perm_weights = zeros(size(X_normalized, 2), 1);
                        for j = 1:size(X_normalized, 2)
                            try
                                r = corr(X_normalized(:, j), shuffled_y, 'rows', 'complete');
                                if ~isnan(r), perm_weights(j) = abs(r); end
                            catch, end
                        end
                        perm_weights = perm_weights / (sum(perm_weights) + eps);

                        perm_weighted_scores = X_normalized * perm_weights;
                        perm_threshold = median(perm_weighted_scores);
                        perm_predictions = double(perm_weighted_scores > perm_threshold);
                        perm_probabilities = (perm_weighted_scores - min(perm_weighted_scores)) / ...
                                           (max(perm_weighted_scores) - min(perm_weighted_scores) + eps);

                    case 'Ttest'
                        % t-test 효과크기 기반 가중치
                        perm_weights = zeros(size(X_normalized, 2), 1);
                        high_idx = shuffled_y == 1;
                        low_idx = shuffled_y == 0;

                        for j = 1:size(X_normalized, 2)
                            try
                                high_scores = X_normalized(high_idx, j);
                                low_scores = X_normalized(low_idx, j);
                                if length(high_scores) > 1 && length(low_scores) > 1
                                    pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                                                     (length(low_scores)-1)*var(low_scores)) / ...
                                                    (length(high_scores) + length(low_scores) - 2));
                                    if pooled_std > 0
                                        cohens_d = abs(mean(high_scores) - mean(low_scores)) / pooled_std;
                                        perm_weights(j) = cohens_d;
                                    end
                                end
                            catch, end
                        end
                        perm_weights = perm_weights / (sum(perm_weights) + eps);

                        perm_weighted_scores = X_normalized * perm_weights;
                        perm_threshold = median(perm_weighted_scores);
                        perm_predictions = double(perm_weighted_scores > perm_threshold);
                        perm_probabilities = (perm_weighted_scores - min(perm_weighted_scores)) / ...
                                           (max(perm_weighted_scores) - min(perm_weighted_scores) + eps);

                    case 'Bootstrap'
                        % 간단한 부트스트랩
                        perm_weights = zeros(size(X_normalized, 2), 1);
                        n_boot_simple = 30; % 더 축약된 버전

                        for boot = 1:n_boot_simple
                            boot_idx = randsample(perm_n, perm_n, true);
                            boot_X = X_normalized(boot_idx, :);
                            boot_y = shuffled_y(boot_idx);

                            for j = 1:size(boot_X, 2)
                                try
                                    r = corr(boot_X(:, j), boot_y, 'rows', 'complete');
                                    if ~isnan(r), perm_weights(j) = perm_weights(j) + abs(r); end
                                catch, end
                            end
                        end
                        perm_weights = perm_weights / n_boot_simple;
                        perm_weights = perm_weights / (sum(perm_weights) + eps);

                        perm_weighted_scores = X_normalized * perm_weights;
                        perm_threshold = median(perm_weighted_scores);
                        perm_predictions = double(perm_weighted_scores > perm_threshold);
                        perm_probabilities = (perm_weighted_scores - min(perm_weighted_scores)) / ...
                                           (max(perm_weighted_scores) - min(perm_weighted_scores) + eps);

                    case 'Logistic'
                        % 로지스틱 회귀 (간소화된 버전)
                        try
                            % 전체 데이터로 한 번에 훈련
                            if length(unique(shuffled_y)) >= 2
                                perm_mdl = fitclinear(X_normalized, shuffled_y, ...
                                    'Learner', 'logistic', 'Regularization', 'lasso', ...
                                    'Lambda', 1e-3, 'Solver', 'sparsa', 'Verbose', 0);

                                [perm_predictions, pred_scores] = predict(perm_mdl, X_normalized);
                                if size(pred_scores, 2) >= 2
                                    perm_probabilities = pred_scores(:, 2);
                                else
                                    perm_probabilities = pred_scores(:, 1);
                                end
                            else
                                % 단일 클래스인 경우 랜덤 예측
                                perm_predictions = double(rand(perm_n, 1) > 0.5);
                                perm_probabilities = rand(perm_n, 1);
                            end
                        catch
                            % 오류 발생 시 랜덤 예측
                            perm_predictions = double(rand(perm_n, 1) > 0.5);
                            perm_probabilities = rand(perm_n, 1);
                        end

                    case 'Ensemble'
                        % 앙상블 방법
                        corr_weights_perm = zeros(size(X_normalized, 2), 1);
                        for j = 1:size(X_normalized, 2)
                            try
                                r = corr(X_normalized(:, j), shuffled_y, 'rows', 'complete');
                                if ~isnan(r), corr_weights_perm(j) = abs(r); end
                            catch, end
                        end
                        corr_weights_perm = corr_weights_perm / (sum(corr_weights_perm) + eps);

                        ttest_weights_perm = zeros(size(X_normalized, 2), 1);
                        high_idx = shuffled_y == 1;
                        low_idx = shuffled_y == 0;
                        for j = 1:size(X_normalized, 2)
                            try
                                high_scores = X_normalized(high_idx, j);
                                low_scores = X_normalized(low_idx, j);
                                if length(high_scores) > 1 && length(low_scores) > 1
                                    cohens_d = abs(mean(high_scores) - mean(low_scores)) / ...
                                              (sqrt((var(high_scores) + var(low_scores))/2) + eps);
                                    ttest_weights_perm(j) = cohens_d;
                                end
                            catch, end
                        end
                        ttest_weights_perm = ttest_weights_perm / (sum(ttest_weights_perm) + eps);

                        ensemble_weights_perm = (corr_weights_perm + ttest_weights_perm) / 2;

                        perm_weighted_scores = X_normalized * ensemble_weights_perm;
                        perm_threshold = median(perm_weighted_scores);
                        perm_predictions = double(perm_weighted_scores > perm_threshold);
                        perm_probabilities = (perm_weighted_scores - min(perm_weighted_scores)) / ...
                                           (max(perm_weighted_scores) - min(perm_weighted_scores) + eps);

                    otherwise
                        perm_predictions = double(rand(perm_n, 1) > 0.5);
                        perm_probabilities = rand(perm_n, 1);
                end

                % AUC 계산
                try
                    [~, ~, ~, perm_auc] = perfcurve(shuffled_y, perm_probabilities, 1);
                    if isnan(perm_auc), perm_auc = 0.5; end
                catch
                    perm_auc = 0.5;
                end

                % F1 계산
                perm_TP = sum(perm_predictions == 1 & shuffled_y == 1);
                perm_FP = sum(perm_predictions == 1 & shuffled_y == 0);
                perm_FN = sum(perm_predictions == 0 & shuffled_y == 1);
                perm_precision = perm_TP / (perm_TP + perm_FP + eps);
                perm_recall = perm_TP / (perm_TP + perm_FN + eps);
                perm_f1 = 2 * (perm_precision * perm_recall) / (perm_precision + perm_recall + eps);

                null_auc_distribution(perm) = perm_auc;
                null_f1_distribution(perm) = perm_f1;

            catch
                null_auc_distribution(perm) = 0.5;
                null_f1_distribution(perm) = 0;
                failed_permutations = failed_permutations + 1;
            end

            % 진행률 표시
            if mod(perm, update_interval) == 0 || perm == n_permutations
                progress = 100 * perm / n_permutations;
                fprintf('\b\b\b\b%3.0f%% ', progress);
            end
        end

        fprintf('\n');
        elapsed_time = toc;

        % 통계 계산
        mean_null_auc = mean(null_auc_distribution);
        std_null_auc = std(null_auc_distribution);
        p_value_auc = sum(null_auc_distribution >= observed_auc) / n_permutations;
        z_score_auc = (observed_auc - mean_null_auc) / (std_null_auc + eps);

        mean_null_f1 = mean(null_f1_distribution);
        std_null_f1 = std(null_f1_distribution);
        p_value_f1 = sum(null_f1_distribution >= observed_f1) / n_permutations;
        z_score_f1 = (observed_f1 - mean_null_f1) / (std_null_f1 + eps);

        % 결과 저장
        method_results = struct();
        method_results.auc = struct();
        method_results.auc.observed = observed_auc;
        method_results.auc.p_value = p_value_auc;
        method_results.auc.mean_null = mean_null_auc;
        method_results.auc.std_null = std_null_auc;
        method_results.auc.z_score = z_score_auc;
        method_results.auc.null_distribution = null_auc_distribution;

        method_results.f1 = struct();
        method_results.f1.observed = observed_f1;
        method_results.f1.p_value = p_value_f1;
        method_results.f1.mean_null = mean_null_f1;
        method_results.f1.std_null = std_null_f1;
        method_results.f1.z_score = z_score_f1;
        method_results.f1.null_distribution = null_f1_distribution;

        method_results.meta = struct();
        method_results.meta.n_permutations = n_permutations;
        method_results.meta.failed_permutations = failed_permutations;
        method_results.meta.elapsed_time = elapsed_time;

        all_permutation_results.(current_method) = method_results;

        % 캐시 저장
        method_perm_cache = struct();
        method_perm_cache.cache_key = cache_key;
        method_perm_cache.results = method_results;
        method_perm_cache.timestamp = datestr(now);

        try
            save(method_cache_file, 'method_perm_cache');
            fprintf('  ✓ 캐시 저장: %s\n', method_cache_file);
        catch
            fprintf('  ⚠ 캐시 저장 실패\n');
        end

        % 결과 요약 출력
        fprintf('\n  【%s 결과】\n', current_method);
        fprintf('    관찰 AUC: %.4f, p-value: %.4f\n', observed_auc, p_value_auc);
        fprintf('    관찰 F1:  %.4f, p-value: %.4f\n', observed_f1, p_value_f1);
        fprintf('    실행시간: %.1f초, 실패횟수: %d\n', elapsed_time, failed_permutations);
        if p_value_auc < 0.05 || p_value_f1 < 0.05
            fprintf('    ✓ 통계적으로 유의함\n');
        else
            fprintf('    - 통계적으로 유의하지 않음\n');
        end
    end
end

% 비교 요약 테이블 생성 (사용 가능한 방법들만)
fprintf('\n【%d가지 방법 퍼뮤테이션 비교 요약】\n', n_available_methods);
fprintf('════════════════════════════════════════════════════════════\n');

if n_available_methods > 0
    comparison_table = table();
    comparison_table.Method = available_methods';
    comparison_table.Observed_AUC = zeros(n_available_methods, 1);
    comparison_table.AUC_p_value = zeros(n_available_methods, 1);
    comparison_table.Observed_F1 = zeros(n_available_methods, 1);
    comparison_table.F1_p_value = zeros(n_available_methods, 1);
    comparison_table.AUC_Effect_Size = zeros(n_available_methods, 1);
    comparison_table.F1_Effect_Size = zeros(n_available_methods, 1);
    comparison_table.Overall_Significance = cell(n_available_methods, 1);

    for i = 1:n_available_methods
        method = available_methods{i};
        if isfield(all_permutation_results, method)
            results = all_permutation_results.(method);
            comparison_table.Observed_AUC(i) = results.auc.observed;
            comparison_table.AUC_p_value(i) = results.auc.p_value;
            comparison_table.Observed_F1(i) = results.f1.observed;
            comparison_table.F1_p_value(i) = results.f1.p_value;
            comparison_table.AUC_Effect_Size(i) = results.auc.z_score;
            comparison_table.F1_Effect_Size(i) = results.f1.z_score;

            % 전체 유의성 판단
            min_p = min(results.auc.p_value, results.f1.p_value);
            if min_p < 0.001
                comparison_table.Overall_Significance{i} = '***';
            elseif min_p < 0.01
                comparison_table.Overall_Significance{i} = '**';
            elseif min_p < 0.05
                comparison_table.Overall_Significance{i} = '*';
            else
                comparison_table.Overall_Significance{i} = '';
            end
        end
    end

    % 테이블 출력
    disp(comparison_table);

    % 가장 robust한 방법 식별
    valid_auc_p = comparison_table.AUC_p_value(comparison_table.AUC_p_value > 0);
    valid_f1_p = comparison_table.F1_p_value(comparison_table.F1_p_value > 0);

    if ~isempty(valid_auc_p)
        [~, best_auc_idx] = min(comparison_table.AUC_p_value);
        fprintf('\n【최고 성능 방법】\n');
        fprintf('  AUC 기준:  %s (p=%.4f)\n', available_methods{best_auc_idx}, comparison_table.AUC_p_value(best_auc_idx));
    end

    if ~isempty(valid_f1_p)
        [~, best_f1_idx] = min(comparison_table.F1_p_value);
        fprintf('  F1 기준:   %s (p=%.4f)\n', available_methods{best_f1_idx}, comparison_table.F1_p_value(best_f1_idx));
    end
end

fprintf('\n✅ STEP 22.6 모든 가중치 방법 퍼뮤테이션 테스트 비교 완료\n');
fprintf('   - 처리된 방법: %d개\n', n_available_methods);
fprintf('   - 건너뛴 방법: %d개\n', n_all_methods - n_available_methods);

end  % STEP 22.6 else 블록 닫기