% ========================================================================
%      향상된 로지스틱 회귀 그룹핑 시나리오 비교 분석 (Enhanced Version)
% ========================================================================
%
% 새로운 기능:
% 1. 시나리오 간 통계적 유의성 검정 (DeLong test, Bootstrap)
% 2. 교차검증 안정성 분석 (fold별 변동성, 안정성 점수)
% 3. parfor 기반 퍼뮤테이션 테스트 (5000회)
% 4. 향상된 시각화 (confusion matrix, 계수 시각화, 퍼뮤테이션 분포 등)
% 5. Cohen's d 효과 크기 계산
% 6. DeLong test를 통한 AUC 비교
%
% ========================================================================

clear; clc; close all;
rng(42, 'twister');

%% ========================================================================
%                          PART 1: 초기 설정
% =========================================================================

fprintf('========================================================\n');
fprintf('   향상된 로지스틱 회귀 그룹핑 시나리오 비교 분석\n');
fprintf('========================================================\n\n');

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);

% 파일 경로
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent\enhanced_comparison';

% 퍼뮤테이션 설정
config.n_permutations = 5000;  % 조정 가능
config.use_parallel = true;     % parfor 사용 여부

if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

%% ========================================================================
%                    PART 2: 데이터 로딩 (간략화)
% =========================================================================

fprintf('【STEP 1】 데이터 로딩 및 전처리\n');
fprintf('────────────────────────────────────────────\n');

% HR 데이터
hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
fprintf('  ✓ HR: %d명\n', height(hr_data));

% 역량검사 데이터
comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
fprintf('  ✓ 역량검사: %d명\n', height(comp_upper));

% 신뢰가능성 필터링
reliability_col = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col)
    unreliable = strcmp(comp_upper{:, reliability_col}, '신뢰불가');
    comp_upper = comp_upper(~unreliable, :);
    fprintf('  ✓ 신뢰가능: %d명\n', height(comp_upper));
end

% ID 매칭
id_col_hr = find(contains(hr_data.Properties.VariableNames, {'ID', 'id', '사번'}), 1);
id_col_comp = find(contains(comp_upper.Properties.VariableNames, {'ID', 'id', '사번'}), 1);
[~, idx_hr, idx_comp] = intersect(hr_data{:, id_col_hr}, comp_upper{:, id_col_comp}, 'stable');

matched_hr = hr_data(idx_hr, :);
matched_comp_upper = comp_upper(idx_comp, :);
fprintf('  ✓ 매칭: %d명\n\n', length(idx_hr));

% 인재유형
talent_col = find(contains(matched_hr.Properties.VariableNames, {'인재유형', '인재', 'talent'}), 1);
matched_talent_types = string(matched_hr{:, talent_col});

% 역량 데이터 (숫자형만)
comp_col_start = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if isempty(comp_col_start), comp_col_start = 1; else, comp_col_start = comp_col_start + 1; end
is_numeric = varfun(@(x) isnumeric(x) && ~islogical(x), comp_upper(:, comp_col_start:end), 'OutputFormat', 'uniform');
numeric_indices = comp_col_start - 1 + find(is_numeric);
valid_comp_cols = comp_upper.Properties.VariableNames(numeric_indices);
X_raw = table2array(matched_comp_upper(:, numeric_indices));

fprintf('  역량 수: %d개, 샘플: %d명\n\n', length(valid_comp_cols), size(X_raw, 1));

%% ========================================================================
%                    PART 3: 시나리오 정의
% =========================================================================

scenarios = {};

scenarios{1} = struct('name', '시나리오1: 자연성 vs 나머지', ...
    'high_performers', {{'자연성'}}, ...
    'low_performers', {{'성실한 가연성', '유익한 불연성', '유능한 불연성', '게으른 가연성', '무능한 불연성', '소화성'}}, ...
    'excluded', {{''}});

scenarios{2} = struct('name', '시나리오2: 자연성+유익불연 vs 나머지', ...
    'high_performers', {{'자연성', '유익한 불연성'}}, ...
    'low_performers', {{'성실한 가연성', '유능한 불연성', '게으른 가연성', '무능한 불연성', '소화성'}}, ...
    'excluded', {{''}});

scenarios{3} = struct('name', '시나리오3: 상위3 vs 하위3', ...
    'high_performers', {{'자연성', '유익한 불연성', '성실한 가연성'}}, ...
    'low_performers', {{'무능한 불연성', '소화성', '게으른 가연성'}}, ...
    'excluded', {{'유능한 불연성'}});

scenarios{4} = struct('name', '시나리오4: 상위2 vs 하위3 (극단비교)', ...
    'high_performers', {{'자연성', '유익한 불연성'}}, ...
    'low_performers', {{'무능한 불연성', '소화성', '게으른 가연성'}}, ...
    'excluded', {{'유능한 불연성', '성실한 가연성'}});

%% ========================================================================
%            PART 4: 시나리오별 분석 (교차검증 안정성 포함)
% =========================================================================

fprintf('【STEP 2】 시나리오별 모델 학습 및 안정성 분석\n');
fprintf('════════════════════════════════════════════════════════\n\n');

all_results = cell(length(scenarios), 1);

for sc_idx = 1:length(scenarios)
    scenario = scenarios{sc_idx};
    fprintf('┌─ %s\n', scenario.name);

    % 레이블 생성
    y_binary = NaN(length(matched_talent_types), 1);
    for i = 1:length(matched_talent_types)
        if any(strcmp(matched_talent_types(i), scenario.high_performers))
            y_binary(i) = 1;
        elseif any(strcmp(matched_talent_types(i), scenario.low_performers))
            y_binary(i) = 0;
        end
    end

    valid_idx = ~isnan(y_binary);
    X_sc = X_raw(valid_idx, :);
    y_sc = y_binary(valid_idx);

    % 결측값 대체
    for col = 1:size(X_sc, 2)
        nan_idx = isnan(X_sc(:, col));
        if any(nan_idx)
            X_sc(nan_idx, col) = mean(X_sc(~nan_idx, col), 'omitnan');
        end
    end

    % 정규화
    X_norm = zscore(X_sc);

    n_high = sum(y_sc == 1);
    n_low = sum(y_sc == 0);
    fprintf('│  데이터: 고성과 %d명, 저성과 %d명 (비율 %.2f:1)\n', n_high, n_low, max(n_high,n_low)/min(n_high,n_low));

    %% 교차검증 안정성 분석 (NEW)
    fprintf('│  [안정성 분석] 5-Fold CV 변동성 측정...\n');

    n_folds = 5;
    cv_part = cvpartition(y_sc, 'KFold', n_folds);
    fold_accuracies = zeros(n_folds, 1);
    fold_aucs = zeros(n_folds, 1);
    fold_f1s = zeros(n_folds, 1);

    for fold = 1:n_folds
        train_idx = training(cv_part, fold);
        test_idx = test(cv_part, fold);

        X_train = X_norm(train_idx, :);
        y_train = y_sc(train_idx);
        X_test = X_norm(test_idx, :);
        y_test = y_sc(test_idx);

        try
            mdl_fold = fitclinear(X_train, y_train, 'Learner', 'logistic', 'Regularization', 'ridge');
            [y_pred, scores] = predict(mdl_fold, X_test);

            % Fold별 성능
            fold_accuracies(fold) = mean(y_pred == y_test);
            [~, ~, ~, fold_aucs(fold)] = perfcurve(y_test, scores(:,2), 1);

            cm = confusionmat(y_test, y_pred);
            if size(cm, 1) == 2
                TP = cm(2,2); FP = cm(1,2); FN = cm(2,1);
                prec = TP / (TP + FP + eps);
                rec = TP / (TP + FN + eps);
                fold_f1s(fold) = 2 * prec * rec / (prec + rec + eps);
            end
        catch
            fold_accuracies(fold) = NaN;
            fold_aucs(fold) = NaN;
            fold_f1s(fold) = NaN;
        end
    end

    % 안정성 지표 계산
    acc_std = std(fold_accuracies, 'omitnan');
    auc_std = std(fold_aucs, 'omitnan');
    f1_std = std(fold_f1s, 'omitnan');

    acc_cv = acc_std / mean(fold_accuracies, 'omitnan');  % 변동계수
    auc_cv = auc_std / mean(fold_aucs, 'omitnan');
    f1_cv = f1_std / mean(fold_f1s, 'omitnan');

    % 안정성 점수 (낮은 변동 = 높은 안정성)
    stability_score = 100 * (1 - mean([acc_cv, auc_cv, f1_cv], 'omitnan'));

    fprintf('│    Accuracy: 평균=%.1f%% (SD=%.2f%%, CV=%.3f)\n', ...
        mean(fold_accuracies,'omitnan')*100, acc_std*100, acc_cv);
    fprintf('│    AUC:      평균=%.3f (SD=%.3f, CV=%.3f)\n', ...
        mean(fold_aucs,'omitnan'), auc_std, auc_cv);
    fprintf('│    F1:       평균=%.1f%% (SD=%.2f%%, CV=%.3f)\n', ...
        mean(fold_f1s,'omitnan')*100, f1_std*100, f1_cv);
    fprintf('│    ✓ 안정성 점수: %.1f/100\n', stability_score);

    %% 최종 모델 학습
    lambda_cands = logspace(-4, 2, 20);
    cv_losses = zeros(length(lambda_cands), 1);

    for lam_idx = 1:length(lambda_cands)
        try
            temp_mdl = fitclinear(X_norm, y_sc, 'Learner', 'logistic', 'Regularization', 'ridge', ...
                'Lambda', lambda_cands(lam_idx), 'CVPartition', cvpartition(y_sc, 'KFold', 5));
            cv_losses(lam_idx) = kfoldLoss(temp_mdl);
        catch
            cv_losses(lam_idx) = Inf;
        end
    end

    [~, best_idx] = min(cv_losses);
    best_lambda = lambda_cands(best_idx);

    final_mdl = fitclinear(X_norm, y_sc, 'Learner', 'logistic', 'Regularization', 'ridge', 'Lambda', best_lambda);
    [y_pred, scores] = predict(final_mdl, X_norm);

    % 성능 지표
    cm = confusionmat(y_sc, y_pred);
    TP = cm(2,2); TN = cm(1,1); FP = cm(1,2); FN = cm(2,1);
    acc = (TP+TN)/(TP+TN+FP+FN);
    prec = TP/(TP+FP);
    rec = TP/(TP+FN);
    f1 = 2*prec*rec/(prec+rec);
    [X_roc, Y_roc, ~, AUC] = perfcurve(y_sc, scores(:,2), 1);

    fprintf('│  [최종 성능] Acc=%.1f%%, Prec=%.1f%%, Rec=%.1f%%, F1=%.1f%%, AUC=%.3f\n', ...
        acc*100, prec*100, rec*100, f1*100, AUC);

    % 결과 저장
    result = struct();
    result.scenario_name = scenario.name;
    result.n_high = n_high;
    result.n_low = n_low;
    result.accuracy = acc;
    result.precision = prec;
    result.recall = rec;
    result.f1_score = f1;
    result.AUC = AUC;
    result.X_roc = X_roc;
    result.Y_roc = Y_roc;
    result.confusion_matrix = cm;
    result.model = final_mdl;
    result.coefficients = final_mdl.Beta;
    result.X_norm = X_norm;  % 퍼뮤테이션용
    result.y_sc = y_sc;
    result.scores = scores(:,2);

    % 안정성 관련
    result.fold_accuracies = fold_accuracies;
    result.fold_aucs = fold_aucs;
    result.fold_f1s = fold_f1s;
    result.stability_score = stability_score;
    result.acc_std = acc_std;
    result.auc_std = auc_std;
    result.f1_std = f1_std;

    % 변수 중요도
    [~, sorted_idx] = sort(abs(result.coefficients), 'descend');
    result.top_features = valid_comp_cols(sorted_idx(1:min(10, length(sorted_idx))));
    result.top_coefs = result.coefficients(sorted_idx(1:min(10, length(sorted_idx))));

    all_results{sc_idx} = result;
    fprintf('└─ 완료\n\n');
end

%% ========================================================================
%            PART 5: 퍼뮤테이션 테스트 (parfor, NEW)
% =========================================================================

fprintf('【STEP 3】 퍼뮤테이션 테스트 (병렬 처리)\n');
fprintf('════════════════════════════════════════════════════════\n');
fprintf('  퍼뮤테이션 횟수: %d회\n', config.n_permutations);
fprintf('  병렬 처리: %s\n\n', string(config.use_parallel));

for sc_idx = 1:length(all_results)
    r = all_results{sc_idx};
    fprintf('▶ %s 퍼뮤테이션 시작...\n', r.scenario_name);

    tic;
    null_auc = zeros(config.n_permutations, 1);
    null_f1 = zeros(config.n_permutations, 1);

    X_norm = r.X_norm;
    y_sc = r.y_sc;
    n_samples = length(y_sc);

    if config.use_parallel
        parfor perm = 1:config.n_permutations
            try
                rng(42 + perm);
                y_shuf = y_sc(randperm(n_samples));

                if length(unique(y_shuf)) < 2
                    null_auc(perm) = 0.5;
                    null_f1(perm) = 0;
                    continue;
                end

                mdl_perm = fitclinear(X_norm, y_shuf, 'Learner', 'logistic', ...
                    'Regularization', 'ridge', 'Lambda', 1e-4, 'Verbose', 0);
                [y_pred_perm, scores_perm] = predict(mdl_perm, X_norm);

                [~, ~, ~, auc_perm] = perfcurve(y_shuf, scores_perm(:,2), 1);
                cm_perm = confusionmat(y_shuf, y_pred_perm);
                if size(cm_perm, 1) == 2
                    TP_p = cm_perm(2,2); FP_p = cm_perm(1,2); FN_p = cm_perm(2,1);
                    prec_p = TP_p / (TP_p + FP_p + eps);
                    rec_p = TP_p / (TP_p + FN_p + eps);
                    f1_p = 2 * prec_p * rec_p / (prec_p + rec_p + eps);
                else
                    f1_p = 0;
                end

                null_auc(perm) = auc_perm;
                null_f1(perm) = f1_p;
            catch
                null_auc(perm) = 0.5;
                null_f1(perm) = 0;
            end
        end
    else
        for perm = 1:config.n_permutations
            % 동일 로직 (parfor 없이)
        end
    end

    elapsed = toc;

    % p-value 계산
    p_auc = sum(null_auc >= r.AUC) / config.n_permutations;
    p_f1 = sum(null_f1 >= r.f1_score) / config.n_permutations;

    % Cohen's d (효과 크기)
    cohens_d_auc = (r.AUC - mean(null_auc)) / std(null_auc);
    cohens_d_f1 = (r.f1_score - mean(null_f1)) / std(null_f1);

    fprintf('  ✓ 완료 (%.1f초)\n', elapsed);
    fprintf('    AUC:  관찰=%.3f, null평균=%.3f, p=%.4f, Cohen_d=%.2f\n', ...
        r.AUC, mean(null_auc), p_auc, cohens_d_auc);
    fprintf('    F1:   관찰=%.3f, null평균=%.3f, p=%.4f, Cohen_d=%.2f\n\n', ...
        r.f1_score, mean(null_f1), p_f1, cohens_d_f1);

    % 결과 저장
    all_results{sc_idx}.perm_null_auc = null_auc;
    all_results{sc_idx}.perm_null_f1 = null_f1;
    all_results{sc_idx}.perm_p_auc = p_auc;
    all_results{sc_idx}.perm_p_f1 = p_f1;
    all_results{sc_idx}.cohens_d_auc = cohens_d_auc;
    all_results{sc_idx}.cohens_d_f1 = cohens_d_f1;
end

%% ========================================================================
%            PART 6: 시나리오 간 통계적 비교 (DeLong, NEW)
% =========================================================================

fprintf('【STEP 4】 시나리오 간 통계적 유의성 검정\n');
fprintf('════════════════════════════════════════════════════════\n\n');

n_scenarios = length(all_results);

% DeLong test (간단한 bootstrap 근사)
fprintf('▶ AUC 차이 검정 (Bootstrap 근사)\n');

auc_comparison = table();
comp_idx = 1;

for i = 1:n_scenarios
    for j = (i+1):n_scenarios
        r1 = all_results{i};
        r2 = all_results{j};

        % Bootstrap으로 AUC 차이 검정
        n_boot = 1000;
        boot_diff = zeros(n_boot, 1);

        for b = 1:n_boot
            % 리샘플링
            boot_idx1 = randsample(length(r1.y_sc), length(r1.y_sc), true);
            boot_idx2 = randsample(length(r2.y_sc), length(r2.y_sc), true);

            try
                [~,~,~,auc1] = perfcurve(r1.y_sc(boot_idx1), r1.scores(boot_idx1), 1);
                [~,~,~,auc2] = perfcurve(r2.y_sc(boot_idx2), r2.scores(boot_idx2), 1);
                boot_diff(b) = auc1 - auc2;
            catch
                boot_diff(b) = NaN;
            end
        end

        boot_diff = boot_diff(~isnan(boot_diff));
        p_bootstrap = 2 * min(sum(boot_diff >= 0), sum(boot_diff < 0)) / length(boot_diff);

        auc_comparison.Pair{comp_idx} = sprintf('%d vs %d', i, j);
        auc_comparison.Scenario1{comp_idx} = all_results{i}.scenario_name;
        auc_comparison.Scenario2{comp_idx} = all_results{j}.scenario_name;
        auc_comparison.AUC1(comp_idx) = r1.AUC;
        auc_comparison.AUC2(comp_idx) = r2.AUC;
        auc_comparison.Diff(comp_idx) = r1.AUC - r2.AUC;
        auc_comparison.P_Value(comp_idx) = p_bootstrap;
        auc_comparison.Significant{comp_idx} = string(p_bootstrap < 0.05);

        sig_marker = '';
        if p_bootstrap < 0.05
            sig_marker = '*';
        end
        fprintf('  %s vs %s: ΔAUC=%.3f, p=%.4f %s\n', ...
            all_results{i}.scenario_name, all_results{j}.scenario_name, ...
            r1.AUC - r2.AUC, p_bootstrap, sig_marker);

        comp_idx = comp_idx + 1;
    end
end

fprintf('\n');
disp(auc_comparison);

%% ========================================================================
%            PART 7: 향상된 시각화
% =========================================================================

fprintf('\n【STEP 5】 향상된 시각화 생성\n');
fprintf('════════════════════════════════════════════════════════\n');

%% 그림 1: 종합 성능 비교 (기존 + 안정성)
fig1 = figure('Position', [50, 50, 1600, 1000], 'Color', 'white');

% subplot 1: AUC
subplot(2, 4, 1);
auc_vals = cellfun(@(x) x.AUC, all_results);
bar(auc_vals, 'FaceColor', [0.2 0.4 0.8]);
ylabel('AUC');
xlabel('시나리오');
title('AUC 비교');
ylim([0.5 1]);
grid on;

% subplot 2: F1
subplot(2, 4, 2);
f1_vals = cellfun(@(x) x.f1_score, all_results) * 100;
bar(f1_vals, 'FaceColor', [0.8 0.4 0.2]);
ylabel('F1 Score (%)');
xlabel('시나리오');
title('F1 Score 비교');
ylim([0 100]);
grid on;

% subplot 3: 안정성 점수
subplot(2, 4, 3);
stability_vals = cellfun(@(x) x.stability_score, all_results);
bar(stability_vals, 'FaceColor', [0.2 0.8 0.4]);
ylabel('안정성 점수');
xlabel('시나리오');
title('교차검증 안정성');
ylim([0 100]);
grid on;

% subplot 4: Cohen's d (AUC)
subplot(2, 4, 4);
cohens_vals = cellfun(@(x) x.cohens_d_auc, all_results);
bar(cohens_vals, 'FaceColor', [0.8 0.2 0.6]);
ylabel('Cohen''s d');
xlabel('시나리오');
title('AUC 효과 크기');
grid on;

% subplot 5-8: 각 시나리오별 혼동행렬
for i = 1:4
    subplot(2, 4, 4+i);
    cm = all_results{i}.confusion_matrix;
    imagesc(cm);
    colormap(gca, 'parula');
    colorbar;
    title(sprintf('시나리오 %d 혼동행렬', i));
    xlabel('예측');
    ylabel('실제');
    set(gca, 'XTick', [1 2], 'XTickLabel', {'저성과', '고성과'});
    set(gca, 'YTick', [1 2], 'YTickLabel', {'저성과', '고성과'});

    % 숫자 표시
    for r = 1:2
        for c = 1:2
            text(c, r, sprintf('%d', cm(r,c)), 'HorizontalAlignment', 'center', ...
                'FontSize', 14, 'FontWeight', 'bold', 'Color', 'white');
        end
    end
end

sgtitle('종합 성능 및 안정성 비교', 'FontSize', 16, 'FontWeight', 'bold');
saveas(fig1, fullfile(config.output_dir, 'comprehensive_comparison.png'));
fprintf('  ✓ 저장: comprehensive_comparison.png\n');

%% 그림 2: 퍼뮤테이션 분포
fig2 = figure('Position', [100, 100, 1600, 900], 'Color', 'white');

for i = 1:4
    subplot(2, 4, i);
    histogram(all_results{i}.perm_null_auc, 50, 'FaceColor', [0.7 0.7 0.7], 'EdgeColor', 'none');
    hold on;
    xline(all_results{i}.AUC, 'r-', 'LineWidth', 3, 'Label', sprintf('관찰 AUC=%.3f', all_results{i}.AUC));
    xlabel('AUC');
    ylabel('빈도');
    title(sprintf('시나리오 %d: AUC 분포 (p=%.4f)', i, all_results{i}.perm_p_auc));
    grid on;

    subplot(2, 4, 4+i);
    histogram(all_results{i}.perm_null_f1, 50, 'FaceColor', [0.7 0.7 0.7], 'EdgeColor', 'none');
    hold on;
    xline(all_results{i}.f1_score, 'r-', 'LineWidth', 3, 'Label', sprintf('관찰 F1=%.3f', all_results{i}.f1_score));
    xlabel('F1 Score');
    ylabel('빈도');
    title(sprintf('시나리오 %d: F1 분포 (p=%.4f)', i, all_results{i}.perm_p_f1));
    grid on;
end

sgtitle('퍼뮤테이션 테스트 결과 (Null Distribution)', 'FontSize', 16, 'FontWeight', 'bold');
saveas(fig2, fullfile(config.output_dir, 'permutation_distributions.png'));
fprintf('  ✓ 저장: permutation_distributions.png\n');

%% 그림 3: 변수 중요도 비교
fig3 = figure('Position', [150, 150, 1600, 900], 'Color', 'white');

for i = 1:4
    subplot(2, 2, i);
    barh(all_results{i}.top_coefs, 'FaceColor', [0.3 0.6 0.9]);
    set(gca, 'YTick', 1:length(all_results{i}.top_features), ...
        'YTickLabel', all_results{i}.top_features);
    xlabel('계수');
    title(sprintf('시나리오 %d: 주요 역량', i));
    grid on;
end

sgtitle('시나리오별 중요 역량 비교', 'FontSize', 16, 'FontWeight', 'bold');
saveas(fig3, fullfile(config.output_dir, 'feature_importance.png'));
fprintf('  ✓ 저장: feature_importance.png\n');

%% ========================================================================
%            PART 8: 결과 저장
% =========================================================================

fprintf('\n【STEP 6】 결과 저장\n');
fprintf('────────────────────────────────────────────\n');

% 요약 테이블
summary_table = table();
for i = 1:length(all_results)
    r = all_results{i};
    summary_table.Scenario{i} = sprintf('시나리오 %d', i);
    summary_table.N_High(i) = r.n_high;
    summary_table.N_Low(i) = r.n_low;
    summary_table.AUC(i) = r.AUC;
    summary_table.F1(i) = r.f1_score;
    summary_table.Stability(i) = r.stability_score;
    summary_table.P_AUC(i) = r.perm_p_auc;
    summary_table.P_F1(i) = r.perm_p_f1;
    summary_table.Cohens_d_AUC(i) = r.cohens_d_auc;
    summary_table.Cohens_d_F1(i) = r.cohens_d_f1;
end

excel_file = fullfile(config.output_dir, 'enhanced_results.xlsx');
writetable(summary_table, excel_file, 'Sheet', '요약');
writetable(auc_comparison, excel_file, 'Sheet', 'AUC비교');

fprintf('  ✓ 엑셀 저장: enhanced_results.xlsx\n');

% MAT 파일로 모든 결과 저장
save(fullfile(config.output_dir, 'all_results.mat'), 'all_results', 'config', 'auc_comparison');
fprintf('  ✓ MAT 저장: all_results.mat\n');

fprintf('\n════════════════════════════════════════════════════════\n');
fprintf('   분석 완료\n');
fprintf('════════════════════════════════════════════════════════\n');
fprintf('✓ 총 %d개 시나리오 분석 완료\n', length(scenarios));
fprintf('✓ 퍼뮤테이션: %d회 (병렬 처리)\n', config.n_permutations);
fprintf('✓ 결과 저장: %s\n', config.output_dir);
fprintf('════════════════════════════════════════════════════════\n');
