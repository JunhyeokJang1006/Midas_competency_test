% ========================================================================
%    로지스틱 회귀 기반 고성과자/저성과자 그룹핑 시나리오 전체 비교 분석
% ========================================================================
% 목적: 다양한 그룹핑 시나리오에 대해 데이터 로딩부터 모델 학습, 예측,
%       평가까지 전체 파이프라인을 실행하고 결과를 비교
%
% 분석 시나리오:
% 1. 자연성 vs 나머지
% 2. 자연성+유익한불연성 vs 나머지
% 3. 자연성+유익한불연성+성실한가연성 vs 무능한불연성+소화성+게으른가연성
%    (유능한불연성 제외)
% 4. 자연성+유익한불연성 vs 무능한불연성+소화성+게으른가연성
%    (유능한불연성+성실한가연성 제외)
%
% ========================================================================

clear; clc; close all;
rng(42, 'twister');  % 재현성 보장

%% ========================================================================
%                          PART 1: 초기 설정
% =========================================================================

fprintf('========================================================\n');
fprintf('   로지스틱 회귀 그룹핑 시나리오 전체 비교 분석\n');
fprintf('========================================================\n\n');

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent\scenario_comparison';

% 출력 디렉토리 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
    fprintf('✓ 출력 디렉토리 생성: %s\n\n', config.output_dir);
end

%% ========================================================================
%                    PART 2: 데이터 로딩 및 전처리
% =========================================================================

fprintf('【STEP 1】 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

try
    % HR 데이터 로딩
    fprintf('▶ HR 데이터 로딩 중...\n');
    hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
    fprintf('  ✓ HR 데이터: %d명\n', height(hr_data));

    % 역량검사 데이터 로딩
    fprintf('▶ 역량검사 데이터 로딩 중...\n');
    comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 상위항목 데이터: %d명\n\n', height(comp_upper));
catch ME
    error('데이터 로딩 실패: %s', ME.message);
end

%% 신뢰가능성 필터링
fprintf('【STEP 2】 신뢰가능성 필터링\n');
fprintf('────────────────────────────────────────────\n');

reliability_col_idx = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
original_count = height(comp_upper);

if ~isempty(reliability_col_idx)
    reliability_data = comp_upper{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(comp_upper), 1);
    end

    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));
    comp_upper = comp_upper(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능 데이터: %d명 (제거율: %.1f%%)\n\n', ...
        height(comp_upper), sum(unreliable_idx)/original_count*100);
else
    fprintf('  ⚠ 신뢰가능성 컬럼 없음. 모든 데이터 사용\n\n');
end

%% 데이터 매칭
fprintf('【STEP 3】 HR-역량검사 데이터 매칭\n');
fprintf('────────────────────────────────────────────\n');

% ID 컬럼 찾기
id_col_hr = find(contains(hr_data.Properties.VariableNames, {'ID', 'id', '사번'}), 1);
id_col_comp = find(contains(comp_upper.Properties.VariableNames, {'ID', 'id', '사번'}), 1);

if isempty(id_col_hr) || isempty(id_col_comp)
    error('ID 컬럼을 찾을 수 없습니다.');
end

% 매칭
hr_ids = hr_data{:, id_col_hr};
comp_ids = comp_upper{:, id_col_comp};
[common_ids, idx_hr, idx_comp] = intersect(hr_ids, comp_ids, 'stable');

matched_hr = hr_data(idx_hr, :);
matched_comp_upper = comp_upper(idx_comp, :);

fprintf('  원본 HR: %d명\n', height(hr_data));
fprintf('  원본 역량검사: %d명\n', height(comp_upper));
fprintf('  ✓ 매칭 성공: %d명 (매칭률: %.1f%%)\n\n', ...
    length(common_ids), length(common_ids)/min(height(hr_data), height(comp_upper))*100);

%% 인재유형 추출
fprintf('【STEP 4】 인재유형 추출\n');
fprintf('────────────────────────────────────────────\n');

talent_type_col = find(contains(matched_hr.Properties.VariableNames, ...
    {'인재유형', '인재', 'talent', '유형'}), 1);

if isempty(talent_type_col)
    error('인재유형 컬럼을 찾을 수 없습니다.');
end

matched_talent_types = matched_hr{:, talent_type_col};
if iscell(matched_talent_types)
    matched_talent_types = cellfun(@string, matched_talent_types);
else
    matched_talent_types = string(matched_talent_types);
end

% 인재유형 분포 확인
unique_types = unique(matched_talent_types);
fprintf('  인재유형 종류: %d개\n', length(unique_types));
for i = 1:length(unique_types)
    count = sum(strcmp(matched_talent_types, unique_types(i)));
    fprintf('    - %s: %d명 (%.1f%%)\n', unique_types(i), count, count/length(matched_talent_types)*100);
end
fprintf('\n');

%% 역량 데이터 추출
fprintf('【STEP 5】 역량 데이터 추출\n');
fprintf('────────────────────────────────────────────\n');

% 역량 컬럼 범위 찾기
comp_col_start = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if isempty(comp_col_start)
    comp_col_start = 1;
else
    comp_col_start = comp_col_start + 1;
end

% 숫자형 컬럼만 선택 (double 타입)
all_cols = comp_upper.Properties.VariableNames(comp_col_start:end);
is_numeric = varfun(@(x) isnumeric(x) && ~islogical(x), comp_upper(:, comp_col_start:end), 'OutputFormat', 'uniform');
numeric_col_indices = comp_col_start - 1 + find(is_numeric);

valid_comp_cols = comp_upper.Properties.VariableNames(numeric_col_indices);
matched_comp = matched_comp_upper(:, numeric_col_indices);
X_raw = table2array(matched_comp);

fprintf('  총 역량 수: %d개\n', length(valid_comp_cols));
fprintf('  샘플 수: %d명\n', size(X_raw, 1));
fprintf('  결측값 비율: %.2f%%\n', sum(isnan(X_raw(:)))/numel(X_raw)*100);

% 역량명 샘플 출력
fprintf('\n  역량명 샘플 (처음 5개):\n');
for i = 1:min(5, length(valid_comp_cols))
    fprintf('    %d. %s\n', i, valid_comp_cols{i});
end
fprintf('\n');

%% ========================================================================
%                    PART 3: 시나리오 정의 및 분석
% =========================================================================

fprintf('【STEP 6】 분석 시나리오 정의\n');
fprintf('────────────────────────────────────────────\n');

% 4가지 시나리오 정의
scenarios = {};

scenarios{1} = struct(...
    'name', '시나리오 1: 자연성 vs 나머지', ...
    'description', '자연성만을 고성과자로 분류', ...
    'high_performers', {{'자연성'}}, ...
    'low_performers', {{'성실한 가연성', '유익한 불연성', '유능한 불연성', ...
                        '게으른 가연성', '무능한 불연성', '소화성'}}, ...
    'excluded', {{''}});

scenarios{2} = struct(...
    'name', '시나리오 2: 자연성+유익한불연성 vs 나머지', ...
    'description', '자연성과 유익한불연성을 고성과자로 분류', ...
    'high_performers', {{'자연성', '유익한 불연성'}}, ...
    'low_performers', {{'성실한 가연성', '유능한 불연성', ...
                        '게으른 가연성', '무능한 불연성', '소화성'}}, ...
    'excluded', {{''}});

scenarios{3} = struct(...
    'name', '시나리오 3: 상위3 vs 하위3 (유능한불연성 제외)', ...
    'description', '상위 3개 vs 하위 3개, 중위권(유능한불연성) 제외', ...
    'high_performers', {{'자연성', '유익한 불연성', '성실한 가연성'}}, ...
    'low_performers', {{'무능한 불연성', '소화성', '게으른 가연성'}}, ...
    'excluded', {{'유능한 불연성'}});

scenarios{4} = struct(...
    'name', '시나리오 4: 상위2 vs 하위3 (중위2 제외)', ...
    'description', '명확한 상위 2개 vs 명확한 하위 3개, 중위권 모두 제외', ...
    'high_performers', {{'자연성', '유익한 불연성'}}, ...
    'low_performers', {{'무능한 불연성', '소화성', '게으른 가연성'}}, ...
    'excluded', {{'유능한 불연성', '성실한 가연성'}});

% 시나리오 정보 출력
for i = 1:length(scenarios)
    fprintf('\n%s\n', scenarios{i}.name);
    fprintf('  설명: %s\n', scenarios{i}.description);
    fprintf('  고성과자: %s\n', strjoin(scenarios{i}.high_performers, ', '));
    fprintf('  저성과자: %s\n', strjoin(scenarios{i}.low_performers, ', '));
    if ~isempty(scenarios{i}.excluded{1})
        fprintf('  제외: %s\n', strjoin(scenarios{i}.excluded, ', '));
    end
end
fprintf('\n');

%% ========================================================================
%                PART 4: 시나리오별 전체 파이프라인 실행
% =========================================================================

fprintf('【STEP 7】 시나리오별 분석 실행\n');
fprintf('════════════════════════════════════════════════════════\n\n');

% 결과 저장용
all_results = cell(length(scenarios), 1);

for scenario_idx = 1:length(scenarios)
    scenario = scenarios{scenario_idx};

    fprintf('┌────────────────────────────────────────────────────────┐\n');
    fprintf('│ %s\n', scenario.name);
    fprintf('│ %s\n', scenario.description);
    fprintf('└────────────────────────────────────────────────────────┘\n\n');

    %% 4.1 시나리오별 레이블 생성
    fprintf('▶ [단계 1/6] 레이블 생성\n');

    y_binary = NaN(length(matched_talent_types), 1);
    for i = 1:length(matched_talent_types)
        type_name = matched_talent_types(i);
        if any(strcmp(type_name, scenario.high_performers))
            y_binary(i) = 1;  % 고성과자
        elseif any(strcmp(type_name, scenario.low_performers))
            y_binary(i) = 0;  % 저성과자
        elseif any(strcmp(type_name, scenario.excluded))
            y_binary(i) = NaN;  % 제외
        end
    end

    % 유효한 데이터만 선택
    valid_idx = ~isnan(y_binary);
    X_scenario = X_raw(valid_idx, :);
    y_scenario = y_binary(valid_idx);
    talent_types_scenario = matched_talent_types(valid_idx);

    n_high = sum(y_scenario == 1);
    n_low = sum(y_scenario == 0);
    n_total = length(y_scenario);

    fprintf('  고성과자: %d명 (%.1f%%)\n', n_high, n_high/n_total*100);
    fprintf('  저성과자: %d명 (%.1f%%)\n', n_low, n_low/n_total*100);
    fprintf('  제외: %d명\n', sum(~valid_idx));
    fprintf('  분석 대상: %d명\n\n', n_total);

    %% 4.2 결측값 처리
    fprintf('▶ [단계 2/6] 결측값 처리\n');

    missing_ratio_before = sum(isnan(X_scenario(:))) / numel(X_scenario) * 100;
    fprintf('  결측값 비율 (처리 전): %.2f%%\n', missing_ratio_before);

    % 역량별 평균으로 대체
    X_processed = X_scenario;
    for col = 1:size(X_processed, 2)
        col_data = X_processed(:, col);
        nan_idx = isnan(col_data);
        if any(nan_idx)
            col_mean = mean(col_data(~nan_idx));
            if ~isnan(col_mean)
                X_processed(nan_idx, col) = col_mean;
            else
                X_processed(nan_idx, col) = 50; % 전체가 NaN인 경우 중간값
            end
        end
    end

    missing_ratio_after = sum(isnan(X_processed(:))) / numel(X_processed) * 100;
    fprintf('  결측값 비율 (처리 후): %.2f%%\n', missing_ratio_after);
    fprintf('  ✓ 결측값 처리 완료\n\n');

    %% 4.3 데이터 정규화
    fprintf('▶ [단계 3/6] 데이터 정규화 (Z-score)\n');

    X_normalized = zscore(X_processed);
    fprintf('  ✓ 정규화 완료 (평균=0, 표준편차=1)\n\n');

    %% 4.4 클래스 불균형 분석
    fprintf('▶ [단계 4/6] 클래스 불균형 분석\n');

    imbalance_ratio = max(n_high, n_low) / min(n_high, n_low);
    fprintf('  불균형 비율: %.2f:1\n', imbalance_ratio);

    if imbalance_ratio > 3
        fprintf('  ⚠ 심각한 클래스 불균형 (>3:1)\n');
        balance_status = 'severe';
    elseif imbalance_ratio > 2
        fprintf('  △ 중간 클래스 불균형 (2:1~3:1)\n');
        balance_status = 'moderate';
    elseif imbalance_ratio > 1.5
        fprintf('  △ 경미한 클래스 불균형 (1.5:1~2:1)\n');
        balance_status = 'mild';
    else
        fprintf('  ✓ 균형잡힌 클래스 분포\n');
        balance_status = 'balanced';
    end
    fprintf('\n');

    %% 4.5 로지스틱 회귀 모델 학습
    fprintf('▶ [단계 5/6] 로지스틱 회귀 모델 학습\n');

    try
        % 교차검증 설정 (5-Fold)
        cv_partitions = cvpartition(y_scenario, 'KFold', 5);

        fprintf('  모델: Logistic Regression (Ridge 정규화)\n');
        fprintf('  교차검증: 5-Fold CV\n');

        % 하이퍼파라미터 튜닝을 위한 Lambda 후보
        lambda_candidates = logspace(-4, 2, 20);
        cv_losses = zeros(length(lambda_candidates), 1);

        fprintf('  하이퍼파라미터 튜닝 중...\n');
        for lambda_idx = 1:length(lambda_candidates)
            try
                temp_mdl = fitclinear(X_normalized, y_scenario, ...
                    'Learner', 'logistic', ...
                    'Regularization', 'ridge', ...
                    'Lambda', lambda_candidates(lambda_idx), ...
                    'CVPartition', cv_partitions);
                cv_losses(lambda_idx) = kfoldLoss(temp_mdl);
            catch
                cv_losses(lambda_idx) = Inf;
            end
        end

        [min_loss, best_lambda_idx] = min(cv_losses);
        best_lambda = lambda_candidates(best_lambda_idx);

        fprintf('  최적 Lambda: %.6f (CV Loss: %.4f)\n', best_lambda, min_loss);

        % 최적 하이퍼파라미터로 최종 모델 학습
        final_mdl = fitclinear(X_normalized, y_scenario, ...
            'Learner', 'logistic', ...
            'Regularization', 'ridge', ...
            'Lambda', best_lambda);

        fprintf('  ✓ 모델 학습 완료\n\n');

        %% 4.6 모델 평가
        fprintf('▶ [단계 6/6] 모델 평가\n');

        % 예측
        [y_pred, scores] = predict(final_mdl, X_normalized);

        % 혼동행렬
        cm = confusionmat(y_scenario, y_pred);

        % 성능 지표
        TP = cm(2, 2);
        TN = cm(1, 1);
        FP = cm(1, 2);
        FN = cm(2, 1);

        accuracy = (TP + TN) / (TP + TN + FP + FN);
        precision = TP / (TP + FP);
        recall = TP / (TP + FN);
        specificity = TN / (TN + FP);
        f1_score = 2 * (precision * recall) / (precision + recall);

        % AUC-ROC
        [X_roc, Y_roc, ~, AUC] = perfcurve(y_scenario, scores(:,2), 1);

        fprintf('  혼동행렬:\n');
        fprintf('                예측 저성과    예측 고성과\n');
        fprintf('  실제 저성과    %6d         %6d\n', TN, FP);
        fprintf('  실제 고성과    %6d         %6d\n\n', FN, TP);

        fprintf('  성능 지표:\n');
        fprintf('    Accuracy:    %.2f%%\n', accuracy * 100);
        fprintf('    Precision:   %.2f%%\n', precision * 100);
        fprintf('    Recall:      %.2f%%\n', recall * 100);
        fprintf('    Specificity: %.2f%%\n', specificity * 100);
        fprintf('    F1 Score:    %.2f%%\n', f1_score * 100);
        fprintf('    AUC-ROC:     %.4f\n', AUC);

        % 성능 등급 판정
        if AUC >= 0.9
            performance_grade = 'Excellent';
        elseif AUC >= 0.8
            performance_grade = 'Good';
        elseif AUC >= 0.7
            performance_grade = 'Fair';
        else
            performance_grade = 'Poor';
        end
        fprintf('    성능 등급:   %s\n\n', performance_grade);

        % 변수 중요도 (계수)
        coefficients = final_mdl.Beta;
        [sorted_coef, sorted_idx] = sort(abs(coefficients), 'descend');

        fprintf('  중요 역량 TOP 10:\n');
        for i = 1:min(10, length(sorted_coef))
            coef_val = coefficients(sorted_idx(i));
            direction = '고성과자 ↑';
            if coef_val < 0
                direction = '저성과자 ↑';
            end
            fprintf('    %2d. %-30s  계수: %+.4f  (%s)\n', ...
                i, valid_comp_cols{sorted_idx(i)}, coef_val, direction);
        end

        model_success = true;

    catch ME
        fprintf('  ✗ 모델 학습/평가 실패: %s\n', ME.message);
        model_success = false;

        % 실패 시 기본값
        accuracy = NaN;
        precision = NaN;
        recall = NaN;
        specificity = NaN;
        f1_score = NaN;
        AUC = NaN;
        cm = [];
        coefficients = [];
        sorted_idx = [];
        performance_grade = 'Failed';
    end

    fprintf('\n');

    %% 결과 저장
    result = struct();
    result.scenario_name = scenario.name;
    result.scenario_description = scenario.description;
    result.high_performers = scenario.high_performers;
    result.low_performers = scenario.low_performers;
    result.excluded = scenario.excluded;

    result.n_total = n_total;
    result.n_high = n_high;
    result.n_low = n_low;
    result.n_excluded = sum(~valid_idx);
    result.imbalance_ratio = imbalance_ratio;
    result.balance_status = balance_status;

    result.missing_ratio_before = missing_ratio_before;
    result.missing_ratio_after = missing_ratio_after;

    if model_success
        result.model = final_mdl;
        result.best_lambda = best_lambda;
        result.cv_loss = min_loss;
        result.confusion_matrix = cm;
        result.accuracy = accuracy;
        result.precision = precision;
        result.recall = recall;
        result.specificity = specificity;
        result.f1_score = f1_score;
        result.AUC = AUC;
        result.performance_grade = performance_grade;
        result.coefficients = coefficients;
        result.top_feature_indices = sorted_idx(1:min(10, length(sorted_idx)));
        result.top_features = valid_comp_cols(sorted_idx(1:min(10, length(sorted_idx))));
        result.top_coef_values = coefficients(sorted_idx(1:min(10, length(sorted_idx))));
        result.X_roc = X_roc;
        result.Y_roc = Y_roc;
    else
        result.model = [];
        result.accuracy = NaN;
        result.precision = NaN;
        result.recall = NaN;
        result.specificity = NaN;
        result.f1_score = NaN;
        result.AUC = NaN;
        result.performance_grade = 'Failed';
    end

    all_results{scenario_idx} = result;

    fprintf('════════════════════════════════════════════════════════\n\n');
end

%% ========================================================================
%                    PART 5: 시나리오 비교 및 시각화
% =========================================================================

fprintf('【STEP 8】 시나리오 간 성능 비교\n');
fprintf('════════════════════════════════════════════════════════\n\n');

% 비교 테이블 생성
comparison_table = table();
for i = 1:length(all_results)
    r = all_results{i};
    comparison_table.Scenario{i} = sprintf('시나리오 %d', i);
    comparison_table.N_Total(i) = r.n_total;
    comparison_table.N_High(i) = r.n_high;
    comparison_table.N_Low(i) = r.n_low;
    comparison_table.Imbalance(i) = r.imbalance_ratio;
    comparison_table.Accuracy(i) = r.accuracy;
    comparison_table.Precision(i) = r.precision;
    comparison_table.Recall(i) = r.recall;
    comparison_table.F1_Score(i) = r.f1_score;
    comparison_table.AUC(i) = r.AUC;
    comparison_table.Grade{i} = r.performance_grade;
end

fprintf('성능 비교 요약:\n\n');
disp(comparison_table);

% 최적 시나리오 추천
fprintf('\n【추천 시나리오】\n');
fprintf('────────────────────────────────────────────\n');

[max_auc, best_auc_idx] = max(comparison_table.AUC);
fprintf('✓ AUC 기준 최고: %s (AUC=%.4f, 등급: %s)\n', ...
    comparison_table.Scenario{best_auc_idx}, max_auc, comparison_table.Grade{best_auc_idx});

[max_f1, best_f1_idx] = max(comparison_table.F1_Score);
fprintf('✓ F1 Score 기준 최고: %s (F1=%.2f%%)\n', ...
    comparison_table.Scenario{best_f1_idx}, max_f1*100);

[min_imbalance, best_balance_idx] = min(comparison_table.Imbalance);
fprintf('✓ 클래스 균형 최고: %s (비율=%.2f:1)\n\n', ...
    comparison_table.Scenario{best_balance_idx}, min_imbalance);

%% 시각화
fprintf('【STEP 9】 시각화 생성\n');
fprintf('────────────────────────────────────────────\n');

% 그림 1: 성능 지표 비교
fig1 = figure('Position', [50, 50, 1400, 900], 'Color', 'white');

% Subplot 1: Accuracy
subplot(2, 3, 1);
bar(comparison_table.Accuracy * 100, 'FaceColor', [0.2 0.4 0.8], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('Accuracy (%)');
xlabel('시나리오');
title('정확도', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0 100]);
grid on;
set(gca, 'XTickLabel', 1:length(all_results));

% Subplot 2: Precision
subplot(2, 3, 2);
bar(comparison_table.Precision * 100, 'FaceColor', [0.8 0.4 0.2], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('Precision (%)');
xlabel('시나리오');
title('정밀도', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0 100]);
grid on;
set(gca, 'XTickLabel', 1:length(all_results));

% Subplot 3: Recall
subplot(2, 3, 3);
bar(comparison_table.Recall * 100, 'FaceColor', [0.2 0.8 0.4], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('Recall (%)');
xlabel('시나리오');
title('재현율', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0 100]);
grid on;
set(gca, 'XTickLabel', 1:length(all_results));

% Subplot 4: F1 Score
subplot(2, 3, 4);
bar(comparison_table.F1_Score * 100, 'FaceColor', [0.8 0.2 0.6], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('F1 Score (%)');
xlabel('시나리오');
title('F1 점수', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0 100]);
grid on;
set(gca, 'XTickLabel', 1:length(all_results));

% Subplot 5: AUC
subplot(2, 3, 5);
bar(comparison_table.AUC, 'FaceColor', [0.6 0.2 0.8], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('AUC');
xlabel('시나리오');
title('AUC-ROC', 'FontSize', 14, 'FontWeight', 'bold');
ylim([0.5 1.0]);
grid on;
set(gca, 'XTickLabel', 1:length(all_results));

% Subplot 6: 클래스 불균형
subplot(2, 3, 6);
bar(comparison_table.Imbalance, 'FaceColor', [0.5 0.5 0.5], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('불균형 비율');
xlabel('시나리오');
title('클래스 불균형', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
set(gca, 'XTickLabel', 1:length(all_results));

sgtitle('로지스틱 회귀 시나리오별 성능 비교', 'FontSize', 18, 'FontWeight', 'bold');

saveas(fig1, fullfile(config.output_dir, 'performance_comparison.png'));
fprintf('  ✓ 저장: performance_comparison.png\n');

% 그림 2: ROC Curve 비교
fig2 = figure('Position', [100, 100, 1000, 800], 'Color', 'white');
hold on;

colors = [0.2 0.4 0.8; 0.8 0.4 0.2; 0.2 0.8 0.4; 0.8 0.2 0.6];
legend_entries = {};

for i = 1:length(all_results)
    if isfield(all_results{i}, 'X_roc') && ~isempty(all_results{i}.X_roc)
        plot(all_results{i}.X_roc, all_results{i}.Y_roc, ...
            'LineWidth', 3, 'Color', colors(i, :));
        legend_entries{end+1} = sprintf('시나리오 %d (AUC=%.3f)', i, all_results{i}.AUC);
    end
end

% 랜덤 분류선
plot([0 1], [0 1], 'k--', 'LineWidth', 2);
legend_entries{end+1} = 'Random (AUC=0.5)';

xlabel('False Positive Rate', 'FontSize', 14);
ylabel('True Positive Rate', 'FontSize', 14);
title('ROC Curve 비교', 'FontSize', 16, 'FontWeight', 'bold');
legend(legend_entries, 'Location', 'southeast', 'FontSize', 11);
grid on;
hold off;

saveas(fig2, fullfile(config.output_dir, 'roc_curves_comparison.png'));
fprintf('  ✓ 저장: roc_curves_comparison.png\n');

% 그림 3: 상위 역량 비교
fig3 = figure('Position', [150, 150, 1600, 900], 'Color', 'white');

for i = 1:length(all_results)
    if isfield(all_results{i}, 'top_features') && ~isempty(all_results{i}.top_features)
        subplot(2, 2, i);

        top_n = min(8, length(all_results{i}.top_features));
        features = all_results{i}.top_features(1:top_n);
        coefs = all_results{i}.top_coef_values(1:top_n);

        barh(1:top_n, coefs, 'FaceColor', 'flat', 'EdgeColor', 'k', 'LineWidth', 1.2);
        set(gca, 'YTick', 1:top_n, 'YTickLabel', features);
        xlabel('계수 (Coefficient)', 'FontSize', 11);
        title(sprintf('시나리오 %d: 주요 역량', i), 'FontSize', 13, 'FontWeight', 'bold');
        grid on;

        % 0 기준선
        hold on;
        plot([0 0], [0.5 top_n+0.5], 'k--', 'LineWidth', 1.5);
        hold off;
    end
end

sgtitle('시나리오별 주요 역량 비교 (계수 기준)', 'FontSize', 18, 'FontWeight', 'bold');

saveas(fig3, fullfile(config.output_dir, 'top_features_comparison.png'));
fprintf('  ✓ 저장: top_features_comparison.png\n');

%% 엑셀 저장
fprintf('\n【STEP 10】 엑셀 결과 저장\n');
fprintf('────────────────────────────────────────────\n');

excel_file = fullfile(config.output_dir, 'scenario_comparison_results.xlsx');

% 시트 1: 성능 비교 요약
writetable(comparison_table, excel_file, 'Sheet', '성능비교');

% 시트 2-5: 각 시나리오 상세
for i = 1:length(all_results)
    r = all_results{i};

    % 시나리오 정보
    info_items = {'시나리오명'; '설명'; '고성과자'; '저성과자'; '제외'; ...
        '총샘플'; '고성과자수'; '저성과자수'; '불균형비율'; ...
        'Accuracy'; 'Precision'; 'Recall'; 'F1 Score'; 'AUC'; '성능등급'};
    info_values = {
        r.scenario_name;
        r.scenario_description;
        strjoin(r.high_performers, ', ');
        strjoin(r.low_performers, ', ');
        strjoin(r.excluded, ', ');
        r.n_total;
        r.n_high;
        r.n_low;
        r.imbalance_ratio;
        r.accuracy;
        r.precision;
        r.recall;
        r.f1_score;
        r.AUC;
        r.performance_grade
    };
    info_table = table(info_items, info_values);

    writetable(info_table, excel_file, 'Sheet', sprintf('시나리오%d_요약', i));

    % 주요 역량
    if isfield(r, 'top_features') && ~isempty(r.top_features)
        % cell array를 column vector로 변환
        comp_names = r.top_features(:);  % 이미 column이면 그대로, row면 column으로
        comp_coefs = r.top_coef_values(:);
        feature_table = table(comp_names, comp_coefs, 'VariableNames', {'역량명', '계수'});
        writetable(feature_table, excel_file, 'Sheet', sprintf('시나리오%d_주요역량', i));
    end
end

fprintf('  ✓ 저장: scenario_comparison_results.xlsx\n');

%% 최종 요약
fprintf('\n════════════════════════════════════════════════════════\n');
fprintf('   분석 완료\n');
fprintf('════════════════════════════════════════════════════════\n');
fprintf('✓ 총 시나리오: %d개\n', length(scenarios));
fprintf('✓ 결과 저장 위치: %s\n\n', config.output_dir);
fprintf('생성된 파일:\n');
fprintf('  - performance_comparison.png (성능 지표 비교)\n');
fprintf('  - roc_curves_comparison.png (ROC 곡선 비교)\n');
fprintf('  - top_features_comparison.png (주요 역량 비교)\n');
fprintf('  - scenario_comparison_results.xlsx (상세 결과)\n\n');
fprintf('추천 시나리오: %s\n', comparison_table.Scenario{best_auc_idx});
fprintf('  AUC: %.4f\n', max_auc);
fprintf('  F1 Score: %.2f%%\n', comparison_table.F1_Score(best_auc_idx)*100);
fprintf('  성능 등급: %s\n', comparison_table.Grade{best_auc_idx});
fprintf('════════════════════════════════════════════════════════\n');
