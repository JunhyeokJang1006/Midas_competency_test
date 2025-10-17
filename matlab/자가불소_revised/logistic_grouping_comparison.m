% ========================================================================
%          로지스틱 회귀 기반 고성과자/저성과자 그룹핑 비교 분석
% ========================================================================
% 목적: 다양한 그룹핑 시나리오를 비교하여 최적의 분류 기준 도출
%
% 분석 시나리오:
% 1. 자연성 vs 나머지
% 2. 자연성+유익한불연성 vs 나머지
% 3. 자연성+유익한불연성+성실한가연성 vs 무능한불연성+소화성+게으른가연성 (유능한불연성 제외)
% 4. 자연성+유익한불연성 vs 무능한불연성+소화성+게으른가연성 (유능한불연성+성실한가연성 제외)
%
% ========================================================================

clear; clc; close all;
rng(42, 'twister');  % 재현성 보장

%% ========================================================================
%                          PART 1: 초기 설정 및 데이터 로딩
% =========================================================================

fprintf('=======================================================\n');
fprintf('   로지스틱 회귀 기반 그룹핑 비교 분석 시작\n');
fprintf('=======================================================\n\n');

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent';

% 출력 디렉토리 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
    fprintf('✓ 출력 디렉토리 생성: %s\n\n', config.output_dir);
end

%% 데이터 로딩
fprintf('【STEP 1】 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

try
    % HR 데이터 로딩
    fprintf('▶ HR 데이터 로딩 중...\n');
    hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
    fprintf('  ✓ HR 데이터: %d명 로드 완료\n', height(hr_data));

    % 역량검사 데이터 로딩
    fprintf('▶ 역량검사 데이터 로딩 중...\n');
    comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
    comp_total = readtable(config.comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 상위항목 데이터: %d명\n', height(comp_upper));
    fprintf('  ✓ 종합점수 데이터: %d명\n\n', height(comp_total));
catch ME
    error('데이터 로딩 실패: %s', ME.message);
end

%% 신뢰가능성 필터링
fprintf('【STEP 2】 신뢰가능성 필터링\n');
fprintf('────────────────────────────────────────────\n');

reliability_col_idx = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col_idx)
    reliability_data = comp_upper{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(comp_upper), 1);
    end

    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));
    comp_upper = comp_upper(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능한 데이터: %d명\n\n', height(comp_upper));
else
    fprintf('  ⚠ 신뢰가능성 컬럼이 없습니다. 모든 데이터를 사용합니다.\n\n');
end

%% 데이터 매칭
fprintf('【STEP 3】 HR 데이터와 역량검사 데이터 매칭\n');
fprintf('────────────────────────────────────────────\n');

% ID 컬럼 찾기
id_col_hr = find(contains(hr_data.Properties.VariableNames, {'ID', 'id', '사번'}), 1);
id_col_comp = find(contains(comp_upper.Properties.VariableNames, {'ID', 'id', '사번'}), 1);

if isempty(id_col_hr) || isempty(id_col_comp)
    error('ID 컬럼을 찾을 수 없습니다.');
end

% 공통 ID 찾기
hr_ids = hr_data{:, id_col_hr};
comp_ids = comp_upper{:, id_col_comp};

% 매칭되는 데이터만 추출
[common_ids, idx_hr, idx_comp] = intersect(hr_ids, comp_ids, 'stable');
matched_hr = hr_data(idx_hr, :);
matched_comp_upper = comp_upper(idx_comp, :);

fprintf('  ✓ 매칭된 데이터: %d명\n\n', length(common_ids));

%% 인재유형 컬럼 찾기
talent_type_col = find(contains(matched_hr.Properties.VariableNames, {'인재유형', '인재', 'talent'}), 1);
if isempty(talent_type_col)
    error('인재유형 컬럼을 찾을 수 없습니다.');
end

matched_talent_types = matched_hr{:, talent_type_col};
if iscell(matched_talent_types)
    matched_talent_types = cellfun(@string, matched_talent_types);
else
    matched_talent_types = string(matched_talent_types);
end

% 역량 점수 데이터 추출
comp_col_start = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1) + 1;
if isempty(comp_col_start)
    comp_col_start = 2; % ID 다음부터
end

valid_comp_cols = comp_upper.Properties.VariableNames(comp_col_start:end);
matched_comp = matched_comp_upper(:, comp_col_start:end);
X_raw = table2array(matched_comp);

fprintf('【STEP 4】 역량 데이터 확인\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  샘플 수: %d\n', size(X_raw, 1));
fprintf('  역량 수: %d\n', size(X_raw, 2));
fprintf('  결측값 비율: %.1f%%\n\n', sum(isnan(X_raw(:)))/numel(X_raw)*100);

%% ========================================================================
%                    PART 2: 그룹핑 시나리오 정의 및 분석
% =========================================================================

fprintf('【STEP 5】 그룹핑 시나리오 정의\n');
fprintf('────────────────────────────────────────────\n');

% 4가지 그룹핑 시나리오 정의
scenarios = {};

% 시나리오 1: 자연성 vs 나머지
scenarios{1} = struct(...
    'name', '시나리오 1: 자연성 vs 나머지', ...
    'high_performers', {{'자연성'}}, ...
    'low_performers', {{'성실한 가연성', '유익한 불연성', '유능한 불연성', '게으른 가연성', '무능한 불연성', '소화성'}}, ...
    'excluded', {{''}});

% 시나리오 2: 자연성+유익한불연성 vs 나머지
scenarios{2} = struct(...
    'name', '시나리오 2: 자연성+유익한불연성 vs 나머지', ...
    'high_performers', {{'자연성', '유익한 불연성'}}, ...
    'low_performers', {{'성실한 가연성', '유능한 불연성', '게으른 가연성', '무능한 불연성', '소화성'}}, ...
    'excluded', {{''}});

% 시나리오 3: 자연성+유익한불연성+성실한가연성 vs 무능한불연성+소화성+게으른가연성 (유능한불연성 제외)
scenarios{3} = struct(...
    'name', '시나리오 3: 상위3 vs 하위3 (유능한불연성 제외)', ...
    'high_performers', {{'자연성', '유익한 불연성', '성실한 가연성'}}, ...
    'low_performers', {{'무능한 불연성', '소화성', '게으른 가연성'}}, ...
    'excluded', {{'유능한 불연성'}});

% 시나리오 4: 자연성+유익한불연성 vs 무능한불연성+소화성+게으른가연성 (유능한불연성+성실한가연성 제외)
scenarios{4} = struct(...
    'name', '시나리오 4: 상위2 vs 하위3 (중위2 제외)', ...
    'high_performers', {{'자연성', '유익한 불연성'}}, ...
    'low_performers', {{'무능한 불연성', '소화성', '게으른 가연성'}}, ...
    'excluded', {{'유능한 불연성', '성실한 가연성'}});

% 시나리오 정보 출력
for i = 1:length(scenarios)
    fprintf('\n%s\n', scenarios{i}.name);
    fprintf('  고성과자: %s\n', strjoin(scenarios{i}.high_performers, ', '));
    fprintf('  저성과자: %s\n', strjoin(scenarios{i}.low_performers, ', '));
    if ~isempty(scenarios{i}.excluded{1})
        fprintf('  제외: %s\n', strjoin(scenarios{i}.excluded, ', '));
    end
end
fprintf('\n');

%% ========================================================================
%                    PART 3: 각 시나리오별 로지스틱 회귀 분석
% =========================================================================

fprintf('【STEP 6】 시나리오별 로지스틱 회귀 분석 실행\n');
fprintf('========================================================\n\n');

% 결과 저장용 구조체
results = struct();

for scenario_idx = 1:length(scenarios)
    scenario = scenarios{scenario_idx};

    fprintf('┌──────────────────────────────────────────────────────┐\n');
    fprintf('│ %s\n', scenario.name);
    fprintf('└──────────────────────────────────────────────────────┘\n\n');

    % 레이블 생성
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
    X_valid = X_raw(valid_idx, :);
    y_valid = y_binary(valid_idx);

    % 결측값 처리 (평균 대체)
    for col = 1:size(X_valid, 2)
        col_data = X_valid(:, col);
        nan_idx = isnan(col_data);
        if any(nan_idx)
            col_mean = mean(col_data(~nan_idx));
            X_valid(nan_idx, col) = col_mean;
        end
    end

    % 그룹별 샘플 수
    n_high = sum(y_valid == 1);
    n_low = sum(y_valid == 0);
    n_total = length(y_valid);

    fprintf('▶ 데이터 요약\n');
    fprintf('  고성과자: %d명 (%.1f%%)\n', n_high, n_high/n_total*100);
    fprintf('  저성과자: %d명 (%.1f%%)\n', n_low, n_low/n_total*100);
    fprintf('  총 샘플: %d명\n\n', n_total);

    % 클래스 불균형 확인
    imbalance_ratio = max(n_high, n_low) / min(n_high, n_low);
    fprintf('  클래스 불균형 비율: %.2f:1\n', imbalance_ratio);
    if imbalance_ratio > 3
        fprintf('  ⚠ 심각한 클래스 불균형\n\n');
    elseif imbalance_ratio > 1.5
        fprintf('  △ 중간 정도의 클래스 불균형\n\n');
    else
        fprintf('  ✓ 클래스 균형\n\n');
    end

    %% 마할라노비스 거리 계산 (그룹 분리 타당성)
    fprintf('▶ 그룹 분리 타당성 검증 (마할라노비스 거리)\n');

    X_high = X_valid(y_valid == 1, :);
    X_low = X_valid(y_valid == 0, :);

    if n_high >= 3 && n_low >= 3
        % 공통 공분산 행렬 계산
        cov_high = cov(X_high);
        cov_low = cov(X_low);
        pooled_cov = ((n_high - 1) * cov_high + (n_low - 1) * cov_low) / (n_high + n_low - 2);

        % 평균 차이
        mean_high = mean(X_high);
        mean_low = mean(X_low);
        mean_diff = mean_high - mean_low;

        % 마할라노비스 거리 계산
        pooled_cov_reg = pooled_cov + eye(size(pooled_cov)) * 1e-6;
        try
            L = chol(pooled_cov_reg, 'lower');
            v = L \ mean_diff';
            mahal_distance = sqrt(v' * v);

            fprintf('  마할라노비스 거리: %.4f\n', mahal_distance);

            if mahal_distance >= 3.0
                fprintf('  ✓ 매우 우수한 그룹 분리\n');
            elseif mahal_distance >= 2.0
                fprintf('  ✓ 우수한 그룹 분리\n');
            elseif mahal_distance >= 1.5
                fprintf('  △ 보통 수준의 그룹 분리\n');
            else
                fprintf('  ⚠ 약한 그룹 분리\n');
            end
        catch
            mahal_distance = NaN;
            fprintf('  ⚠ 마할라노비스 거리 계산 실패\n');
        end
    else
        mahal_distance = NaN;
        fprintf('  ⚠ 샘플 수가 부족하여 계산 불가\n');
    end
    fprintf('\n');

    %% 로지스틱 회귀 모델 학습 (5-Fold Cross-Validation)
    fprintf('▶ 로지스틱 회귀 모델 학습 (5-Fold CV)\n');

    try
        % 교차검증 설정
        cv = cvpartition(y_valid, 'KFold', 5);

        % 모델 학습 (Ridge 정규화)
        mdl = fitclinear(X_valid, y_valid, ...
            'Learner', 'logistic', ...
            'Regularization', 'ridge', ...
            'CVPartition', cv);

        % 교차검증 정확도
        cv_accuracy = 1 - kfoldLoss(mdl);

        % 전체 데이터로 최종 모델 학습
        final_mdl = fitclinear(X_valid, y_valid, ...
            'Learner', 'logistic', ...
            'Regularization', 'ridge', ...
            'Lambda', mdl.Lambda);

        % 예측 및 평가
        [y_pred, scores] = predict(final_mdl, X_valid);

        % 혼동행렬
        cm = confusionmat(y_valid, y_pred);

        % 성능 지표 계산
        TP = cm(2, 2);  % True Positive
        TN = cm(1, 1);  % True Negative
        FP = cm(1, 2);  % False Positive
        FN = cm(2, 1);  % False Negative

        accuracy = (TP + TN) / (TP + TN + FP + FN);
        precision = TP / (TP + FP);
        recall = TP / (TP + FN);
        f1_score = 2 * (precision * recall) / (precision + recall);
        specificity = TN / (TN + FP);

        % AUC 계산
        [X_roc, Y_roc, T_roc, AUC] = perfcurve(y_valid, scores(:,2), 1);

        fprintf('  ✓ 모델 학습 완료\n');
        fprintf('  CV 정확도: %.2f%%\n', cv_accuracy * 100);
        fprintf('  최종 정확도: %.2f%%\n', accuracy * 100);
        fprintf('  Precision: %.2f%%\n', precision * 100);
        fprintf('  Recall (Sensitivity): %.2f%%\n', recall * 100);
        fprintf('  F1 Score: %.2f%%\n', f1_score * 100);
        fprintf('  Specificity: %.2f%%\n', specificity * 100);
        fprintf('  AUC: %.4f\n', AUC);

        % 변수 중요도 (계수의 절대값)
        coefficients = final_mdl.Beta;
        [sorted_coef, sorted_idx] = sort(abs(coefficients), 'descend');

        fprintf('\n  상위 5개 중요 역량:\n');
        for i = 1:min(5, length(sorted_coef))
            fprintf('    %d. %s: %.4f\n', i, valid_comp_cols{sorted_idx(i)}, coefficients(sorted_idx(i)));
        end

    catch ME
        fprintf('  ✗ 모델 학습 실패: %s\n', ME.message);
        accuracy = NaN;
        cv_accuracy = NaN;
        precision = NaN;
        recall = NaN;
        f1_score = NaN;
        AUC = NaN;
        mahal_distance = NaN;
    end

    fprintf('\n');

    % 결과 저장
    results(scenario_idx).scenario_name = scenario.name;
    results(scenario_idx).n_high = n_high;
    results(scenario_idx).n_low = n_low;
    results(scenario_idx).n_total = n_total;
    results(scenario_idx).imbalance_ratio = imbalance_ratio;
    results(scenario_idx).mahal_distance = mahal_distance;
    results(scenario_idx).cv_accuracy = cv_accuracy;
    results(scenario_idx).accuracy = accuracy;
    results(scenario_idx).precision = precision;
    results(scenario_idx).recall = recall;
    results(scenario_idx).f1_score = f1_score;
    results(scenario_idx).AUC = AUC;

    if exist('final_mdl', 'var')
        results(scenario_idx).model = final_mdl;
        results(scenario_idx).coefficients = coefficients;
        results(scenario_idx).top_features = valid_comp_cols(sorted_idx(1:min(10, length(sorted_idx))));
        results(scenario_idx).top_coef = coefficients(sorted_idx(1:min(10, length(sorted_idx))));
    end

    fprintf('════════════════════════════════════════════════════════\n\n');
end

%% ========================================================================
%                    PART 4: 시나리오 비교 및 시각화
% =========================================================================

fprintf('【STEP 7】 시나리오 비교 결과\n');
fprintf('========================================================\n\n');

% 결과 테이블 생성
comparison_table = table();
for i = 1:length(results)
    comparison_table.Scenario{i} = sprintf('시나리오 %d', i);
    comparison_table.High_N(i) = results(i).n_high;
    comparison_table.Low_N(i) = results(i).n_low;
    comparison_table.Imbalance(i) = results(i).imbalance_ratio;
    comparison_table.Mahal_Dist(i) = results(i).mahal_distance;
    comparison_table.CV_Accuracy(i) = results(i).cv_accuracy;
    comparison_table.Accuracy(i) = results(i).accuracy;
    comparison_table.Precision(i) = results(i).precision;
    comparison_table.Recall(i) = results(i).recall;
    comparison_table.F1_Score(i) = results(i).f1_score;
    comparison_table.AUC(i) = results(i).AUC;
end

disp(comparison_table);

%% 최적 시나리오 추천
fprintf('\n【추천 시나리오】\n');
fprintf('────────────────────────────────────────────\n');

% F1 스코어 기준 최적 시나리오
[max_f1, best_idx] = max(comparison_table.F1_Score);
fprintf('✓ F1 Score 기준 최적: %s (F1=%.2f%%)\n', comparison_table.Scenario{best_idx}, max_f1*100);

% AUC 기준 최적 시나리오
[max_auc, best_auc_idx] = max(comparison_table.AUC);
fprintf('✓ AUC 기준 최적: %s (AUC=%.4f)\n', comparison_table.Scenario{best_auc_idx}, max_auc);

% 마할라노비스 거리 기준
[max_mahal, best_mahal_idx] = max(comparison_table.Mahal_Dist);
fprintf('✓ 그룹 분리도 기준 최적: %s (D=%.4f)\n\n', comparison_table.Scenario{best_mahal_idx}, max_mahal);

%% 시각화: 성능 지표 비교
fprintf('【STEP 8】 시각화 생성\n');
fprintf('────────────────────────────────────────────\n');

% 그림 1: 성능 지표 비교 바 차트
fig1 = figure('Position', [100, 100, 1200, 600], 'Color', 'white');

metrics = {'Accuracy', 'Precision', 'Recall', 'F1_Score', 'AUC'};
metric_names = {'정확도', 'Precision', 'Recall', 'F1 Score', 'AUC'};
colors = [0.2 0.4 0.8; 0.8 0.4 0.2; 0.2 0.8 0.4; 0.8 0.2 0.6; 0.6 0.2 0.8];

for m = 1:length(metrics)
    subplot(2, 3, m);

    metric_values = comparison_table.(metrics{m});

    % AUC는 스케일이 다르므로 별도 처리
    if strcmp(metrics{m}, 'AUC')
        b = bar(metric_values, 'FaceColor', colors(m, :), 'EdgeColor', 'k', 'LineWidth', 1.5);
        ylim([0.5, 1.0]);
    else
        b = bar(metric_values * 100, 'FaceColor', colors(m, :), 'EdgeColor', 'k', 'LineWidth', 1.5);
        ylim([0, 100]);
    end

    % 값 표시
    for i = 1:length(metric_values)
        if strcmp(metrics{m}, 'AUC')
            text(i, metric_values(i) + 0.02, sprintf('%.3f', metric_values(i)), ...
                'HorizontalAlignment', 'center', 'FontSize', 10, 'FontWeight', 'bold');
        else
            text(i, metric_values(i)*100 + 2, sprintf('%.1f%%', metric_values(i)*100), ...
                'HorizontalAlignment', 'center', 'FontSize', 10, 'FontWeight', 'bold');
        end
    end

    xlabel('시나리오');
    if strcmp(metrics{m}, 'AUC')
        ylabel(metric_names{m});
    else
        ylabel([metric_names{m}, ' (%)']);
    end
    title(metric_names{m}, 'FontSize', 14, 'FontWeight', 'bold');
    grid on;
    set(gca, 'XTickLabel', 1:4);
end

% 그림 2: 마할라노비스 거리 및 클래스 불균형
subplot(2, 3, 6);
yyaxis left
bar(comparison_table.Mahal_Dist, 'FaceColor', [0.3 0.6 0.8], 'EdgeColor', 'k', 'LineWidth', 1.5);
ylabel('마할라노비스 거리');
ylim([0, max(comparison_table.Mahal_Dist) * 1.2]);

yyaxis right
plot(1:4, comparison_table.Imbalance, '-o', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'MarkerFaceColor', [0.8 0.3 0.3], 'Color', [0.8 0.3 0.3]);
ylabel('클래스 불균형 비율');

xlabel('시나리오');
title('그룹 분리도 & 클래스 불균형', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
legend({'마할라노비스 거리', '불균형 비율'}, 'Location', 'best');

sgtitle('로지스틱 회귀 그룹핑 시나리오 비교', 'FontSize', 16, 'FontWeight', 'bold');

% 저장
saveas(fig1, fullfile(config.output_dir, 'logistic_scenario_comparison.png'));
fprintf('  ✓ 시각화 저장: logistic_scenario_comparison.png\n');

%% 그림 2: 주요 역량 비교 (시나리오별 상위 5개)
fig2 = figure('Position', [150, 150, 1400, 800], 'Color', 'white');

for scenario_idx = 1:length(results)
    if isfield(results(scenario_idx), 'top_features') && ~isempty(results(scenario_idx).top_features)
        subplot(2, 2, scenario_idx);

        top_n = min(5, length(results(scenario_idx).top_features));
        top_features = results(scenario_idx).top_features(1:top_n);
        top_coef = results(scenario_idx).top_coef(1:top_n);

        % 양수/음수에 따라 색상 구분
        bar_colors = zeros(top_n, 3);
        for i = 1:top_n
            if top_coef(i) > 0
                bar_colors(i, :) = [0.2, 0.6, 0.8];  % 파란색 (고성과자에 긍정적)
            else
                bar_colors(i, :) = [0.8, 0.4, 0.2];  % 주황색 (고성과자에 부정적)
            end
        end

        b = barh(1:top_n, top_coef, 'FaceColor', 'flat', 'EdgeColor', 'k', 'LineWidth', 1.2);
        b.CData = bar_colors;

        set(gca, 'YTick', 1:top_n, 'YTickLabel', top_features);
        xlabel('계수 (Coefficient)');
        title(sprintf('시나리오 %d: 상위 5개 중요 역량', scenario_idx), 'FontSize', 12, 'FontWeight', 'bold');
        grid on;

        % 계수 값 표시
        for i = 1:top_n
            if top_coef(i) > 0
                text(top_coef(i) + 0.01, i, sprintf('%.3f', top_coef(i)), ...
                    'FontSize', 9, 'VerticalAlignment', 'middle');
            else
                text(top_coef(i) - 0.01, i, sprintf('%.3f', top_coef(i)), ...
                    'FontSize', 9, 'VerticalAlignment', 'middle', 'HorizontalAlignment', 'right');
            end
        end
    end
end

sgtitle('시나리오별 중요 역량 비교', 'FontSize', 16, 'FontWeight', 'bold');

% 저장
saveas(fig2, fullfile(config.output_dir, 'logistic_top_features_comparison.png'));
fprintf('  ✓ 시각화 저장: logistic_top_features_comparison.png\n');

%% 엑셀 파일로 결과 저장
fprintf('\n【STEP 9】 결과 엑셀 파일 저장\n');
fprintf('────────────────────────────────────────────\n');

excel_file = fullfile(config.output_dir, 'logistic_scenario_comparison.xlsx');

% 시트 1: 성능 비교 요약
writetable(comparison_table, excel_file, 'Sheet', '성능비교요약');

% 시트 2-5: 각 시나리오별 상위 역량
for i = 1:length(results)
    if isfield(results(i), 'top_features') && ~isempty(results(i).top_features)
        feature_table = table(results(i).top_features', results(i).top_coef', ...
            'VariableNames', {'역량명', '계수'});
        writetable(feature_table, excel_file, 'Sheet', sprintf('시나리오%d_주요역량', i));
    end
end

fprintf('  ✓ 엑셀 파일 저장: %s\n', excel_file);

%% 최종 요약
fprintf('\n========================================================\n');
fprintf('   분석 완료\n');
fprintf('========================================================\n');
fprintf('✓ 총 %d개 시나리오 비교 완료\n', length(scenarios));
fprintf('✓ 결과 파일 저장 위치: %s\n', config.output_dir);
fprintf('  - logistic_scenario_comparison.png (성능 비교 차트)\n');
fprintf('  - logistic_top_features_comparison.png (중요 역량 비교)\n');
fprintf('  - logistic_scenario_comparison.xlsx (상세 결과)\n');
fprintf('\n권장 시나리오: %s (F1=%.2f%%, AUC=%.4f)\n', ...
    comparison_table.Scenario{best_idx}, max_f1*100, comparison_table.AUC(best_idx));
fprintf('========================================================\n');
