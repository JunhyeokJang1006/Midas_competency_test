%% Final Talent Type Analysis with Radar Chart
% Based on reference codes: corr_CSR_vs_comp_test_cursor.m and mk_rader_chart_wih_statistics.m
% Correct method: Using proper sheet names and creating radar charts for talent types

clear; clc; close all;

%% Global Settings for High-Quality Figures
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 11);
set(0, 'DefaultTextFontSize', 11);
set(0, 'DefaultLineLineWidth', 1.5);

%% 1. Data Loading - Correct Method (Following CSR Reference)
fprintf('========================================\n');
fprintf('📊 Final Talent Type Analysis Start\n');
fprintf('========================================\n\n');

% File paths
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    % HR data with talent types
    fprintf('▶ Loading HR data...\n');
    hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
    fprintf('  ✓ HR data: %d people\n', height(hr_data));

    % Competency test data - Upper level categories (CORRECT METHOD)
    fprintf('▶ Loading competency test data...\n');
    comp_upper = readtable(comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ Upper level competency data: %d people\n', height(comp_upper));

    % Competency test data - Total scores
    comp_total = readtable(comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ Total score data: %d people\n', height(comp_total));

catch ME
    fprintf('✗ Data loading failed: %s\n', ME.message);
    return;
end

%% 2. Clean and Extract Talent Type Data
fprintf('\n=== Talent Type Data Extraction ===\n');

% Find talent type column (following CSR reference pattern)
talent_col_idx = [];
col_names = hr_data.Properties.VariableNames;
for col = 1:width(hr_data)
    colName = col_names{col};
    if contains(colName, '인재유형') || contains(colName, '인재') || contains(colName, '유형')
        talent_col_idx = col;
        break;
    end
end

if isempty(talent_col_idx)
    fprintf('✗ Talent type column not found\n');
    return;
end

talent_col_name = col_names{talent_col_idx};
fprintf('Found talent type column: %s\n', talent_col_name);

% Clean talent type data
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_name}), :);
fprintf('Employees with talent type: %d\n', height(hr_clean));

% Talent type distribution
talent_types = hr_clean{:, talent_col_name};
unique_types = unique(talent_types(~cellfun(@isempty, talent_types)));

fprintf('\nTalent type distribution:\n');
for i = 1:length(unique_types)
    count = sum(strcmp(talent_types, unique_types{i}));
    fprintf('  %s: %d people\n', unique_types{i}, count);
end

%% 3. Process Competency Data (Following CSR Reference)
fprintf('\n=== Competency Data Processing ===\n');

% Find ID column in competency data
comp_id_col = [];
comp_col_names = comp_upper.Properties.VariableNames;
for col = 1:width(comp_upper)
    colName = lower(comp_col_names{col});
    colData = comp_upper{:, col};
    if contains(colName, {'id', '사번'}) && isnumeric(colData) && ~all(isnan(colData))
        comp_id_col = col;
        break;
    end
end

if isempty(comp_id_col)
    fprintf('✗ ID column not found in competency data\n');
    return;
end

fprintf('Found competency ID column: %s\n', comp_col_names{comp_id_col});

% Extract competency scores (starting from reasonable column, avoiding metadata)
fprintf('▶ Extracting competency scores...\n');

valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)  % Start from column 6 to avoid metadata
    col_name = comp_col_names{i};
    col_data = comp_upper{:, i};

    % Select numeric columns with valid data (following CSR pattern)
    if isnumeric(col_data) && ~all(isnan(col_data)) && var(col_data, 'omitnan') > 0
        valid_data = col_data(~isnan(col_data));
        if length(valid_data) >= 5 && all(valid_data >= 0) && all(valid_data <= 100)
            valid_comp_cols{end+1} = col_name;
            valid_comp_indices = [valid_comp_indices, i];
            fprintf('  ✓ %s: range [%.1f, %.1f], valid %d people\n', ...
                col_name, min(valid_data), max(valid_data), length(valid_data));
        end
    end
end

if isempty(valid_comp_cols)
    fprintf('✗ No valid competency columns found\n');
    return;
end

fprintf('Valid competency categories: %d\n', length(valid_comp_cols));

%% 4. Match IDs and Extract Data
fprintf('\n=== ID Matching and Data Extraction ===\n');

% Standardize IDs (following CSR reference)
hr_ids = hr_clean.ID;
comp_ids = comp_upper{:, comp_id_col};

% Convert to string for matching
hr_ids_str = arrayfun(@(x) sprintf('%.0f', x), hr_ids, 'UniformOutput', false);
comp_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_ids, 'UniformOutput', false);

[matched_ids, hr_idx, comp_idx] = intersect(hr_ids_str, comp_ids_str);

if length(matched_ids) < 10
    fprintf('✗ Insufficient matched IDs: %d\n', length(matched_ids));
    return;
end

fprintf('Successfully matched IDs: %d people\n', length(matched_ids));

% Extract matched data
matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, valid_comp_indices);
matched_talent_types = matched_hr{:, talent_col_name};

fprintf('Final analysis dataset: %d people\n', length(matched_ids));

%% 5. Calculate Total Scores (Following CSR Pattern)
fprintf('\n=== Total Score Integration ===\n');

% Also match with total scores
comp_total_ids = comp_total{:, 1};  % Assume first column is ID
comp_total_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_total_ids, 'UniformOutput', false);

[total_matched_ids, ~, total_comp_idx] = intersect(matched_ids, comp_total_ids_str);

if ~isempty(total_matched_ids)
    % Find total score column (usually last column)
    total_score_col = width(comp_total);
    total_scores = comp_total{total_comp_idx, total_score_col};

    fprintf('Integrated total scores: %d people\n', length(total_matched_ids));
    fprintf('Total score column: %s\n', comp_total.Properties.VariableNames{total_score_col});
else
    fprintf('⚠ No total score integration\n');
    total_scores = [];
end

%% 6. Talent Type Performance Analysis
fprintf('\n=== Talent Type Performance Analysis ===\n');

% Define performance ranking (user provided)
performance_ranking = containers.Map();
performance_ranking('자연성') = 8;
performance_ranking('성실한 가연성') = 7;
performance_ranking('유익한 불연성') = 6;
performance_ranking('유능한 불연성') = 5;
performance_ranking('게으른 가연성') = 4;
performance_ranking('무능한 불연성') = 3;
performance_ranking('위장형 소화성') = 2;
performance_ranking('소화성') = 1;

% Calculate statistics by talent type
unique_matched_types = unique(matched_talent_types(~cellfun(@isempty, matched_talent_types)));
type_stats = [];

fprintf('Competency statistics by talent type:\n');
fprintf('Talent Type           | Count | Comp Mean | Comp Std  | Total Mean | Performance\n');
fprintf('----------------------|-------|-----------|-----------|------------|------------\n');

for i = 1:length(unique_matched_types)
    talent_type = unique_matched_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);

    % Count
    count = sum(type_idx);

    % Competency scores statistics
    type_comp_data = matched_comp{type_idx, :};
    comp_mean = nanmean(type_comp_data(:));
    comp_std = nanstd(type_comp_data(:));

    % Total score statistics (if available)
    if ~isempty(total_scores)
        % Find corresponding indices in total scores
        type_ids = matched_ids(type_idx);
        [~, ~, total_type_idx] = intersect(type_ids, total_matched_ids);
        if ~isempty(total_type_idx)
            type_total_scores = total_scores(total_type_idx);
            total_mean = nanmean(type_total_scores);
        else
            total_mean = NaN;
        end
    else
        total_mean = NaN;
    end

    % Performance ranking
    if performance_ranking.isKey(talent_type)
        perf_rank = performance_ranking(talent_type);
    else
        perf_rank = NaN;
    end

    fprintf('%-21s | %5d | %9.2f | %9.2f | %10.2f | %11.0f\n', ...
            talent_type, count, comp_mean, comp_std, total_mean, perf_rank);

    type_stats = [type_stats; {talent_type, count, comp_mean, comp_std, total_mean, perf_rank}];
end

%% 7. Competency Weight Analysis (Top Predictors)
fprintf('\n=== Competency Weight Analysis ===\n');

% Calculate performance scores for matched data
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    talent_type = matched_talent_types{i};
    if performance_ranking.isKey(talent_type)
        performance_scores(i) = performance_ranking(talent_type);
    end
end

% Calculate correlation and importance for each competency
competency_importance = [];
competency_correlations = [];

fprintf('Competency importance for performance prediction:\n');
fprintf('Competency          | Correlation | High Mean | Low Mean  | Difference\n');
fprintf('--------------------|-------------|-----------|-----------|------------\n');

for j = 1:length(valid_comp_cols)
    comp_name = valid_comp_cols{j};
    comp_scores = matched_comp{:, j};

    % Valid data indices
    valid_idx = ~isnan(comp_scores) & performance_scores > 0;

    if sum(valid_idx) >= 10
        valid_comp_scores = comp_scores(valid_idx);
        valid_perf_scores = performance_scores(valid_idx);

        % Calculate correlation
        if var(valid_comp_scores) > 0 && var(valid_perf_scores) > 0
            correlation = corr(valid_comp_scores, valid_perf_scores);
        else
            correlation = 0;
        end

        % High vs Low performance comparison
        high_perf_threshold = median(valid_perf_scores);
        high_idx = valid_perf_scores > high_perf_threshold;
        low_idx = valid_perf_scores <= high_perf_threshold;

        high_mean = mean(valid_comp_scores(high_idx));
        low_mean = mean(valid_comp_scores(low_idx));
        difference = high_mean - low_mean;

        fprintf('%-19s | %11.3f | %9.2f | %9.2f | %10.2f\n', ...
                comp_name, correlation, high_mean, low_mean, difference);

        competency_correlations = [competency_correlations; correlation];
        competency_importance = [competency_importance; abs(correlation)];
    else
        competency_correlations = [competency_correlations; 0];
        competency_importance = [competency_importance; 0];
    end
end

% Get top 5 competencies
[sorted_importance, sort_idx] = sort(competency_importance, 'descend');
top_5_idx = sort_idx(1:min(5, length(sort_idx)));
top_5_competencies = valid_comp_cols(top_5_idx);
top_5_weights = sorted_importance(1:length(top_5_idx));

% Normalize weights
top_5_weights_normalized = top_5_weights / sum(top_5_weights);

fprintf('\nTop 5 competencies for performance prediction:\n');
for i = 1:length(top_5_competencies)
    fprintf('  %d. %s: %.1f%%\n', i, top_5_competencies{i}, top_5_weights_normalized(i) * 100);
end

%% 8. Create Radar Charts for Talent Types
fprintf('\n=== Creating Radar Charts ===\n');

% Color palette for talent types
talent_colors = [
    0.2157 0.4941 0.7216;  % Blue
    0.8941 0.1020 0.1098;  % Red
    0.3020 0.6863 0.2902;  % Green
    0.5961 0.3059 0.6392;  % Purple
    1.0000 0.4980 0.0000;  % Orange
    0.6510 0.3373 0.1569;  % Brown
    0.9 0.7 0.1;           % Yellow
    0.5 0.5 0.5;           % Gray
];

% Calculate overall means for baseline
overall_means = nanmean(matched_comp{:, :}, 1);

% Create subplot radar chart for all talent types
fprintf('✓ Creating talent type radar charts...\n');

figure('Position', [100 100 1400 1000], 'Color', 'white', 'Name', '인재유형별 상위 역량 프로필');

n_types = length(unique_matched_types);
n_cols = ceil(sqrt(n_types));
n_rows = ceil(n_types / n_cols);

for type_idx = 1:n_types
    subplot(n_rows, n_cols, type_idx);

    talent_type = unique_matched_types{type_idx};
    type_data_idx = strcmp(matched_talent_types, talent_type);

    % Calculate mean scores for this talent type
    type_comp_data = matched_comp{type_data_idx, :};
    type_means = nanmean(type_comp_data, 1);

    % Normalize to 0-1 scale
    type_means_norm = type_means / 100;
    overall_means_norm = overall_means / 100;

    % Use top 5 competencies for cleaner visualization
    radar_means = type_means_norm(top_5_idx);
    radar_baseline = overall_means_norm(top_5_idx);
    radar_labels = top_5_competencies;

    % Create radar chart
    draw_talent_radar_chart(type_idx, talent_type, radar_means, radar_baseline, ...
                           radar_labels, talent_colors, true);
end

sgtitle('인재유형별 상위 역량 프로필 (TOP 5 핵심 역량)', 'FontSize', 16, 'FontWeight', 'bold');

%% 9. Save Results
fprintf('\n=== Saving Results ===\n');

% Save analysis results
results = struct();
results.matched_ids = matched_ids;
results.matched_talent_types = matched_talent_types;
results.matched_comp = matched_comp;
results.valid_comp_cols = valid_comp_cols;
results.type_stats = type_stats;
results.top_5_competencies = top_5_competencies;
results.top_5_weights_normalized = top_5_weights_normalized;
results.competency_correlations = competency_correlations;
results.performance_ranking = performance_ranking;

save('final_talent_analysis_results.mat', 'results');

% Create summary table
summary_table = cell2table(type_stats, 'VariableNames', ...
    {'TalentType', 'Count', 'CompetencyMean', 'CompetencyStd', 'TotalScoreMean', 'PerformanceRank'});

try
    writetable(summary_table, 'talent_type_analysis_summary.xlsx', 'Sheet', 'TalentTypeSummary');

    % Top competencies table
    top_comp_table = table();
    top_comp_table.Rank = (1:length(top_5_competencies))';
    top_comp_table.Competency = top_5_competencies';
    top_comp_table.Weight = top_5_weights_normalized;
    top_comp_table.WeightPercent = top_5_weights_normalized * 100;
    top_comp_table.Correlation = competency_correlations(top_5_idx);

    writetable(top_comp_table, 'talent_type_analysis_summary.xlsx', 'Sheet', 'TopCompetencies');

    fprintf('✓ Excel file saved: talent_type_analysis_summary.xlsx\n');
catch
    fprintf('⚠ Excel save failed\n');
end

fprintf('✓ MATLAB file saved: final_talent_analysis_results.mat\n');

%% 10. Final Summary
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 70));
fprintf('               Final Talent Type Analysis Complete\n');
fprintf('%s\n', repmat('=', 1, 70));

fprintf('\n📊 Analysis Summary:\n');
fprintf('  • Analyzed people: %d\n', length(matched_ids));
fprintf('  • Talent types: %d\n', length(unique_matched_types));
fprintf('  • Competency categories: %d\n', length(valid_comp_cols));
fprintf('  • Top predictive competencies: %d\n', length(top_5_competencies));

fprintf('\n🏆 Talent Type Performance Ranking (by expected performance):\n');
% Sort by performance ranking
[~, rank_order] = sort(cell2mat(type_stats(:, 6)), 'descend');
for i = 1:length(rank_order)
    idx = rank_order(i);
    fprintf('  %d. %s (rank: %.0f, count: %d)\n', i, type_stats{idx, 1}, ...
            type_stats{idx, 6}, type_stats{idx, 2});
end

fprintf('\n🎯 Top 5 Predictive Competencies:\n');
for i = 1:length(top_5_competencies)
    fprintf('  %d. %s: %.1f%% (r=%.3f)\n', i, top_5_competencies{i}, ...
            top_5_weights_normalized(i) * 100, competency_correlations(top_5_idx(i)));
end

fprintf('\n✅ Analysis complete with radar charts!\n');
fprintf('%s\n', repmat('=', 1, 70));

%% Helper Function: Draw Talent Radar Chart
function draw_talent_radar_chart(type_idx, talent_type, type_means, baseline_means, labels, colors, is_subplot)
    n_vars = length(type_means);

    % Calculate angles
    angles = linspace(0, 2*pi, n_vars+1);
    type_data = [type_means, type_means(1)];
    baseline_data = [baseline_means, baseline_means(1)];

    hold on;

    % Draw grid circles
    for r = 0.2:0.2:1
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
    end

    % Draw radial grid
    for j = 1:n_vars
        [gx, gy] = pol2cart(angles(j), [0 1]);
        plot(gx, gy, 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
    end

    % Draw baseline (overall mean)
    [base_x, base_y] = pol2cart(angles, baseline_data);
    plot(base_x, base_y, '--', 'Color', [0.5 0.5 0.5], 'LineWidth', 2);

    % Draw talent type data
    [type_x, type_y] = pol2cart(angles, type_data);
    color_idx = mod(type_idx-1, size(colors, 1)) + 1;
    patch(type_x, type_y, colors(color_idx,:), ...
        'FaceAlpha', 0.3, ...
        'EdgeColor', colors(color_idx,:), ...
        'LineWidth', 2.5);
    scatter(type_x(1:end-1), type_y(1:end-1), 50, colors(color_idx,:), 'filled');

    % Add labels
    for j = 1:n_vars
        [lx, ly] = pol2cart(angles(j), 1.2);

        if is_subplot
            txt = sprintf('%s\n%.0f', labels{j}, type_means(j)*100);
            font_size = 8;
        else
            txt = sprintf('%s\n%.0f%%', labels{j}, type_means(j)*100);
            font_size = 10;
        end

        text(lx, ly, txt, ...
            'HorizontalAlignment', 'center', ...
            'FontSize', font_size, 'FontWeight', 'bold');
    end

    % Title
    title(talent_type, 'FontSize', 12, 'FontWeight', 'bold');

    axis equal; axis([-1.4 1.4 -1.4 1.4]); axis off;
    hold off;
end


%% Advanced ML Model for Talent Type Analysis with Full Evaluation
% 교차검증, 하이퍼파라미터 튜닝, Permutation Importance, 성능평가 포함
clear; clc; close all;

%% 1. 데이터 준비
fprintf('========================================\n');
fprintf('🚀 Advanced ML-based Talent Type Analysis\n');
fprintf('========================================\n\n');

% 기존 분석 결과 로드
load('final_talent_analysis_results.mat', 'results');

% 데이터 추출
X = table2array(results.matched_comp);  % 역량 점수 (특성)
y = results.matched_talent_types;        % 인재유형 (레이블)
feature_names = results.valid_comp_cols;

% 레이블 인코딩
[unique_labels, ~, y_encoded] = unique(y);
n_classes = length(unique_labels);
n_features = size(X, 2);
n_samples = size(X, 1);

% 데이터 정규화 (0-1 스케일링)
X_normalized = (X - min(X)) ./ (max(X) - min(X));

fprintf('📊 Data Overview:\n');
fprintf('  • Samples: %d\n', n_samples);
fprintf('  • Features: %d\n', n_features);
fprintf('  • Classes: %d\n', n_classes);
fprintf('  • Class distribution:\n');
for i = 1:n_classes
    fprintf('    - %s: %d samples (%.1f%%)\n', unique_labels{i}, ...
        sum(y_encoded == i), sum(y_encoded == i)/n_samples*100);
end

%% 2. 교차검증 설정
fprintf('\n=== ⚙️ Cross-Validation Setup ===\n');

% Stratified K-fold 교차검증
k_folds = 5;
cv = cvpartition(y_encoded, 'KFold', k_folds, 'Stratify', true);
fprintf('Using %d-fold stratified cross-validation\n', k_folds);

% 성능 메트릭 저장용
performance_metrics = struct();

%% 3. Random Forest with Hyperparameter Tuning
fprintf('\n=== 🌲 Random Forest with Hyperparameter Tuning ===\n');

% 하이퍼파라미터 그리드
rf_params = struct();
rf_params.n_trees = [50, 100, 200];
rf_params.min_leaf = [1, 3, 5];
rf_params.max_splits = [10, 20, 50];

% Grid Search
best_rf_score = 0;
best_rf_params = struct();
rf_cv_scores = [];

fprintf('Performing grid search...\n');
total_combinations = length(rf_params.n_trees) * length(rf_params.min_leaf) * length(rf_params.max_splits);
combo_count = 0;

for n_tree = rf_params.n_trees
    for min_leaf = rf_params.min_leaf
        for max_split = rf_params.max_splits
            combo_count = combo_count + 1;
            
            % 교차검증 점수
            cv_scores = zeros(k_folds, 1);
            
            for fold = 1:k_folds
                % 훈련/검증 데이터 분할
                train_idx = cv.training(fold);
                val_idx = cv.test(fold);
                
                X_train = X_normalized(train_idx, :);
                y_train = y_encoded(train_idx);
                X_val = X_normalized(val_idx, :);
                y_val = y_encoded(val_idx);
                
                % 모델 학습
                rf_temp = TreeBagger(n_tree, X_train, y_train, ...
                    'Method', 'classification', ...
                    'MinLeafSize', min_leaf, ...
                    'MaxNumSplits', max_split);
                
                % 예측 및 평가
                y_pred = cellfun(@str2num, predict(rf_temp, X_val));
                cv_scores(fold) = sum(y_pred == y_val) / length(y_val);
            end
            
            mean_score = mean(cv_scores);
            
            % 최고 성능 파라미터 저장
            if mean_score > best_rf_score
                best_rf_score = mean_score;
                best_rf_params.n_trees = n_tree;
                best_rf_params.min_leaf = min_leaf;
                best_rf_params.max_splits = max_split;
            end
            
            if mod(combo_count, 5) == 0
                fprintf('  Progress: %d/%d combinations tested\n', combo_count, total_combinations);
            end
        end
    end
end

fprintf('\nBest RF Parameters:\n');
fprintf('  • Trees: %d\n', best_rf_params.n_trees);
fprintf('  • Min Leaf Size: %d\n', best_rf_params.min_leaf);
fprintf('  • Max Splits: %d\n', best_rf_params.max_splits);
fprintf('  • CV Accuracy: %.4f\n', best_rf_score);

% 최적 파라미터로 최종 모델 학습
rf_final = TreeBagger(best_rf_params.n_trees, X_normalized, y_encoded, ...
    'Method', 'classification', ...
    'MinLeafSize', best_rf_params.min_leaf, ...
    'MaxNumSplits', best_rf_params.max_splits, ...
    'OOBPrediction', 'on', ...
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', feature_names);

% Feature importance 추출
rf_importance = rf_final.OOBPermutedPredictorDeltaError;
rf_importance_norm = rf_importance / sum(rf_importance);

performance_metrics.rf_cv_accuracy = best_rf_score;
performance_metrics.rf_oob_error = error(rf_final);

%% 4. Gradient Boosting with Hyperparameter Tuning
fprintf('\n=== 🚀 Gradient Boosting with Hyperparameter Tuning ===\n');

% 하이퍼파라미터 그리드
gb_params = struct();
gb_params.n_cycles = [50, 100, 150];
gb_params.learn_rate = [0.05, 0.1, 0.2];
gb_params.max_splits = [5, 10, 20];

% Grid Search
best_gb_score = 0;
best_gb_params = struct();

fprintf('Performing grid search...\n');
total_gb_combinations = length(gb_params.n_cycles) * length(gb_params.learn_rate) * length(gb_params.max_splits);
gb_combo_count = 0;

for n_cycle = gb_params.n_cycles
    for learn_rate = gb_params.learn_rate
        for max_split = gb_params.max_splits
            gb_combo_count = gb_combo_count + 1;
            
            % 교차검증 점수
            cv_scores = zeros(k_folds, 1);
            
            for fold = 1:k_folds
                % 훈련/검증 데이터 분할
                train_idx = cv.training(fold);
                val_idx = cv.test(fold);
                
                X_train = X_normalized(train_idx, :);
                y_train = y_encoded(train_idx);
                X_val = X_normalized(val_idx, :);
                y_val = y_encoded(val_idx);
                
                % 모델 학습
                gb_temp = fitcensemble(X_train, y_train, ...
                    'Method', 'AdaBoostM2', ...
                    'NumLearningCycles', n_cycle, ...
                    'Learners', templateTree('MaxNumSplits', max_split), ...
                    'LearnRate', learn_rate);
                
                % 예측 및 평가
                y_pred = predict(gb_temp, X_val);
                cv_scores(fold) = sum(y_pred == y_val) / length(y_val);
            end
            
            mean_score = mean(cv_scores);
            
            % 최고 성능 파라미터 저장
            if mean_score > best_gb_score
                best_gb_score = mean_score;
                best_gb_params.n_cycles = n_cycle;
                best_gb_params.learn_rate = learn_rate;
                best_gb_params.max_splits = max_split;
            end
            
            if mod(gb_combo_count, 5) == 0
                fprintf('  Progress: %d/%d combinations tested\n', gb_combo_count, total_gb_combinations);
            end
        end
    end
end

fprintf('\nBest GB Parameters:\n');
fprintf('  • Learning Cycles: %d\n', best_gb_params.n_cycles);
fprintf('  • Learning Rate: %.2f\n', best_gb_params.learn_rate);
fprintf('  • Max Splits: %d\n', best_gb_params.max_splits);
fprintf('  • CV Accuracy: %.4f\n', best_gb_score);

% 최적 파라미터로 최종 모델 학습
gb_final = fitcensemble(X_normalized, y_encoded, ...
    'Method', 'AdaBoostM2', ...
    'NumLearningCycles', best_gb_params.n_cycles, ...
    'Learners', templateTree('MaxNumSplits', best_gb_params.max_splits), ...
    'LearnRate', best_gb_params.learn_rate);

% Feature importance 추출
gb_importance = predictorImportance(gb_final);
gb_importance_norm = gb_importance / sum(gb_importance);

performance_metrics.gb_cv_accuracy = best_gb_score;

%% 5. Permutation Importance 계산 (교차검증 기반)
fprintf('\n=== 🔄 Permutation Importance Analysis ===\n');

n_permutations = 20;
perm_importance_cv = zeros(k_folds, n_features);

fprintf('Computing permutation importance with CV...\n');

for fold = 1:k_folds
    fprintf('  Fold %d/%d: ', fold, k_folds);
    
    % 데이터 분할
    train_idx = cv.training(fold);
    val_idx = cv.test(fold);
    
    X_train = X_normalized(train_idx, :);
    y_train = y_encoded(train_idx);
    X_val = X_normalized(val_idx, :);
    y_val = y_encoded(val_idx);
    
    % 이 fold의 모델 학습
    fold_model = TreeBagger(best_rf_params.n_trees, X_train, y_train, ...
        'Method', 'classification', ...
        'MinLeafSize', best_rf_params.min_leaf);
    
    % 기준 성능
    y_pred_base = cellfun(@str2num, predict(fold_model, X_val));
    baseline_acc = sum(y_pred_base == y_val) / length(y_val);
    
    % 각 특성별 permutation importance
    for feat = 1:n_features
        perm_scores = zeros(n_permutations, 1);
        
        for perm = 1:n_permutations
            % 특성 섞기
            X_val_perm = X_val;
            X_val_perm(:, feat) = X_val_perm(randperm(size(X_val, 1)), feat);
            
            % 성능 측정
            y_pred_perm = cellfun(@str2num, predict(fold_model, X_val_perm));
            perm_scores(perm) = sum(y_pred_perm == y_val) / length(y_val);
        end
        
        % Importance = baseline - permuted accuracy
        perm_importance_cv(fold, feat) = baseline_acc - mean(perm_scores);
    end
    
    fprintf('Done\n');
end

% 평균 permutation importance
perm_importance = mean(perm_importance_cv);
perm_importance(perm_importance < 0) = 0;
perm_importance_norm = perm_importance / sum(perm_importance);

% 표준편차 계산 (불확실성 추정)
perm_importance_std = std(perm_importance_cv);

fprintf('\nTop 5 features by Permutation Importance:\n');
[sorted_perm, idx_perm] = sort(perm_importance_norm, 'descend');
for i = 1:min(5, length(sorted_perm))
    fprintf('  %d. %s: %.2f%% (±%.2f%%)\n', i, feature_names{idx_perm(i)}, ...
        sorted_perm(i)*100, perm_importance_std(idx_perm(i))*100);
end

%% 6. 종합 성능 평가
fprintf('\n=== 📈 Comprehensive Performance Evaluation ===\n');

% 홀드아웃 테스트 세트 생성 (최종 평가용)
cv_final = cvpartition(y_encoded, 'HoldOut', 0.2, 'Stratify', true);
X_train_final = X_normalized(cv_final.training, :);
y_train_final = y_encoded(cv_final.training);
X_test_final = X_normalized(cv_final.test, :);
y_test_final = y_encoded(cv_final.test);
y_test_labels = y(cv_final.test);

% Random Forest 성능
rf_test = TreeBagger(best_rf_params.n_trees, X_train_final, y_train_final, ...
    'Method', 'classification', ...
    'MinLeafSize', best_rf_params.min_leaf, ...
    'MaxNumSplits', best_rf_params.max_splits);

y_pred_rf = cellfun(@str2num, predict(rf_test, X_test_final));

% Gradient Boosting 성능
gb_test = fitcensemble(X_train_final, y_train_final, ...
    'Method', 'AdaBoostM2', ...
    'NumLearningCycles', best_gb_params.n_cycles, ...
    'Learners', templateTree('MaxNumSplits', best_gb_params.max_splits), ...
    'LearnRate', best_gb_params.learn_rate);

y_pred_gb = predict(gb_test, X_test_final);

% 성능 메트릭 계산 함수
calculate_metrics = @(y_true, y_pred, n_classes) struct(...
    'accuracy', sum(y_true == y_pred) / length(y_true), ...
    'precision', calculate_precision(y_true, y_pred, n_classes), ...
    'recall', calculate_recall(y_true, y_pred, n_classes), ...
    'f1', calculate_f1(y_true, y_pred, n_classes));

% 각 모델별 성능 계산
rf_metrics = calculate_metrics(y_test_final, y_pred_rf, n_classes);
gb_metrics = calculate_metrics(y_test_final, y_pred_gb, n_classes);

% 앙상블 예측 (투표)
y_pred_ensemble = mode([y_pred_rf, y_pred_gb], 2);
ensemble_metrics = calculate_metrics(y_test_final, y_pred_ensemble, n_classes);

% 결과 출력
fprintf('\nModel Performance on Test Set:\n');
fprintf('----------------------------------------\n');
fprintf('Random Forest:\n');
fprintf('  • Accuracy:  %.4f\n', rf_metrics.accuracy);
fprintf('  • Precision: %.4f (macro-avg)\n', mean(rf_metrics.precision));
fprintf('  • Recall:    %.4f (macro-avg)\n', mean(rf_metrics.recall));
fprintf('  • F1-Score:  %.4f (macro-avg)\n', mean(rf_metrics.f1));

fprintf('\nGradient Boosting:\n');
fprintf('  • Accuracy:  %.4f\n', gb_metrics.accuracy);
fprintf('  • Precision: %.4f (macro-avg)\n', mean(gb_metrics.precision));
fprintf('  • Recall:    %.4f (macro-avg)\n', mean(gb_metrics.recall));
fprintf('  • F1-Score:  %.4f (macro-avg)\n', mean(gb_metrics.f1));

fprintf('\nEnsemble (Voting):\n');
fprintf('  • Accuracy:  %.4f\n', ensemble_metrics.accuracy);
fprintf('  • Precision: %.4f (macro-avg)\n', mean(ensemble_metrics.precision));
fprintf('  • Recall:    %.4f (macro-avg)\n', mean(ensemble_metrics.recall));
fprintf('  • F1-Score:  %.4f (macro-avg)\n', mean(ensemble_metrics.f1));

% Confusion Matrix
figure('Position', [100 100 1400 400], 'Color', 'white');

subplot(1, 3, 1);
conf_rf = confusionmat(y_test_final, y_pred_rf);
heatmap(unique_labels, unique_labels, conf_rf);
title('Random Forest Confusion Matrix');
xlabel('Predicted'); ylabel('Actual');

subplot(1, 3, 2);
conf_gb = confusionmat(y_test_final, y_pred_gb);
heatmap(unique_labels, unique_labels, conf_gb);
title('Gradient Boosting Confusion Matrix');
xlabel('Predicted'); ylabel('Actual');

subplot(1, 3, 3);
conf_ensemble = confusionmat(y_test_final, y_pred_ensemble);
heatmap(unique_labels, unique_labels, conf_ensemble);
title('Ensemble Confusion Matrix');
xlabel('Predicted'); ylabel('Actual');

sgtitle('Model Performance Comparison', 'FontSize', 14, 'FontWeight', 'bold');

%% 7. 앙상블 Feature Importance
fprintf('\n=== 🎯 Final Ensemble Feature Importance ===\n');

% 모든 방법의 가중 평균 (성능 기반 가중치)
weight_rf = rf_metrics.accuracy;
weight_gb = gb_metrics.accuracy;
weight_perm = 1.0;  % Permutation importance는 모델 독립적이므로 고정 가중치

total_weight = weight_rf + weight_gb + weight_perm;

ensemble_importance = (rf_importance_norm * weight_rf + ...
                      gb_importance_norm * weight_gb + ...
                      perm_importance_norm * weight_perm) / total_weight;

% 정규화
ensemble_importance_norm = ensemble_importance / sum(ensemble_importance);

fprintf('\nTop 10 Most Important Features (Weighted Ensemble):\n');
fprintf('Feature                          | Importance | RF     | GB     | Perm   \n');
fprintf('--------------------------------|------------|--------|--------|--------\n');

[sorted_ensemble, idx_ensemble] = sort(ensemble_importance_norm, 'descend');
for i = 1:min(10, length(sorted_ensemble))
    feat_idx = idx_ensemble(i);
    fprintf('%-31s | %9.2f%% | %5.2f%% | %5.2f%% | %5.2f%%\n', ...
        feature_names{feat_idx}, ...
        sorted_ensemble(i)*100, ...
        rf_importance_norm(feat_idx)*100, ...
        gb_importance_norm(feat_idx)*100, ...
        perm_importance_norm(feat_idx)*100);
end

%% 8. Feature Importance 시각화
fprintf('\n=== 📊 Creating Visualizations ===\n');

% Figure 1: Feature Importance 비교
figure('Position', [100 100 1400 800], 'Color', 'white');

% Top 15 features
top_n = min(15, n_features);
[~, top_idx] = sort(ensemble_importance_norm, 'descend');
top_features = feature_names(top_idx(1:top_n));

% Subplot 1: Bar chart comparison
subplot(2, 2, 1:2);
x = 1:top_n;
width = 0.25;

bar(x - width, rf_importance_norm(top_idx(1:top_n)), width, 'FaceColor', [0.2 0.5 0.8]);
hold on;
bar(x, gb_importance_norm(top_idx(1:top_n)), width, 'FaceColor', [0.8 0.3 0.3]);
bar(x + width, perm_importance_norm(top_idx(1:top_n)), width, 'FaceColor', [0.3 0.7 0.3]);

set(gca, 'XTick', x, 'XTickLabel', top_features, 'XTickLabelRotation', 45);
ylabel('Importance');
title('Feature Importance Comparison (Top 15)', 'FontWeight', 'bold');
legend('Random Forest', 'Gradient Boosting', 'Permutation', 'Location', 'northeast');
grid on;

% Subplot 2: Ensemble importance
subplot(2, 2, 3);
barh(ensemble_importance_norm(top_idx(top_n:-1:1)), 'FaceColor', [0.5 0.2 0.7]);
set(gca, 'YTick', 1:top_n, 'YTickLabel', top_features(top_n:-1:1));
xlabel('Importance');
title('Final Ensemble Importance', 'FontWeight', 'bold');
grid on;

% Subplot 3: Cumulative importance
subplot(2, 2, 4);
cumsum_importance = cumsum(sorted_ensemble);
plot(1:length(cumsum_importance), cumsum_importance, 'LineWidth', 2, 'Color', [0.2 0.5 0.8]);
hold on;
plot([1 length(cumsum_importance)], [0.8 0.8], 'r--', 'LineWidth', 1.5);
xlabel('Number of Features');
ylabel('Cumulative Importance');
title('Cumulative Feature Importance', 'FontWeight', 'bold');
legend('Cumulative', '80% threshold', 'Location', 'southeast');
grid on;

n_features_80 = find(cumsum_importance >= 0.8, 1);
text(n_features_80, 0.8, sprintf('  %d features\n  explain 80%%', n_features_80), ...
    'FontSize', 10, 'VerticalAlignment', 'bottom');

sgtitle('Feature Importance Analysis', 'FontSize', 16, 'FontWeight', 'bold');

% Figure 2: Permutation Importance with Error Bars
figure('Position', [100 100 1200 600], 'Color', 'white');

errorbar(1:top_n, perm_importance_norm(top_idx(1:top_n)), ...
    perm_importance_std(top_idx(1:top_n)), 'o-', 'LineWidth', 2, ...
    'MarkerSize', 8, 'MarkerFaceColor', [0.3 0.7 0.3]);

set(gca, 'XTick', 1:top_n, 'XTickLabel', top_features, 'XTickLabelRotation', 45);
ylabel('Permutation Importance');
title('Permutation Importance with Uncertainty (±1 std)', 'FontWeight', 'bold', 'FontSize', 14);
grid on;

%% 9. 클래스별 예측 성능 분석
fprintf('\n=== 📊 Class-wise Performance Analysis ===\n');

fprintf('\nPer-class Performance (Ensemble Model):\n');
fprintf('Class                 | Precision | Recall | F1-Score | Support\n');
fprintf('---------------------|-----------|--------|----------|--------\n');

for class = 1:n_classes
    class_precision = ensemble_metrics.precision(class);
    class_recall = ensemble_metrics.recall(class);
    class_f1 = ensemble_metrics.f1(class);
    class_support = sum(y_test_final == class);
    
    fprintf('%-20s | %9.4f | %6.4f | %8.4f | %7d\n', ...
        unique_labels{class}, class_precision, class_recall, class_f1, class_support);
end

%% 10. 결과 저장
fprintf('\n=== 💾 Saving Results ===\n');

% 종합 결과 구조체
final_results = struct();
final_results.feature_names = feature_names;
final_results.ensemble_importance = ensemble_importance_norm;
final_results.rf_importance = rf_importance_norm;
final_results.gb_importance = gb_importance_norm;
final_results.perm_importance = perm_importance_norm;
final_results.perm_importance_std = perm_importance_std;
final_results.best_rf_params = best_rf_params;
final_results.best_gb_params = best_gb_params;
final_results.performance_metrics = performance_metrics;
final_results.rf_metrics = rf_metrics;
final_results.gb_metrics = gb_metrics;
final_results.ensemble_metrics = ensemble_metrics;
final_results.confusion_matrices = struct('rf', conf_rf, 'gb', conf_gb, 'ensemble', conf_ensemble);

save('advanced_ml_talent_results.mat', 'final_results');

% Excel 파일로 내보내기
% Sheet 1: Feature Importance
importance_table = table();
importance_table.Feature = feature_names';
importance_table.Ensemble_Weight = ensemble_importance_norm * 100;
importance_table.RF_Weight = rf_importance_norm * 100;
importance_table.GB_Weight = gb_importance_norm * 100;
importance_table.Perm_Weight = perm_importance_norm * 100;
importance_table.Perm_Std = perm_importance_std * 100;

% Sheet 2: Model Performance
perf_table = table();
perf_table.Model = {'Random Forest'; 'Gradient Boosting'; 'Ensemble'};
perf_table.CV_Accuracy = [best_rf_score; best_gb_score; NaN];
perf_table.Test_Accuracy = [rf_metrics.accuracy; gb_metrics.accuracy; ensemble_metrics.accuracy];
perf_table.Precision = [mean(rf_metrics.precision); mean(gb_metrics.precision); mean(ensemble_metrics.precision)];
perf_table.Recall = [mean(rf_metrics.recall); mean(gb_metrics.recall); mean(ensemble_metrics.recall)];
perf_table.F1_Score = [mean(rf_metrics.f1); mean(gb_metrics.f1); mean(ensemble_metrics.f1)];

% Sheet 3: Per-class Performance
class_perf_table = table();
class_perf_table.TalentType = unique_labels;
class_perf_table.Precision = ensemble_metrics.precision;
class_perf_table.Recall = ensemble_metrics.recall;
class_perf_table.F1_Score = ensemble_metrics.f1;

for class = 1:n_classes
    class_perf_table.Support(class) = sum(y_test_final == class);
end

% Excel 파일 쓰기
writetable(importance_table, 'advanced_ml_analysis.xlsx', 'Sheet', 'FeatureImportance');
writetable(perf_table, 'advanced_ml_analysis.xlsx', 'Sheet', 'ModelPerformance');
writetable(class_perf_table, 'advanced_ml_analysis.xlsx', 'Sheet', 'ClassPerformance');

fprintf('✓ Results saved to:\n');
fprintf('  • advanced_ml_talent_results.mat\n');
fprintf('  • advanced_ml_analysis.xlsx\n');

%% 11. 최종 요약 보고서
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 80));
fprintf('                    FINAL ANALYSIS SUMMARY REPORT\n');
fprintf('%s\n', repmat('=', 1, 80));

fprintf('\n📊 DATA SUMMARY:\n');
fprintf('  • Total samples: %d\n', n_samples);
fprintf('  • Features: %d\n', n_features);
fprintf('  • Talent types: %d\n', n_classes);
fprintf('  • Train/Test split: 80%%/20%%\n');
fprintf('  • Cross-validation: %d-fold stratified\n', k_folds);

fprintf('\n🏆 BEST MODEL PERFORMANCE:\n');
[best_acc, best_model_idx] = max([rf_metrics.accuracy, gb_metrics.accuracy, ensemble_metrics.accuracy]);
model_names = {'Random Forest', 'Gradient Boosting', 'Ensemble'};
fprintf('  • Best model: %s\n', model_names{best_model_idx});
fprintf('  • Test accuracy: %.2f%%\n', best_acc * 100);
fprintf('  • CV accuracy: %.2f%%\n', max(best_rf_score, best_gb_score) * 100);

fprintf('\n🎯 TOP 5 CRITICAL FEATURES:\n');
for i = 1:min(5, length(sorted_ensemble))
    fprintf('  %d. %-30s: %5.2f%% (±%.2f%%)\n', ...
        i, feature_names{idx_ensemble(i)}, ...
        sorted_ensemble(i)*100, ...
        perm_importance_std(idx_ensemble(i))*100);
end

fprintf('\n💡 KEY INSIGHTS:\n');
fprintf('  • Features needed for 80%% importance: %d\n', n_features_80);
fprintf('  • Most consistent feature across methods: %s\n', feature_names{idx_ensemble(1)});
fprintf('  • Highest uncertainty feature: %s (±%.2f%%)\n', ...
    feature_names{find(perm_importance_std == max(perm_importance_std), 1)}, ...
    max(perm_importance_std)*100);

% 클래스 불균형 경고
class_counts = histcounts(y_encoded, 1:(n_classes+1));
imbalance_ratio = max(class_counts) / min(class_counts);
if imbalance_ratio > 3
    fprintf('\n⚠️  WARNING: Class imbalance detected (ratio %.1f:1)\n', imbalance_ratio);
    fprintf('    Consider using SMOTE or class weights for better performance\n');
end

fprintf('\n✅ Analysis completed successfully!\n');
fprintf('%s\n', repmat('=', 1, 80));

%% Helper Functions
function precision = calculate_precision(y_true, y_pred, n_classes)
    precision = zeros(n_classes, 1);
    for i = 1:n_classes
        tp = sum(y_true == i & y_pred == i);
        fp = sum(y_true ~= i & y_pred == i);
        if (tp + fp) > 0
            precision(i) = tp / (tp + fp);
        end
    end
end

function recall = calculate_recall(y_true, y_pred, n_classes)
    recall = zeros(n_classes, 1);
    for i = 1:n_classes
        tp = sum(y_true == i & y_pred == i);
        fn = sum(y_true == i & y_pred ~= i);
        if (tp + fn) > 0
            recall(i) = tp / (tp + fn);
        end
    end
end

function f1 = calculate_f1(y_true, y_pred, n_classes)
    precision = calculate_precision(y_true, y_pred, n_classes);
    recall = calculate_recall(y_true, y_pred, n_classes);
    f1 = zeros(n_classes, 1);
    for i = 1:n_classes
        if (precision(i) + recall(i)) > 0
            f1(i) = 2 * (precision(i) * recall(i)) / (precision(i) + recall(i));
        end
    end
end