%% Simplified Performance Prediction Weight Analysis
% Focus on correlation analysis with available data

clear; clc;

%% 1. Load and Prepare Data
fprintf('=== Loading and Preparing Data ===\n');
load('talent_competency_merged_data.mat');

matched_data = analysis_data.matched_data;
matched_performance = analysis_data.matched_performance;
matched_talent_types = analysis_data.matched_talent_types;
competency_headers = analysis_data.competency_headers;

fprintf('Original data: %d people, %d competencies\n', size(matched_data, 1), size(matched_data, 2));

% Remove people with no performance score
valid_people_idx = matched_performance > 0;
X = matched_data(valid_people_idx, :);
y = matched_performance(valid_people_idx);
talent_types = matched_talent_types(valid_people_idx);

fprintf('Valid people with performance scores: %d\n', length(y));

% Remove competencies with all missing values
non_empty_comp_idx = ~all(isnan(X), 1);
X_filtered = X(:, non_empty_comp_idx);
filtered_competencies = competency_headers(non_empty_comp_idx);

fprintf('Competencies with some data: %d\n', sum(non_empty_comp_idx));

%% 2. Calculate Correlations and Differences
fprintf('\n=== Calculating Correlations and Group Differences ===\n');

% High vs Low performance groups
high_perf_threshold = median(y);
high_perf_idx = y > high_perf_threshold;
low_perf_idx = y <= high_perf_threshold;

fprintf('High performance group: %d people (score > %.1f)\n', sum(high_perf_idx), high_perf_threshold);
fprintf('Low performance group: %d people (score <= %.1f)\n', sum(low_perf_idx), high_perf_threshold);

% Analyze each competency
results = [];
for i = 1:size(X_filtered, 2)
    comp_scores = X_filtered(:, i);
    valid_data_idx = ~isnan(comp_scores);

    if sum(valid_data_idx) >= 5  % At least 5 valid data points
        scores_valid = comp_scores(valid_data_idx);
        perf_valid = y(valid_data_idx);

        % Correlation with performance
        if length(scores_valid) > 2 && var(scores_valid) > 0 && var(perf_valid) > 0
            correlation = sum((scores_valid - mean(scores_valid)) .* (perf_valid - mean(perf_valid))) / ...
                         (sqrt(sum((scores_valid - mean(scores_valid)).^2)) * sqrt(sum((perf_valid - mean(perf_valid)).^2)));
        else
            correlation = 0;
        end

        % Group difference
        high_group_scores = comp_scores(valid_data_idx & high_perf_idx);
        low_group_scores = comp_scores(valid_data_idx & low_perf_idx);

        if ~isempty(high_group_scores) && ~isempty(low_group_scores)
            mean_diff = nanmean(high_group_scores) - nanmean(low_group_scores);
            effect_size = abs(mean_diff) / (nanstd(comp_scores) + eps);
        else
            mean_diff = 0;
            effect_size = 0;
        end

        % Data quality score
        data_completeness = sum(valid_data_idx) / length(comp_scores);

        % Store results
        results = [results; i, correlation, mean_diff, effect_size, data_completeness, sum(valid_data_idx)];
    end
end

if isempty(results)
    fprintf('Warning: No competencies met the minimum data requirement\n');
    return;
end

fprintf('Competencies analyzed: %d\n', size(results, 1));

%% 3. Calculate Composite Weights
fprintf('\n=== Calculating Composite Weights ===\n');

% Extract metrics
comp_indices = results(:, 1);
correlations = results(:, 2);
mean_differences = results(:, 3);
effect_sizes = results(:, 4);
data_completeness = results(:, 5);
sample_sizes = results(:, 6);

% Normalize metrics (0-1 scale)
norm_correlation = abs(correlations) / (max(abs(correlations)) + eps);
norm_effect_size = effect_sizes / (max(effect_sizes) + eps);
norm_completeness = data_completeness;

% Combined weight (weighted average of normalized metrics)
weight_correlation = 0.4;  % 40% weight for correlation
weight_effect = 0.3;       % 30% weight for effect size
weight_completeness = 0.3; % 30% weight for data completeness

composite_weights = weight_correlation * norm_correlation + ...
                   weight_effect * norm_effect_size + ...
                   weight_completeness * norm_completeness;

%% 4. Rank and Display Results
fprintf('\n=== Top Competencies for Performance Prediction ===\n');

% Sort by composite weight
[sorted_weights, sort_idx] = sort(composite_weights, 'descend');

fprintf('Rank | Competency                    | Weight | Corr   | Diff   | Effect | Complete | N\n');
fprintf('-----|-------------------------------|--------|--------|--------|--------|----------|---\n');

top_competencies = {};
top_weights = [];

for i = 1:min(15, length(sort_idx))
    result_idx = sort_idx(i);
    comp_idx = comp_indices(result_idx);
    comp_name = filtered_competencies{comp_idx};

    fprintf('%4d | %-29s | %6.3f | %6.3f | %6.2f | %6.3f | %8.1f%% | %2.0f\n', ...
            i, comp_name, ...
            composite_weights(result_idx), ...
            correlations(result_idx), ...
            mean_differences(result_idx), ...
            effect_sizes(result_idx), ...
            data_completeness(result_idx) * 100, ...
            sample_sizes(result_idx));

    top_competencies{end+1} = comp_name;
    top_weights = [top_weights; composite_weights(result_idx)];
end

%% 5. Talent Type Analysis
fprintf('\n=== Talent Type Analysis ===\n');

unique_types = unique(talent_types);
fprintf('Performance by talent type:\n');

for i = 1:length(unique_types)
    type_idx = strcmp(talent_types, unique_types{i});
    type_performance = y(type_idx);

    fprintf('  %s: %.2f ± %.2f (n=%d)\n', ...
            unique_types{i}, ...
            mean(type_performance), ...
            std(type_performance), ...
            sum(type_idx));
end

%% 6. Create Recommendations
fprintf('\n=== Recommendations for Performance Prediction ===\n');

% Top 5 competencies for prediction model
top_5_indices = sort_idx(1:min(5, length(sort_idx)));
top_5_names = {};
top_5_final_weights = [];

total_weight = sum(composite_weights(top_5_indices));
for i = 1:length(top_5_indices)
    result_idx = top_5_indices(i);
    comp_idx = comp_indices(result_idx);
    comp_name = filtered_competencies{comp_idx};

    normalized_weight = composite_weights(result_idx) / total_weight;

    top_5_names{end+1} = comp_name;
    top_5_final_weights = [top_5_final_weights; normalized_weight];
end

fprintf('Top 5 competencies for performance prediction model:\n');
for i = 1:length(top_5_names)
    fprintf('  %d. %s: %.1f%% weight\n', i, top_5_names{i}, top_5_final_weights(i) * 100);
end

%% 7. Save Results
fprintf('\n=== Saving Results ===\n');

% Create results table
results_table = table();
for i = 1:length(sort_idx)
    result_idx = sort_idx(i);
    comp_idx = comp_indices(result_idx);

    results_table.Rank(i) = i;
    results_table.Competency(i) = filtered_competencies(comp_idx);
    results_table.Composite_Weight(i) = composite_weights(result_idx);
    results_table.Correlation(i) = correlations(result_idx);
    results_table.Mean_Difference(i) = mean_differences(result_idx);
    results_table.Effect_Size(i) = effect_sizes(result_idx);
    results_table.Data_Completeness(i) = data_completeness(result_idx);
    results_table.Sample_Size(i) = sample_sizes(result_idx);
end

% Save final recommendations
recommendations = struct();
recommendations.top_5_competencies = top_5_names;
recommendations.top_5_weights = top_5_final_weights;
recommendations.all_results = results_table;
recommendations.talent_type_performance = containers.Map(unique_types, ...
    cellfun(@(x) mean(y(strcmp(talent_types, x))), unique_types));

save('performance_prediction_recommendations.mat', 'recommendations');
writetable(results_table, 'competency_analysis_results.xlsx');

fprintf('Results saved:\n');
fprintf('  - MATLAB file: performance_prediction_recommendations.mat\n');
fprintf('  - Excel file: competency_analysis_results.xlsx\n');

fprintf('\n=== Analysis Complete ===\n');
fprintf('성과 예측을 위한 역량 가중치 분석이 완료되었습니다.\n');
fprintf('상위 5개 역량항목을 활용하여 성과 예측 모델을 구축할 수 있습니다.\n');