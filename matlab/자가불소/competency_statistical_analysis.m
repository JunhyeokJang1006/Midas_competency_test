%% Competency Statistical Analysis by Talent Type
% Stage 2: Analyze competency characteristics by talent type

clear; clc;

%% 1. Load Merged Data
fprintf('=== Loading Merged Data ===\n');
load('talent_competency_merged_data.mat');

fprintf('Loaded data for %d people\n', length(analysis_data.matched_ids));
fprintf('Competency items: %d\n', length(analysis_data.competency_headers));

%% 2. Analyze Data Quality
fprintf('\n=== Data Quality Analysis ===\n');

matched_data = analysis_data.matched_data;
matched_performance = analysis_data.matched_performance;
matched_talent_types = analysis_data.matched_talent_types;
competency_headers = analysis_data.competency_headers;

% Check for missing values
missing_ratio = sum(isnan(matched_data), 1) ./ size(matched_data, 1);
fprintf('Missing value ratios by competency:\n');
for i = 1:min(10, length(competency_headers))
    fprintf('  %s: %.2f%%\n', competency_headers{i}, missing_ratio(i)*100);
end

% Remove competencies with too many missing values (>50%)
valid_comp_idx = missing_ratio < 0.5;
valid_competencies = competency_headers(valid_comp_idx);
valid_data = matched_data(:, valid_comp_idx);

fprintf('\nValid competencies (missing < 50%%): %d\n', sum(valid_comp_idx));

%% 3. Talent Type Performance Analysis
fprintf('\n=== Talent Type Performance Analysis ===\n');

unique_talent_types = unique(matched_talent_types);
performance_by_type = [];
competency_means_by_type = [];
competency_stds_by_type = [];

fprintf('Performance by talent type:\n');
for i = 1:length(unique_talent_types)
    talent_type = unique_talent_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);

    % Performance statistics
    type_performance = matched_performance(type_idx);
    mean_performance = mean(type_performance);
    count = sum(type_idx);

    fprintf('  %s: %.2f (n=%d)\n', talent_type, mean_performance, count);
    performance_by_type = [performance_by_type; mean_performance];

    % Competency statistics for this talent type
    type_data = valid_data(type_idx, :);
    type_means = nanmean(type_data, 1);
    type_stds = nanstd(type_data, 1);

    competency_means_by_type = [competency_means_by_type; type_means];
    competency_stds_by_type = [competency_stds_by_type; type_stds];
end

%% 4. Competency Analysis by Performance Level
fprintf('\n=== Competency Analysis by Performance Level ===\n');

% Divide into high and low performance groups
high_perf_threshold = median(matched_performance);
high_perf_idx = matched_performance > high_perf_threshold;
low_perf_idx = matched_performance <= high_perf_threshold;

fprintf('High performance group: %d people (performance > %.1f)\n', ...
        sum(high_perf_idx), high_perf_threshold);
fprintf('Low performance group: %d people (performance <= %.1f)\n', ...
        sum(low_perf_idx), high_perf_threshold);

% Compare competency scores between high and low performance groups
high_perf_data = valid_data(high_perf_idx, :);
low_perf_data = valid_data(low_perf_idx, :);

high_perf_means = nanmean(high_perf_data, 1);
low_perf_means = nanmean(low_perf_data, 1);

competency_differences = high_perf_means - low_perf_means;

%% 5. Statistical Tests for Competency Differences
fprintf('\n=== Statistical Tests ===\n');

p_values = [];
t_stats = [];

for i = 1:length(valid_competencies)
    high_scores = high_perf_data(:, i);
    low_scores = low_perf_data(:, i);

    % Remove NaN values
    high_scores = high_scores(~isnan(high_scores));
    low_scores = low_scores(~isnan(low_scores));

    if length(high_scores) >= 3 && length(low_scores) >= 3
        [h, p, ci, stats] = ttest2(high_scores, low_scores);
        p_values = [p_values, p];
        t_stats = [t_stats, stats.tstat];
    else
        p_values = [p_values, NaN];
        t_stats = [t_stats, NaN];
    end
end

% Significant differences (p < 0.05)
significant_idx = p_values < 0.05 & ~isnan(p_values);
fprintf('Significant competency differences (p < 0.05): %d\n', sum(significant_idx));

%% 6. Identify Top Discriminating Competencies
fprintf('\n=== Top Discriminating Competencies ===\n');

% Combine effect size and significance
effect_sizes = abs(competency_differences) ./ (nanstd(valid_data, 1) + eps);
discrimination_score = effect_sizes .* (1 - p_values);  % Higher is better

% Sort by discrimination score
[sorted_scores, sort_idx] = sort(discrimination_score, 'descend');

fprintf('Top 10 discriminating competencies:\n');
for i = 1:min(10, length(sort_idx))
    idx = sort_idx(i);
    if ~isnan(sorted_scores(i))
        fprintf('  %d. %s\n', i, valid_competencies{idx});
        fprintf('      Difference: %.2f, Effect size: %.3f, p-value: %.3f\n', ...
                competency_differences(idx), effect_sizes(idx), p_values(idx));
    end
end

%% 7. Correlation Analysis
fprintf('\n=== Correlation Analysis ===\n');

% Calculate correlations between competencies and performance
correlations = [];
for i = 1:size(valid_data, 2)
    comp_scores = valid_data(:, i);
    valid_pairs_idx = ~isnan(comp_scores) & ~isnan(matched_performance);

    if sum(valid_pairs_idx) >= 10  % At least 10 valid pairs
        corr_coef = corr(comp_scores(valid_pairs_idx), ...
                        matched_performance(valid_pairs_idx));
        correlations = [correlations, corr_coef];
    else
        correlations = [correlations, NaN];
    end
end

% Sort by absolute correlation
[sorted_corr, corr_sort_idx] = sort(abs(correlations), 'descend');

fprintf('Top 10 competencies by correlation with performance:\n');
for i = 1:min(10, length(corr_sort_idx))
    idx = corr_sort_idx(i);
    if ~isnan(sorted_corr(i))
        fprintf('  %d. %s: r = %.3f\n', i, valid_competencies{idx}, correlations(idx));
    end
end

%% 8. Save Analysis Results
fprintf('\n=== Saving Analysis Results ===\n');

% Create comprehensive results structure
statistical_results = struct();
statistical_results.unique_talent_types = unique_talent_types;
statistical_results.performance_by_type = performance_by_type;
statistical_results.competency_means_by_type = competency_means_by_type;
statistical_results.competency_stds_by_type = competency_stds_by_type;
statistical_results.valid_competencies = valid_competencies;
statistical_results.high_perf_means = high_perf_means;
statistical_results.low_perf_means = low_perf_means;
statistical_results.competency_differences = competency_differences;
statistical_results.p_values = p_values;
statistical_results.t_stats = t_stats;
statistical_results.effect_sizes = effect_sizes;
statistical_results.discrimination_score = discrimination_score;
statistical_results.correlations = correlations;
statistical_results.top_discriminating_idx = sort_idx(1:min(10, length(sort_idx)));
statistical_results.top_correlation_idx = corr_sort_idx(1:min(10, length(corr_sort_idx)));

% Save results
save('competency_statistical_results.mat', 'statistical_results');
fprintf('Statistical analysis results saved: competency_statistical_results.mat\n');

fprintf('\n=== Stage 2 Complete ===\n');
fprintf('Next: Calculate weights for performance prediction\n');