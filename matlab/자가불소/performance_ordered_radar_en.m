%% Performance-Ordered Talent Type Radar Chart Analysis (English Version)
% Creating radar charts ordered by performance ranking and analyzing data usage

clear; clc; close all;

%% Global Settings
set(0, 'DefaultAxesFontName', 'Arial');
set(0, 'DefaultTextFontName', 'Arial');
set(0, 'DefaultAxesFontSize', 10);
set(0, 'DefaultTextFontSize', 10);

%% 1. Data Loading and Analysis
fprintf('========================================\n');
fprintf('Performance-Ordered Talent Analysis\n');
fprintf('========================================\n\n');

% File paths
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

% Load data
fprintf('Loading data...\n');
hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
comp_upper = readtable(comp_file, 'Sheet', 3, 'VariableNamingRule', 'preserve');
comp_total = readtable(comp_file, 'Sheet', 4, 'VariableNamingRule', 'preserve');

fprintf('  HR data: %d people\n', height(hr_data));
fprintf('  Upper competency data: %d people\n', height(comp_upper));
fprintf('  Total score data: %d people\n', height(comp_total));

%% 2. Data Usage Analysis
fprintf('\n=== Data Usage Analysis ===\n');

% Find talent type column
talent_col_names = hr_data.Properties.VariableNames;
talent_col_found = '';
for i = 1:length(talent_col_names)
    if contains(talent_col_names{i}, 'talent', 'IgnoreCase', true) || ...
       contains(talent_col_names{i}, '인재유형')
        talent_col_found = talent_col_names{i};
        break;
    end
end

if isempty(talent_col_found)
    fprintf('Warning: Talent type column not found\n');
    return;
end

fprintf('Found talent type column: %s\n', talent_col_found);

% Filter data with talent types
talent_data = hr_data.(talent_col_found);
valid_talent_idx = ~cellfun(@isempty, talent_data) & ~ismissing(talent_data);
hr_clean = hr_data(valid_talent_idx, :);

fprintf('Total HR data: %d people\n', height(hr_data));
fprintf('People with talent types: %d people (%.1f%%)\n', height(hr_clean), height(hr_clean)/height(hr_data)*100);
fprintf('Missing talent types: %d people (%.1f%%)\n', sum(~valid_talent_idx), sum(~valid_talent_idx)/height(hr_data)*100);

% Show talent type distribution
unique_talent_types = unique(talent_data(valid_talent_idx));
fprintf('\nTalent type distribution:\n');
for i = 1:length(unique_talent_types)
    count = sum(strcmp(talent_data, unique_talent_types{i}));
    fprintf('  %s: %d people\n', unique_talent_types{i}, count);
end

%% 3. ID Matching Analysis
fprintf('\n=== ID Matching Analysis ===\n');

hr_clean_ids = hr_clean.ID;
comp_upper_ids = comp_upper.ID;
comp_total_ids = comp_total.ID;

[matched_ids, hr_idx, comp_idx] = intersect(hr_clean_ids, comp_upper_ids);
[total_matched_ids, hr_total_idx, comp_total_idx] = intersect(hr_clean_ids, comp_total_ids);

fprintf('HR with talent types: %d people\n', height(hr_clean));
fprintf('HR matched with upper competency: %d people (%.1f%%)\n', ...
        length(matched_ids), length(matched_ids)/height(hr_clean)*100);
fprintf('HR matched with total scores: %d people (%.1f%%)\n', ...
        length(total_matched_ids), length(total_matched_ids)/height(hr_clean)*100);

% Final analysis dataset
matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, :);
matched_total = comp_total(comp_total_idx, :);

matched_talent_types = matched_hr.(talent_col_found);
total_scores = matched_total{:, end}; % Last column is total score

fprintf('\nFinal analysis dataset: %d people\n', length(matched_ids));
fprintf('Data usage efficiency: %.1f%% of total HR data\n', length(matched_ids)/height(hr_data)*100);

%% 4. Extract Competency Categories
fprintf('\n=== Competency Categories ===\n');

% Find valid competency columns (from 6th column onwards)
valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(matched_comp)
    col_name = matched_comp.Properties.VariableNames{i};
    col_data = matched_comp{:, i};

    % Select numeric columns with less than 50% missing values
    if isnumeric(col_data) && sum(~isnan(col_data)) >= length(col_data) * 0.5
        valid_comp_cols{end+1} = col_name;
        valid_comp_indices = [valid_comp_indices, i];

        data_completeness = sum(~isnan(col_data)) / length(col_data) * 100;
        fprintf('  %d. %s (%.1f%% complete)\n', length(valid_comp_cols), col_name, data_completeness);
    end
end

fprintf('Valid competency categories: %d\n', length(valid_comp_cols));

%% 5. Performance Ranking Definition
performance_ranking = containers.Map();
performance_ranking('자연성') = 8;
performance_ranking('성실한 가연성') = 7;
performance_ranking('유익한 불연성') = 6;
performance_ranking('유능한 불연성') = 5;
performance_ranking('게으른 가연성') = 4;
performance_ranking('무능한 불연성') = 3;
performance_ranking('위장형 소화성') = 2;
performance_ranking('소화성') = 1;

%% 6. Talent Type Analysis
fprintf('\n=== Talent Type Performance Analysis ===\n');

unique_types = unique(matched_talent_types);
type_stats = [];

fprintf('Talent Type Analysis:\n');
fprintf('Talent Type           | Count | Comp Avg | Total Avg | Perf Rank | Data Quality\n');
fprintf('----------------------|-------|----------|-----------|-----------|-------------\n');

for i = 1:length(unique_types)
    talent_type = unique_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);

    % Basic statistics
    count = sum(type_idx);

    % Competency statistics
    type_comp_data = matched_comp{type_idx, valid_comp_indices};
    comp_mean = nanmean(type_comp_data(:));

    % Total score statistics
    type_total_scores = total_scores(type_idx);
    total_mean = nanmean(type_total_scores);

    % Performance rank
    if performance_ranking.isKey(talent_type)
        perf_rank = performance_ranking(talent_type);
    else
        perf_rank = 0;
    end

    % Data quality (missing data ratio)
    data_quality = sum(~isnan(type_comp_data(:))) / numel(type_comp_data) * 100;

    fprintf('%-21s | %5d | %8.1f | %9.1f | %9d | %10.1f%%\n', ...
            talent_type, count, comp_mean, total_mean, perf_rank, data_quality);

    type_stats = [type_stats; {talent_type, count, comp_mean, total_mean, perf_rank, data_quality}];
end

%% 7. Competency Weight Analysis
fprintf('\n=== Competency Weight Analysis ===\n');

% Map performance scores
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    if performance_ranking.isKey(matched_talent_types{i})
        performance_scores(i) = performance_ranking(matched_talent_types{i});
    end
end

% Analyze each competency
comp_importance = [];
comp_correlations = [];

fprintf('Competency Importance for Performance Prediction:\n');
fprintf('Competency          | Correlation | High Mean | Low Mean  | Difference | Importance\n');
fprintf('--------------------|-------------|-----------|-----------|------------|------------\n');

for j = 1:length(valid_comp_cols)
    comp_name = valid_comp_cols{j};
    comp_scores = matched_comp{:, valid_comp_indices(j)};

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

        % Importance score (correlation + effect size)
        importance = abs(correlation) * 0.7 + abs(difference/std(valid_comp_scores)) * 0.3;

        fprintf('%-19s | %11.3f | %9.1f | %9.1f | %10.1f | %10.3f\n', ...
                comp_name, correlation, high_mean, low_mean, difference, importance);

        comp_correlations = [comp_correlations; correlation];
        comp_importance = [comp_importance; importance];
    else
        comp_correlations = [comp_correlations; 0];
        comp_importance = [comp_importance; 0];
    end
end

% Select top 5 competencies
[~, top_idx] = sort(comp_importance, 'descend');
top_5_idx = top_idx(1:min(5, length(top_idx)));
top_5_competencies = valid_comp_cols(top_5_idx);
top_5_weights = comp_importance(top_5_idx);
top_5_weights_normalized = top_5_weights / sum(top_5_weights);

fprintf('\nTop 5 Key Competencies:\n');
for i = 1:length(top_5_competencies)
    fprintf('  %d. %s: %.1f%%\n', i, top_5_competencies{i}, top_5_weights_normalized(i)*100);
end

%% 8. Performance-Ordered Radar Charts
fprintf('\n=== Creating Performance-Ordered Radar Charts ===\n');

% Sort talent types by performance ranking
type_performance = [];
for i = 1:length(unique_types)
    if performance_ranking.isKey(unique_types{i})
        type_performance = [type_performance; performance_ranking(unique_types{i})];
    else
        type_performance = [type_performance; 0];
    end
end

[~, perf_sort_idx] = sort(type_performance, 'descend');
sorted_types = unique_types(perf_sort_idx);
sorted_performance = type_performance(perf_sort_idx);

% Performance color palette (high to low performance)
performance_colors = [
    0.0000 0.4470 0.7410;  % Dark blue (highest)
    0.0940 0.6940 0.1250;  % Dark green
    0.4940 0.1840 0.5560;  % Dark purple
    0.8500 0.3250 0.0980;  % Dark orange
    0.6350 0.0780 0.1840;  % Dark red
    0.4660 0.6740 0.1880;  % Olive
    0.3010 0.7450 0.9330;  % Sky blue
    0.9290 0.6940 0.1250;  % Yellow (lowest)
];

% Calculate overall means for baseline
overall_means = nanmean(matched_comp{:, valid_comp_indices}, 1);

% Create main radar chart figure
figure('Position', [50 50 1600 1200], 'Color', 'white', 'Name', 'Performance-Ordered Talent Type Competency Profiles');

n_types = length(sorted_types);
n_cols = 3;
n_rows = ceil(n_types / n_cols);

fprintf('Creating radar charts for %d talent types...\n', n_types);

for i = 1:n_types
    subplot(n_rows, n_cols, i);

    talent_type = sorted_types{i};
    perf_rank = sorted_performance(i);
    type_idx = strcmp(matched_talent_types, talent_type);

    % Calculate mean competency scores for this talent type
    type_comp_data = matched_comp{type_idx, valid_comp_indices};
    type_means = nanmean(type_comp_data, 1);

    % Use top 5 competencies only
    radar_data = type_means(top_5_idx);
    radar_baseline = overall_means(top_5_idx);

    % Normalize to 0-100 scale
    radar_data = max(0, min(100, radar_data));
    radar_baseline = max(0, min(100, radar_baseline));

    % Draw radar chart
    create_performance_radar_en(radar_data, radar_baseline, top_5_competencies, ...
                              talent_type, perf_rank, performance_colors(min(i,8), :), sum(type_idx));
end

sgtitle('Talent Type Competency Profiles (Performance Order: High to Low)', ...
        'FontSize', 18, 'FontWeight', 'bold');

%% 9. Data Usage Summary
fprintf('\n=== Data Usage Summary ===\n');
fprintf('Original HR data: %d people\n', height(hr_data));
fprintf('With talent types: %d people (%.1f%%)\n', height(hr_clean), height(hr_clean)/height(hr_data)*100);
fprintf('Final analysis: %d people (%.1f%%)\n', length(matched_ids), length(matched_ids)/height(hr_data)*100);
fprintf('\nData loss reasons:\n');
fprintf('  1. Missing talent type: %d people (%.1f%%)\n', ...
        height(hr_data)-height(hr_clean), (height(hr_data)-height(hr_clean))/height(hr_data)*100);
fprintf('  2. No competency test: %d people (%.1f%%)\n', ...
        height(hr_clean)-length(matched_ids), (height(hr_clean)-length(matched_ids))/height(hr_data)*100);

%% 10. Save Results
fprintf('\n=== Saving Results ===\n');

% Performance summary table
performance_summary = table();
performance_summary.Rank = (1:length(sorted_types))';
performance_summary.TalentType = sorted_types;
performance_summary.PerformanceScore = sorted_performance;

for i = 1:length(sorted_types)
    talent_type = sorted_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);
    performance_summary.Count(i) = sum(type_idx);

    type_comp_data = matched_comp{type_idx, valid_comp_indices};
    performance_summary.AvgCompetency(i) = nanmean(type_comp_data(:));

    type_total = total_scores(type_idx);
    performance_summary.AvgTotalScore(i) = nanmean(type_total);
end

% Save to Excel
try
    writetable(performance_summary, 'performance_ordered_analysis_en.xlsx', 'Sheet', 'PerformanceSummary');

    top_comp_table = table();
    top_comp_table.Rank = (1:length(top_5_competencies))';
    top_comp_table.Competency = top_5_competencies';
    top_comp_table.Weight = top_5_weights_normalized;
    top_comp_table.WeightPercent = top_5_weights_normalized * 100;
    top_comp_table.Correlation = comp_correlations(top_5_idx);

    writetable(top_comp_table, 'performance_ordered_analysis_en.xlsx', 'Sheet', 'TopCompetencies');

    fprintf('Excel file saved: performance_ordered_analysis_en.xlsx\n');
catch
    fprintf('Excel save failed\n');
end

% Save MATLAB results
results = struct();
results.performance_summary = performance_summary;
results.top_competencies = top_5_competencies;
results.competency_weights = top_5_weights_normalized;
results.sorted_types = sorted_types;
results.data_usage_efficiency = length(matched_ids)/height(hr_data)*100;

save('performance_ordered_results_en.mat', 'results');
fprintf('MATLAB file saved: performance_ordered_results_en.mat\n');

%% 11. Final Summary
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 80));
fprintf('                 Performance-Ordered Analysis Complete\n');
fprintf('%s\n', repmat('=', 1, 80));

fprintf('\nPerformance Ranking (High to Low):\n');
for i = 1:length(sorted_types)
    type_count = sum(strcmp(matched_talent_types, sorted_types{i}));
    fprintf('  %d. %s (%d people, score: %d)\n', i, sorted_types{i}, type_count, sorted_performance(i));
end

fprintf('\nTop 5 Predictive Competencies:\n');
for i = 1:length(top_5_competencies)
    fprintf('  %d. %s: %.1f%%\n', i, top_5_competencies{i}, top_5_weights_normalized(i)*100);
end

fprintf('\nData Usage Analysis:\n');
fprintf('  Efficiency: %.1f%% of original data used\n', length(matched_ids)/height(hr_data)*100);
fprintf('  Main loss: %.1f%% missing talent types, %.1f%% no competency test\n', ...
        (height(hr_data)-height(hr_clean))/height(hr_data)*100, ...
        (height(hr_clean)-length(matched_ids))/height(hr_data)*100);

fprintf('\nRadar charts created in performance order!\n');
fprintf('%s\n', repmat('=', 1, 80));

%% Helper Function: Performance Radar Chart (English)
function create_performance_radar_en(data, baseline, labels, title_text, perf_rank, color, count)
    % Calculate angles
    N = length(data);
    theta = linspace(0, 2*pi, N+1);

    % Prepare data (close the circle)
    data_plot = [data, data(1)];
    baseline_plot = [baseline, baseline(1)];

    % Convert to Cartesian coordinates
    [x_data, y_data] = pol2cart(theta, data_plot);
    [x_base, y_base] = pol2cart(theta, baseline_plot);

    % Plot
    hold on;

    % Baseline (overall average)
    plot(x_base, y_base, '--', 'Color', [0.5 0.5 0.5], 'LineWidth', 1.5);
    fill(x_base, y_base, [0.9 0.9 0.9], 'FaceAlpha', 0.3, 'EdgeColor', 'none');

    % Talent type data
    plot(x_data, y_data, '-', 'Color', color, 'LineWidth', 3);
    fill(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2);

    % Axis settings
    axis equal;
    max_val = max([max(data), max(baseline)]) * 1.1;
    xlim([-max_val, max_val]);
    ylim([-max_val, max_val]);

    % Draw grid
    for r = 20:20:100
        [x_circle, y_circle] = pol2cart(theta, repmat(r, size(theta)));
        plot(x_circle, y_circle, ':', 'Color', [0.7 0.7 0.7], 'LineWidth', 0.5);
    end

    % Draw axis lines
    for i = 1:N
        plot([0, max_val*cos(theta(i))], [0, max_val*sin(theta(i))], ':', 'Color', [0.7 0.7 0.7]);

        % Add labels
        label_r = max_val * 1.15;
        text(label_r*cos(theta(i)), label_r*sin(theta(i)), labels{i}, ...
             'HorizontalAlignment', 'center', 'VerticalAlignment', 'middle', ...
             'FontSize', 9, 'FontWeight', 'bold');
    end

    % Title
    title(sprintf('%s\n(Rank: %d, N=%d)', title_text, perf_rank, count), ...
          'FontSize', 12, 'FontWeight', 'bold', 'Color', color);

    % Legend
    legend({'Overall Avg', '', sprintf('%s', title_text), ''}, ...
           'Location', 'best', 'FontSize', 8);

    axis off;
    hold off;
end