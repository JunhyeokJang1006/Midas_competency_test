%% MIDAS Growth Stage Evaluation System (Claude Optimized Version)
% Date: 2024
% Author: Claude (Complete refactoring of existing code)
% Purpose: Calculate growth scores based on competency evaluation data

clear; clc; close all;

%% ========== 1. Program Initialization ==========
fprintf('===============================================================\n');
fprintf('    MIDAS Growth Stage Analysis System (Claude Optimized)    \n');
fprintf('===============================================================\n');

start_time = datetime('now');
fprintf('Analysis Start: %s\n', datestr(start_time, 'yyyy-mm-dd HH:MM:SS'));

% Global Configuration
config = struct();
config.data_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
config.target_sheet = 3;  % Competency evaluation sheet
config.id_column = 'ID';
config.period_labels = {'23H1', '23H2', '24H1', '24H2', '25H1'};
config.verbose = true;

%% ========== 2. Data Loading and Validation ==========
fprintf('\n[Step 1/6] Data Loading Started\n');

try
    % Check file existence
    if ~exist(config.data_file, 'file')
        error('Data file not found: %s', config.data_file);
    end

    % Check sheet information
    [~, sheet_names] = xlsfinfo(config.data_file);
    if length(sheet_names) < config.target_sheet
        error('Target sheet (%d) does not exist. Found %d sheets', ...
              config.target_sheet, length(sheet_names));
    end

    target_sheet_name = sheet_names{config.target_sheet};
    fprintf('  Target Sheet: %s (Sheet %d)\n', target_sheet_name, config.target_sheet);

    % Load data with multiple methods
    [data_table, data_source] = load_excel_data(config.data_file, target_sheet_name);
    fprintf('  Data Load Complete (Method: %s)\n', data_source);
    fprintf('  Data Size: %d rows x %d columns\n', height(data_table), width(data_table));

catch ME
    fprintf('  Data Loading Failed: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 3. Data Preprocessing ==========
fprintf('\n[Step 2/6] Data Preprocessing\n');

try
    % Column analysis
    column_names = data_table.Properties.VariableNames;
    fprintf('  Column Analysis: %s\n', strjoin(column_names(1:min(5, end)), ', '));

    % Find ID column
    id_col_idx = find_column_index(column_names, config.id_column);
    if id_col_idx == 0
        error('ID column not found');
    end
    fprintf('  ID Column: %s (Column %d)\n', column_names{id_col_idx}, id_col_idx);

    % Find evaluation columns
    eval_col_indices = find_evaluation_columns(column_names);
    if isempty(eval_col_indices)
        error('Evaluation columns not found');
    end
    fprintf('  Found %d Evaluation Columns: [%s]\n', length(eval_col_indices), ...
            num2str(eval_col_indices));

    % Extract and clean data
    employee_ids = data_table{:, id_col_idx};
    evaluation_data = data_table{:, eval_col_indices};

    % Filter valid data only
    [clean_ids, clean_evals, valid_mask] = clean_data(employee_ids, evaluation_data);

    n_employees = length(clean_ids);
    n_periods = size(clean_evals, 2);

    fprintf('  Data Cleaning Complete: %d employees x %d periods\n', n_employees, n_periods);
    fprintf('  Valid Data Rate: %.1f%% (%d/%d)\n', ...
            sum(valid_mask)/length(valid_mask)*100, sum(valid_mask), length(valid_mask));

    % Set period labels
    period_labels = get_period_labels(column_names, eval_col_indices, config.period_labels, n_periods);
    fprintf('  Evaluation Periods: %s\n', strjoin(period_labels, ' -> '));

catch ME
    fprintf('  Data Preprocessing Failed: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 4. Growth Score Calculation ==========
fprintf('\n[Step 3/6] Growth Score Calculation\n');

try
    % Start calculation
    calculation_start = tic;

    % Initialize result variables
    growth_scores = zeros(n_employees, 1);
    growth_patterns = cell(n_employees, 1);
    score_details = cell(n_employees, 1);

    % Progress display setup
    progress_step = max(1, floor(n_employees / 20));

    fprintf('  Progress: ');
    for i = 1:n_employees
        % Calculate individual growth score
        [score, pattern, details] = calculate_growth_score(clean_evals(i, :), period_labels);

        growth_scores(i) = score;
        growth_patterns{i} = pattern;
        score_details{i} = details;

        % Display progress
        if mod(i, progress_step) == 0 || i == n_employees
            fprintf('*');
        end
    end

    calculation_time = toc(calculation_start);
    fprintf(' Complete!\n');
    fprintf('  Calculation Complete (Time: %.2f seconds)\n', calculation_time);

catch ME
    fprintf('  Growth Score Calculation Failed: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 5. Statistical Analysis ==========
fprintf('\n[Step 4/6] Statistical Analysis\n');

try
    % Basic statistics
    stats = calculate_statistics(growth_scores);

    % Pattern analysis
    [pattern_stats, pattern_distribution] = analyze_patterns(growth_patterns);

    % Print results
    print_statistics(stats, pattern_stats, n_employees);

catch ME
    fprintf('  Statistical Analysis Failed: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 6. Visualization ==========
fprintf('\n[Step 5/6] Visualization Generation\n');

try
    create_visualizations(growth_scores, growth_patterns, score_details, ...
                         period_labels, stats, pattern_distribution);
    fprintf('  Visualization Complete (8 charts generated)\n');

catch ME
    fprintf('  Visualization Generation Failed: %s\n', ME.message);
    % Continue even if visualization fails
end

%% ========== 7. Save Results ==========
fprintf('\n[Step 6/6] Saving Results\n');

try
    % Generate filename
    timestamp = datestr(now, 'yyyymmdd_HHMMSS');

    % Save basic results
    basic_results = table(clean_ids, growth_scores, growth_patterns, ...
        'VariableNames', {'ID', 'GrowthScore', 'Pattern'});
    basic_filename = sprintf('MIDAS_Growth_Results_%s.xlsx', timestamp);
    writetable(basic_results, basic_filename);
    fprintf('  Basic Results: %s\n', basic_filename);

    % Save detailed results
    detailed_results = create_detailed_results(clean_ids, clean_evals, growth_scores, ...
                                              growth_patterns, score_details, period_labels);
    detailed_filename = sprintf('MIDAS_Growth_Detailed_%s.xlsx', timestamp);
    writetable(detailed_results, detailed_filename);
    fprintf('  Detailed Results: %s\n', detailed_filename);

    % Save MATLAB data
    mat_filename = sprintf('MIDAS_growth_analysis_%s.mat', timestamp);
    save(mat_filename, 'basic_results', 'detailed_results', 'stats', ...
         'pattern_stats', 'score_details', 'config', 'period_labels');
    fprintf('  Data File: %s\n', mat_filename);

catch ME
    fprintf('  Result Saving Failed: %s\n', ME.message);
    % Analysis results remain in memory even if saving fails
end

%% ========== 8. Program Completion ==========
total_time = seconds(datetime('now') - start_time);
fprintf('\n===============================================================\n');
fprintf('              Analysis Complete - Total Time: %.1f seconds\n', total_time);
fprintf('===============================================================\n');

% Summary statistics output
fprintf('\nFinal Summary:\n');
fprintf('   Analyzed: %d employees\n', n_employees);
fprintf('   Average Score: %.1f points\n', stats.mean);
fprintf('   Highest Score: %.0f points (Pattern: %s)\n', stats.max, pattern_stats{1, 1});
fprintf('   Main Pattern: %s (%.1f%%)\n', pattern_distribution.pattern{1}, pattern_distribution.percentage(1));

%% ========== Helper Functions ==========

function [data_table, data_source] = load_excel_data(filename, sheet_name)
    % Load Excel data with multiple methods
    data_table = [];
    data_source = '';

    % Method 1: readtable (recommended)
    try
        data_table = readtable(filename, 'Sheet', sheet_name, ...
                              'VariableNamingRule', 'preserve', ...
                              'TextType', 'string');
        data_source = 'readtable';
        return;
    catch ME1
        % Method 1 failed
    end

    % Method 2: xlsread + table conversion
    try
        [~, ~, raw_data] = xlsread(filename, sheet_name);
        headers = raw_data(1, :);
        data_body = raw_data(2:end, :);

        % Select valid headers only
        valid_headers = {};
        valid_cols = [];
        for i = 1:length(headers)
            if ~isempty(headers{i}) && (ischar(headers{i}) || isstring(headers{i}))
                valid_headers{end+1} = char(headers{i});
                valid_cols(end+1) = i;
            elseif isnumeric(headers{i}) && ~isnan(headers{i})
                valid_headers{end+1} = sprintf('Var%d', i);
                valid_cols(end+1) = i;
            end
        end

        if ~isempty(valid_cols)
            data_table = table();
            for i = 1:length(valid_cols)
                col_data = data_body(:, valid_cols(i));
                data_table.(valid_headers{i}) = col_data;
            end
            data_source = 'xlsread';
            return;
        end
    catch ME2
        % Method 2 failed
    end

    % Method 3: readmatrix (last resort)
    try
        raw_matrix = readmatrix(filename, 'Sheet', sheet_name);
        n_cols = size(raw_matrix, 2);
        data_table = table();
        for i = 1:n_cols
            data_table.(sprintf('Var%d', i)) = raw_matrix(:, i);
        end
        data_source = 'readmatrix';
        return;
    catch ME3
        % All methods failed
    end

    error('All data loading methods failed');
end

function col_idx = find_column_index(column_names, target_name)
    % Find column index (case insensitive)
    col_idx = 0;
    keywords = {target_name, 'id', 'ID'};

    for i = 1:length(column_names)
        col_name = char(column_names{i});
        for j = 1:length(keywords)
            if contains(col_name, keywords{j}, 'IgnoreCase', true)
                col_idx = i;
                return;
            end
        end
    end
end

function eval_indices = find_evaluation_columns(column_names)
    % Find evaluation columns
    eval_indices = [];
    keywords = {'H1', 'H2', 'h1', 'h2'};

    for i = 1:length(column_names)
        col_name = char(column_names{i});
        for j = 1:length(keywords)
            if contains(col_name, keywords{j}, 'IgnoreCase', true)
                eval_indices(end+1) = i;
                break;
            end
        end
    end

    % Remove duplicates and sort
    eval_indices = unique(eval_indices);
end

function [clean_ids, clean_evals, valid_mask] = clean_data(employee_ids, evaluation_data)
    % Data cleaning and validation
    n_rows = length(employee_ids);
    valid_mask = true(n_rows, 1);

    % ID validity check
    for i = 1:n_rows
        id_val = employee_ids(i);
        if iscell(id_val)
            id_val = id_val{1};
        end

        if isempty(id_val) || (isnumeric(id_val) && isnan(id_val)) || ...
           (ischar(id_val) && isempty(strtrim(id_val)))
            valid_mask(i) = false;
        end
    end

    % Extract valid data only
    clean_ids = employee_ids(valid_mask);
    clean_evals = evaluation_data(valid_mask, :);

    % Convert cell array IDs to numeric if needed
    if iscell(clean_ids)
        numeric_ids = zeros(size(clean_ids));
        for i = 1:length(clean_ids)
            if isnumeric(clean_ids{i})
                numeric_ids(i) = clean_ids{i};
            else
                numeric_ids(i) = str2double(clean_ids{i});
            end
        end
        clean_ids = numeric_ids;
    end
end

function period_labels = get_period_labels(column_names, eval_indices, default_labels, n_periods)
    % Generate period labels
    period_labels = cell(1, n_periods);

    for i = 1:n_periods
        if i <= length(eval_indices)
            col_name = char(column_names{eval_indices(i)});
            % Extract period info from column name (e.g., "23H1 evaluation" -> "23H1")
            tokens = regexp(col_name, '\d+H\d+', 'match');
            if ~isempty(tokens)
                period_labels{i} = tokens{1};
            else
                period_labels{i} = sprintf('P%d', i);
            end
        elseif i <= length(default_labels)
            period_labels{i} = default_labels{i};
        else
            period_labels{i} = sprintf('Period%d', i);
        end
    end
end

function [score, pattern, details] = calculate_growth_score(evaluations, period_labels)
    % Calculate individual employee growth score
    score = 0;
    base_level = 'Open_Lv2';
    consecutive_achievement = 0;

    n_periods = length(evaluations);
    score_history = zeros(1, n_periods);
    base_history = cell(1, n_periods);

    for period = 1:n_periods
        % Safely extract evaluation data
        if period <= length(evaluations)
            try
                eval_data = evaluations{period};
            catch
                eval_data = [];
            end
        else
            eval_data = [];
        end

        current_eval = normalize_evaluation(eval_data);
        period_score = 0;

        % Calculate score based on evaluation result
        if isempty(current_eval) || contains(current_eval, '#N/A')
            % No data
            period_score = 0;

        elseif contains(current_eval, 'Responsibility') || contains(current_eval, '책임')
            % Responsibility level achievement
            if ~strcmp(base_level, 'Responsibility_Lv1')
                period_score = 30;  % First responsibility achievement
                base_level = 'Responsibility_Lv1';
                consecutive_achievement = 0;
            else
                period_score = 0;   % Already at responsibility level
            end

        elseif contains(current_eval, 'Achievement') || contains(current_eval, '성취')
            % Achievement level
            if strcmp(base_level, 'Open_Lv2')
                consecutive_achievement = consecutive_achievement + 1;
                period_score = 10;

                % Promote to Achievement_Lv1 after 3 consecutive achievements
                if consecutive_achievement >= 3
                    base_level = 'Achievement_Lv1';
                    consecutive_achievement = 0;
                end
            else
                period_score = 0;   % Already at higher level
            end

        elseif contains(current_eval, 'Open') && contains(current_eval, 'Lv2') || contains(current_eval, '열린(Lv2)')
            % Open Lv2 maintenance or demotion
            if strcmp(base_level, 'Open_Lv2')
                period_score = 0;   % Status quo
                consecutive_achievement = 0;
            elseif contains(base_level, 'Achievement') || contains(base_level, 'Responsibility')
                period_score = -10; % Demotion penalty
                base_level = 'Open_Lv2';
                consecutive_achievement = 0;
            end

        elseif contains(current_eval, 'Open') && contains(current_eval, 'Lv1') || contains(current_eval, '열린(Lv1)')
            % Open Lv1 - regression
            period_score = -5;
            base_level = 'Open_Lv1';
            consecutive_achievement = 0;
        end

        score = score + period_score;
        score_history(period) = period_score;
        base_history{period} = base_level;
    end

    % Classify growth pattern
    if score >= 30
        pattern = 'High_Growth';
    elseif score >= 20
        pattern = 'Medium_High_Growth';
    elseif score >= 10
        pattern = 'Medium_Growth';
    elseif score > 0
        pattern = 'Low_Growth';
    elseif score == 0
        pattern = 'Stagnation';
    else
        pattern = 'Regression';
    end

    % Save detailed information
    details = struct();
    details.score_history = score_history;
    details.base_history = base_history;
    details.final_base_level = base_level;
    details.consecutive_count = consecutive_achievement;
end

function normalized = normalize_evaluation(eval_data)
    % Normalize evaluation data
    normalized = '#N/A';  % Default value

    try
        % Handle empty input
        if isempty(eval_data)
            return;
        end

        % Handle cell arrays
        if iscell(eval_data)
            if isempty(eval_data{1})
                return;
            end
            eval_data = eval_data{1};
        end

        % Handle missing values (MATLAB specific)
        try
            if isa(eval_data, 'missing') || ismissing(eval_data)
                return;
            end
        catch
            % If ismissing fails, continue with other checks
        end

        % Handle different data types
        if isnumeric(eval_data)
            if isnan(eval_data)
                return;
            else
                normalized = num2str(eval_data);
            end
        elseif isstring(eval_data) || ischar(eval_data)
            eval_str = char(eval_data);
            if strcmp(eval_str, '<missing>') || isempty(strtrim(eval_str))
                return;
            else
                normalized = strtrim(eval_str);
            end
        end

    catch ME
        % If any error occurs, return default '#N/A'
        fprintf('Warning: Error normalizing evaluation data: %s\n', ME.message);
    end
end

function stats = calculate_statistics(scores)
    % Calculate basic statistics
    valid_scores = scores(scores ~= 0 | ~isnan(scores));

    if isempty(valid_scores)
        stats = struct('mean', 0, 'median', 0, 'std', 0, 'min', 0, ...
                      'max', 0, 'quartiles', [0 0 0], 'range', 0, 'cv', 0);
        return;
    end

    stats = struct();
    stats.mean = mean(valid_scores);
    stats.median = median(valid_scores);
    stats.std = std(valid_scores);
    stats.min = min(valid_scores);
    stats.max = max(valid_scores);
    stats.quartiles = quantile(valid_scores, [0.25, 0.5, 0.75]);
    stats.range = stats.max - stats.min;

    if stats.mean ~= 0
        stats.cv = abs(stats.std / stats.mean * 100);
    else
        stats.cv = 0;
    end
end

function [pattern_stats, pattern_distribution] = analyze_patterns(patterns)
    % Pattern analysis
    [unique_patterns, ~, pattern_idx] = unique(patterns);
    counts = accumarray(pattern_idx, 1);
    percentages = counts / length(patterns) * 100;

    % Pattern statistics
    pattern_stats = [unique_patterns, num2cell(counts), num2cell(percentages)];

    % Distribution table
    pattern_distribution = table(unique_patterns, counts, percentages, ...
        'VariableNames', {'pattern', 'count', 'percentage'});

    % Sort by frequency
    [~, sort_idx] = sort(counts, 'descend');
    pattern_distribution = pattern_distribution(sort_idx, :);
end

function print_statistics(stats, pattern_stats, n_total)
    % Print statistical results
    fprintf('  Basic Statistics:\n');
    fprintf('     Mean: %.2f  |  Median: %.2f  |  Std Dev: %.2f\n', ...
            stats.mean, stats.median, stats.std);
    fprintf('     Min: %.0f  |  Max: %.0f  |  Range: %.0f\n', ...
            stats.min, stats.max, stats.range);
    fprintf('     Q1: %.0f  |  Q2: %.0f  |  Q3: %.0f  |  CV: %.1f%%\n', ...
            stats.quartiles(1), stats.quartiles(2), stats.quartiles(3), stats.cv);

    fprintf('\n  Growth Pattern Distribution:\n');
    for i = 1:size(pattern_stats, 1)
        fprintf('     %s: %d employees (%.1f%%)\n', ...
                pattern_stats{i,1}, pattern_stats{i,2}, pattern_stats{i,3});
    end
end

function create_visualizations(scores, patterns, score_details, period_labels, stats, pattern_dist)
    % Create comprehensive visualization
    try
        figure('Position', [50, 50, 1200, 800]);

        % Create 4 main plots instead of 8 to avoid issues

        % 1. Score distribution histogram
        subplot(2, 2, 1);
        if ~isempty(scores) && sum(scores ~= 0) > 0
            valid_scores = scores(scores ~= 0);
            try
                histogram(valid_scores, 15);
                xlabel('Growth Score');
                ylabel('Count');
                title('Growth Score Distribution');
                grid on;
            catch
                hist(valid_scores, 15);
                xlabel('Growth Score');
                ylabel('Count');
                title('Growth Score Distribution');
                grid on;
            end
        else
            text(0.5, 0.5, 'No Valid Data', 'HorizontalAlignment', 'center');
            title('Growth Score Distribution');
        end

        % 2. Pattern distribution (pie chart)
        subplot(2, 2, 2);
        if height(pattern_dist) > 0
            try
                pie(pattern_dist.count, pattern_dist.pattern);
                title('Growth Pattern Distribution');
            catch
                bar(pattern_dist.count);
                title('Growth Pattern Distribution');
                set(gca, 'XTickLabel', pattern_dist.pattern);
                xtickangle(45);
            end
        else
            text(0.5, 0.5, 'No Pattern Data', 'HorizontalAlignment', 'center');
            title('Growth Pattern Distribution');
        end

        % 3. Box plot
        subplot(2, 2, 3);
        if ~isempty(scores) && sum(scores ~= 0) > 0
            valid_scores = scores(scores ~= 0);
            try
                boxplot(valid_scores);
                ylabel('Growth Score');
                title('Score Distribution (Box Plot)');
                grid on;
            catch
                hist(valid_scores, 10);
                ylabel('Count');
                xlabel('Growth Score');
                title('Score Distribution');
                grid on;
            end
        else
            text(0.5, 0.5, 'Insufficient Data', 'HorizontalAlignment', 'center');
            title('Score Distribution');
        end

        % 4. Score ranges
        subplot(2, 2, 4);
        if ~isempty(scores)
            ranges = [-inf, 0, 10, 20, 30, inf];
            labels = {'Regression', 'Stagnation', 'Low', 'Medium', 'High'};
            groups = discretize(scores, ranges);
            valid_groups = groups(~isnan(groups));
            if ~isempty(valid_groups)
                counts = accumarray(valid_groups, 1, [length(labels), 1]);
                bar(counts);
                set(gca, 'XTickLabel', labels);
                xtickangle(45);
                ylabel('Count');
                title('Score Range Distribution');
                grid on;
            end
        else
            text(0.5, 0.5, 'No Data', 'HorizontalAlignment', 'center');
            title('Score Range Distribution');
        end

        % Add overall title
        try
            sgtitle('MIDAS Growth Stage Analysis Results');
        catch
            % If sgtitle not available, skip
        end

    catch ME
        fprintf('  Visualization Error: %s\n', ME.message);
        rethrow(ME);
    end
end

function detailed_table = create_detailed_results(ids, evaluations, scores, patterns, details, period_labels)
    % Create detailed results table
    n_employees = length(ids);
    n_periods = size(evaluations, 2);

    % Basic information
    detailed_table = table();
    detailed_table.ID = ids;

    % Period-wise evaluation results
    for p = 1:n_periods
        col_name = sprintf('%s_Evaluation', period_labels{p});
        detailed_table.(col_name) = evaluations(:, p);

        col_name = sprintf('%s_Score', period_labels{p});
        period_scores = zeros(n_employees, 1);
        for i = 1:n_employees
            if ~isempty(details{i}) && isfield(details{i}, 'score_history') && ...
               length(details{i}.score_history) >= p
                period_scores(i) = details{i}.score_history(p);
            end
        end
        detailed_table.(col_name) = period_scores;
    end

    % Total score and pattern
    detailed_table.TotalScore = scores;
    detailed_table.GrowthPattern = patterns;

    % Final level
    final_levels = cell(n_employees, 1);
    for i = 1:n_employees
        if ~isempty(details{i}) && isfield(details{i}, 'final_base_level')
            final_levels{i} = details{i}.final_base_level;
        else
            final_levels{i} = 'Unknown';
        end
    end
    detailed_table.FinalLevel = final_levels;
end