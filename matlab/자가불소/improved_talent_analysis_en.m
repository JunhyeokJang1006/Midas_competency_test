%% Improved Talent Type Competency Analysis (English Version)
% Using readtable and proper sheet selection for better efficiency

clear; clc;

%% 1. Data Loading - Improved Method
fprintf('=== Improved Data Loading ===\n');

% HR data
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');

% Competency data - Upper level items (Sheet 3)
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
comp_upper = readtable(comp_file, 'Sheet', 3, 'VariableNamingRule', 'preserve');

% Competency data - Total score (Sheet 4)
comp_total = readtable(comp_file, 'Sheet', 4, 'VariableNamingRule', 'preserve');

fprintf('HR data: %d rows x %d columns\n', height(hr_data), width(hr_data));
fprintf('Upper level data: %d rows x %d columns\n', height(comp_upper), width(comp_upper));
fprintf('Total score data: %d rows x %d columns\n', height(comp_total), width(comp_total));

%% 2. Clean Talent Type Data
fprintf('\n=== Clean Talent Type Data ===\n');

% Find talent type column
talent_col_idx = find(contains(hr_data.Properties.VariableNames, '인재유형'));
if isempty(talent_col_idx)
    fprintf('Warning: Talent type column not found\n');
    return;
end

talent_col_name = hr_data.Properties.VariableNames{talent_col_idx};
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

%% 3. Upper Level Competency Analysis
fprintf('\n=== Upper Level Competency Analysis ===\n');

% Match IDs
hr_ids = hr_clean.ID;
comp_upper_ids = comp_upper.ID;

[matched_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_upper_ids);
fprintf('Matched IDs with upper level data: %d\n', length(matched_ids));

if ~isempty(matched_ids)
    % Extract matched data
    matched_hr = hr_clean(hr_idx, :);
    matched_comp_upper = comp_upper(comp_idx, :);

    % Find competency columns (numeric data starting from column 6)
    comp_cols = {};
    comp_col_indices = [];

    for i = 6:width(matched_comp_upper)
        col_name = matched_comp_upper.Properties.VariableNames{i};
        col_data = matched_comp_upper{:, i};

        % Select numeric columns with less than 50% missing values
        if isnumeric(col_data) && sum(~isnan(col_data)) >= length(col_data) * 0.5
            comp_cols{end+1} = col_name;
            comp_col_indices = [comp_col_indices, i];
        end
    end

    fprintf('Valid upper level competencies: %d\n', length(comp_cols));
    fprintf('Upper level competencies:\n');
    for i = 1:length(comp_cols)
        fprintf('  %d. %s\n', i, comp_cols{i});
    end

    % Extract upper level scores
    upper_scores = matched_comp_upper{:, comp_col_indices};

    %% 4. Total Score Analysis
    fprintf('\n=== Total Score Analysis ===\n');

    % Match with total scores
    comp_total_ids = comp_total.ID;
    [total_matched_ids, hr_total_idx, comp_total_idx] = intersect(hr_ids, comp_total_ids);
    fprintf('Matched IDs with total score: %d\n', length(total_matched_ids));

    if ~isempty(total_matched_ids)
        matched_hr_total = hr_clean(hr_total_idx, :);
        matched_comp_total = comp_total(comp_total_idx, :);

        % Find total score column (usually the last numeric column)
        total_score_col = width(matched_comp_total);
        total_scores = matched_comp_total{:, total_score_col};

        fprintf('Total score column: %s\n', matched_comp_total.Properties.VariableNames{total_score_col});

        %% 5. Total Score by Talent Type
        fprintf('\n=== Total Score Analysis by Talent Type ===\n');

        talent_types_total = matched_hr_total{:, talent_col_name};
        unique_types_total = unique(talent_types_total(~cellfun(@isempty, talent_types_total)));

        fprintf('Total Score Statistics by Talent Type:\n');
        fprintf('Talent Type         | Mean     | Std      | Max      | Min      | Count\n');
        fprintf('--------------------|----------|----------|----------|----------|------\n');

        type_stats = [];
        for i = 1:length(unique_types_total)
            talent_type = unique_types_total{i};
            type_idx = strcmp(talent_types_total, talent_type) & ~isnan(total_scores);

            if sum(type_idx) > 0
                type_scores = total_scores(type_idx);

                mean_score = mean(type_scores);
                std_score = std(type_scores);
                max_score = max(type_scores);
                min_score = min(type_scores);
                count = length(type_scores);

                fprintf('%-19s | %8.2f | %8.2f | %8.2f | %8.2f | %5d\n', ...
                        talent_type, mean_score, std_score, max_score, min_score, count);

                type_stats = [type_stats; {talent_type, mean_score, std_score, max_score, min_score, count}];
            end
        end
    end

    %% 6. Upper Level Competency vs Performance Analysis
    fprintf('\n=== Upper Level Competency vs Performance Analysis ===\n');

    % Define performance ranking
    performance_ranking = containers.Map();
    performance_ranking('자연성') = 8;
    performance_ranking('성실한 가연성') = 7;
    performance_ranking('유익한 불연성') = 6;
    performance_ranking('유능한 불연성') = 5;
    performance_ranking('게으른 가연성') = 4;
    performance_ranking('무능한 불연성') = 3;
    performance_ranking('위장형 소화성') = 2;
    performance_ranking('소화성') = 1;

    % Get performance scores for matched data
    talent_types_upper = matched_hr{:, talent_col_name};
    performance_scores = zeros(length(talent_types_upper), 1);

    for i = 1:length(talent_types_upper)
        talent_type = talent_types_upper{i};
        if performance_ranking.isKey(talent_type)
            performance_scores(i) = performance_ranking(talent_type);
        end
    end

    % Analyze correlation between upper level competencies and performance
    fprintf('Upper Level Competency Performance Prediction Analysis:\n');
    fprintf('Competency      | Correlation | High Mean | Low Mean  | Diff   | Effect Size\n');
    fprintf('----------------|-------------|-----------|-----------|--------|------------\n');

    correlations = [];
    competency_importance = [];

    for i = 1:length(comp_cols)
        comp_name = comp_cols{i};
        comp_scores = upper_scores(:, i);
        valid_idx = ~isnan(comp_scores) & performance_scores > 0;

        if sum(valid_idx) >= 10
            valid_comp_scores = comp_scores(valid_idx);
            valid_perf_scores = performance_scores(valid_idx);

            % Calculate correlation
            if var(valid_comp_scores) > 0 && var(valid_perf_scores) > 0
                correlation = sum((valid_comp_scores - mean(valid_comp_scores)) .* ...
                                (valid_perf_scores - mean(valid_perf_scores))) / ...
                             (sqrt(sum((valid_comp_scores - mean(valid_comp_scores)).^2)) * ...
                              sqrt(sum((valid_perf_scores - mean(valid_perf_scores)).^2)));
            else
                correlation = 0;
            end

            % High vs Low performance group comparison
            high_perf_threshold = median(valid_perf_scores);
            high_idx = valid_perf_scores > high_perf_threshold;
            low_idx = valid_perf_scores <= high_perf_threshold;

            high_mean = mean(valid_comp_scores(high_idx));
            low_mean = mean(valid_comp_scores(low_idx));
            difference = high_mean - low_mean;

            % Effect size (Cohen's d)
            pooled_std = sqrt(((sum(high_idx)-1)*var(valid_comp_scores(high_idx)) + ...
                              (sum(low_idx)-1)*var(valid_comp_scores(low_idx))) / ...
                             (sum(high_idx) + sum(low_idx) - 2));
            effect_size = difference / (pooled_std + eps);

            fprintf('%-15s | %11.3f | %9.2f | %9.2f | %6.2f | %11.3f\n', ...
                    comp_name, correlation, high_mean, low_mean, difference, effect_size);

            correlations = [correlations; correlation];
            competency_importance = [competency_importance; abs(correlation) * 0.6 + abs(effect_size) * 0.4];
        end
    end

    %% 7. Final Weights and Recommendations
    fprintf('\n=== Upper Level Competency Weights for Performance Prediction ===\n');

    % Sort by importance
    [sorted_importance, sort_idx] = sort(competency_importance, 'descend');

    fprintf('Rank | Competency      | Weight | Correlation | Effect Size\n');
    fprintf('-----|-----------------|--------|-------------|------------\n');

    top_competencies = {};
    top_weights = [];

    for i = 1:min(5, length(sort_idx))
        idx = sort_idx(i);
        comp_name = comp_cols{idx};
        weight = sorted_importance(i);
        correlation = correlations(idx);

        % Recalculate effect size
        comp_scores = upper_scores(:, idx);
        valid_idx = ~isnan(comp_scores) & performance_scores > 0;
        valid_comp_scores = comp_scores(valid_idx);
        valid_perf_scores = performance_scores(valid_idx);

        high_perf_threshold = median(valid_perf_scores);
        high_idx = valid_perf_scores > high_perf_threshold;
        low_idx = valid_perf_scores <= high_perf_threshold;

        high_mean = mean(valid_comp_scores(high_idx));
        low_mean = mean(valid_comp_scores(low_idx));
        difference = high_mean - low_mean;

        pooled_std = sqrt(((sum(high_idx)-1)*var(valid_comp_scores(high_idx)) + ...
                          (sum(low_idx)-1)*var(valid_comp_scores(low_idx))) / ...
                         (sum(high_idx) + sum(low_idx) - 2));
        effect_size = difference / (pooled_std + eps);

        fprintf('%4d | %-15s | %6.3f | %11.3f | %11.3f\n', ...
                i, comp_name, weight, correlation, effect_size);

        top_competencies{end+1} = comp_name;
        top_weights = [top_weights; weight];
    end

    % Normalize weights (sum to 1)
    if ~isempty(top_weights)
        top_weights_normalized = top_weights / sum(top_weights);

        fprintf('\n=== Final Performance Prediction Weights (Normalized) ===\n');
        for i = 1:length(top_competencies)
            fprintf('%d. %s: %.1f%%\n', i, top_competencies{i}, top_weights_normalized(i) * 100);
        end
    end

    %% 8. Save Results
    fprintf('\n=== Save Results ===\n');

    % Create results structure
    results = struct();
    results.type_stats = type_stats;
    results.top_competencies = top_competencies;
    results.top_weights_normalized = top_weights_normalized;
    results.correlations = correlations;
    results.comp_cols = comp_cols;
    results.analysis_summary = sprintf('Analyzed %d people, %d upper level competencies', ...
        length(matched_ids), length(comp_cols));

    % Save to file
    save('improved_talent_analysis_results.mat', 'results');

    % Save to Excel
    try
        % Talent type statistics table
        type_table = cell2table(type_stats, 'VariableNames', ...
            {'TalentType', 'MeanScore', 'StdScore', 'MaxScore', 'MinScore', 'Count'});
        writetable(type_table, 'talent_type_statistics.xlsx', 'Sheet', 'TalentTypeStats');

        % Upper level competency weights table
        weight_table = table();
        weight_table.Rank = (1:length(top_competencies))';
        weight_table.Competency = top_competencies';
        weight_table.Weight = top_weights_normalized;
        weight_table.WeightPercent = top_weights_normalized * 100;
        writetable(weight_table, 'talent_type_statistics.xlsx', 'Sheet', 'CompetencyWeights');

        fprintf('Excel file saved: talent_type_statistics.xlsx\n');
    catch
        fprintf('Excel save failed (Excel may not be installed)\n');
    end

    fprintf('MATLAB file saved: improved_talent_analysis_results.mat\n');

else
    fprintf('Warning: No matching data found\n');
end

%% 9. Final Summary
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 70));
fprintf('                Improved Talent Type Analysis Complete\n');
fprintf('%s\n', repmat('=', 1, 70));

if exist('type_stats', 'var') && ~isempty(type_stats)
    fprintf('\nTalent Type Total Score Rankings:\n');
    % Sort by mean score
    type_scores = cell2mat(type_stats(:, 2));
    [~, rank_idx] = sort(type_scores, 'descend');

    for i = 1:length(rank_idx)
        idx = rank_idx(i);
        fprintf('  %d. %s: %.2f points (%d people)\n', i, type_stats{idx, 1}, ...
                type_stats{idx, 2}, type_stats{idx, 6});
    end
end

if exist('top_competencies', 'var') && ~isempty(top_competencies)
    fprintf('\nKey Upper Level Competencies for Performance Prediction:\n');
    for i = 1:length(top_competencies)
        fprintf('  %d. %s: %.1f%%\n', i, top_competencies{i}, ...
                top_weights_normalized(i) * 100);
    end
end

fprintf('\nAnalysis complete! More efficient analysis using readtable and sheet-specific access.\n');
fprintf('%s\n', repmat('=', 1, 70));