%% Talent Type Competency Analysis and Weight Calculation
% Goal: Derive competency weights for performance prediction

clear; clc;

%% 1. Data Loading
fprintf('=== Starting Data Loading ===\n');

% HR data (including talent types)
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');

% Competency test data
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
[comp_num, comp_txt, comp_raw] = xlsread(comp_file, 1);

fprintf('HR data: %d rows x %d columns\n', height(hr_data), width(hr_data));
fprintf('Competency data: %d rows x %d columns\n', size(comp_raw, 1), size(comp_raw, 2));

%% 2. Extract Talent Type Data
fprintf('\n=== Extracting Talent Type Data ===\n');

% Filter rows with talent type information
talent_col_idx = find(contains(hr_data.Properties.VariableNames, '인재유형'));
if isempty(talent_col_idx)
    fprintf('Warning: Talent type column not found\n');
    return;
end

talent_col_name = hr_data.Properties.VariableNames{talent_col_idx};
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_name}), :);
fprintf('Employees with talent type data: %d\n', height(hr_clean));

% Check talent type distribution
talent_types = hr_clean{:, talent_col_name};
unique_types = unique(talent_types(~cellfun(@isempty, talent_types)));

fprintf('\nTalent type distribution:\n');
for i = 1:length(unique_types)
    count = sum(strcmp(talent_types, unique_types{i}));
    fprintf('%s: %d people\n', unique_types{i}, count);
end

%% 3. Parse Competency Data Structure
fprintf('\n=== Parsing Competency Data Structure ===\n');

% Header row (row 3) and data start point
headers = comp_raw(3, :);
data_start_row = 4;

% Find competency columns (starting from column 10)
competency_headers = {};
competency_cols = [];

for j = 10:min(50, size(comp_raw, 2))
    header_val = comp_raw{3, j};
    if ischar(header_val) && ~isempty(header_val)
        competency_headers{end+1} = header_val;
        competency_cols = [competency_cols, j];
    end
end

fprintf('Found competency items: %d\n', length(competency_headers));
fprintf('Main competency items:\n');
for i = 1:min(10, length(competency_headers))
    fprintf('  %d. %s\n', i, competency_headers{i});
end

%% 4. Extract Valid Competency Data
fprintf('\n=== Extracting Valid Competency Data ===\n');

% Find rows with valid ID and scores
valid_comp_data = [];
valid_ids = [];

for i = data_start_row:size(comp_raw, 1)
    id_val = comp_raw{i, 1};
    if isnumeric(id_val) && id_val > 1000000
        % Extract competency scores for this row
        scores = [];
        for j = 1:length(competency_cols)
            col_idx = competency_cols(j);
            val = comp_raw{i, col_idx};
            if isnumeric(val) && ~isnan(val) && val >= 1 && val <= 100
                scores = [scores, val];
            else
                scores = [scores, NaN];
            end
        end

        % Include if at least 5 valid scores exist
        if sum(~isnan(scores)) >= 5
            valid_ids = [valid_ids; id_val];
            valid_comp_data = [valid_comp_data; scores];
        end
    end
end

fprintf('Valid competency data: %d people\n', length(valid_ids));

%% 5. Match Talent Types with Competency Data
fprintf('\n=== Matching Data ===\n');

matched_data = [];
matched_talent_types = {};
matched_ids = [];

for i = 1:height(hr_clean)
    hr_id = hr_clean.ID(i);

    % Find corresponding ID in competency data
    comp_idx = find(valid_ids == hr_id);

    if ~isempty(comp_idx) && ~isempty(hr_clean{i, talent_col_name}{1})
        matched_ids = [matched_ids; hr_id];
        matched_talent_types{end+1} = hr_clean{i, talent_col_name}{1};
        matched_data = [matched_data; valid_comp_data(comp_idx, :)];
    end
end

fprintf('Matched data: %d people\n', length(matched_ids));

% Distribution of matched talent types
fprintf('\nMatched talent type distribution:\n');
unique_matched_types = unique(matched_talent_types);
for i = 1:length(unique_matched_types)
    count = sum(strcmp(matched_talent_types, unique_matched_types{i}));
    fprintf('%s: %d people\n', unique_matched_types{i}, count);
end

%% 6. Define Performance Ranking for Talent Types
fprintf('\n=== Defining Talent Type Performance Ranking ===\n');

% Performance ranking provided by user (highest to lowest)
performance_ranking = {
    '자연성', 8;
    '성실한 가연성', 7;
    '유익한 불연성', 6;
    '유능한 불연성', 5;
    '게으른 가연성', 4;
    '무능한 불연성', 3;
    '위장형 소화성', 2;
    '소화성', 1
};

% Create performance score mapping
performance_scores = containers.Map();
for i = 1:size(performance_ranking, 1)
    performance_scores(performance_ranking{i, 1}) = performance_ranking{i, 2};
end

% Assign performance scores to matched data
matched_performance = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    talent_type = matched_talent_types{i};
    if performance_scores.isKey(talent_type)
        matched_performance(i) = performance_scores(talent_type);
    else
        fprintf('Warning: Unknown talent type - %s\n', talent_type);
        matched_performance(i) = 0;
    end
end

fprintf('Data with performance scores assigned: %d people\n', sum(matched_performance > 0));

%% 7. Save and Prepare for Next Steps
fprintf('\n=== Saving Data ===\n');

% Save results in structure
analysis_data = struct();
analysis_data.matched_ids = matched_ids;
analysis_data.matched_talent_types = matched_talent_types;
analysis_data.matched_data = matched_data;
analysis_data.matched_performance = matched_performance;
analysis_data.competency_headers = competency_headers;
analysis_data.performance_ranking = performance_ranking;

% Save to file
save('talent_competency_merged_data.mat', 'analysis_data');
fprintf('Merged data saved: talent_competency_merged_data.mat\n');

fprintf('\n=== Stage 1 Complete ===\n');
fprintf('Next: Talent type competency analysis and weight calculation\n');