%% Performance Prediction Weight Calculation
% Stage 3: Calculate optimal weights for performance prediction

clear; clc;

%% 1. Load Analysis Results
fprintf('=== Loading Analysis Results ===\n');
load('talent_competency_merged_data.mat');
load('competency_statistical_results.mat');

matched_data = analysis_data.matched_data;
matched_performance = analysis_data.matched_performance;
valid_competencies = statistical_results.valid_competencies;
valid_data_idx = ~isnan(sum(matched_data, 2));  % Remove rows with all NaN

% Use only valid data
X = matched_data(valid_data_idx, 1:length(valid_competencies));
y = matched_performance(valid_data_idx);

% Remove competencies with too many missing values
missing_threshold = 0.5;  % 50% threshold
missing_ratio = sum(isnan(X), 1) ./ size(X, 1);
valid_comp_mask = missing_ratio < missing_threshold;

X_clean = X(:, valid_comp_mask);
clean_competencies = valid_competencies(valid_comp_mask);

fprintf('Clean dataset: %d people, %d competencies\n', size(X_clean, 1), size(X_clean, 2));

%% 2. Handle Missing Values
fprintf('\n=== Handling Missing Values ===\n');

% Replace missing values with mean imputation
for i = 1:size(X_clean, 2)
    col_mean = nanmean(X_clean(:, i));
    missing_idx = isnan(X_clean(:, i));
    X_clean(missing_idx, i) = col_mean;

    if sum(missing_idx) > 0
        fprintf('Imputed %d missing values in %s with mean %.2f\n', ...
                sum(missing_idx), clean_competencies{i}, col_mean);
    end
end

%% 3. Multiple Regression Analysis
fprintf('\n=== Multiple Regression Analysis ===\n');

% Standardize features
X_std = zscore(X_clean);
y_std = zscore(y);

% Multiple linear regression
[b, bint, r, rint, stats] = regress(y_std, [ones(size(X_std, 1), 1), X_std]);

R_squared = stats(1);
F_stat = stats(2);
p_value = stats(3);

fprintf('Multiple Regression Results:\n');
fprintf('  R-squared: %.4f\n', R_squared);
fprintf('  F-statistic: %.4f\n', F_stat);
fprintf('  p-value: %.4f\n', p_value);

% Extract coefficients (excluding intercept)
regression_weights = b(2:end);

%% 4. Ridge Regression for Regularization
fprintf('\n=== Ridge Regression Analysis ===\n');

% Cross-validation to find optimal lambda
lambda_values = logspace(-3, 2, 20);  % Lambda from 0.001 to 100
cv_errors = zeros(size(lambda_values));

% 5-fold cross-validation
k_folds = 5;
n = size(X_std, 1);
% Manual k-fold split
fold_size = floor(n / k_folds);
indices = zeros(n, 1);
for i = 1:k_folds-1
    start_idx = (i-1) * fold_size + 1;
    end_idx = i * fold_size;
    indices(start_idx:end_idx) = i;
end
indices((k_folds-1)*fold_size+1:end) = k_folds;

for i = 1:length(lambda_values)
    lambda = lambda_values(i);
    fold_errors = zeros(k_folds, 1);

    for fold = 1:k_folds
        train_idx = indices ~= fold;
        test_idx = indices == fold;

        X_train = X_std(train_idx, :);
        y_train = y_std(train_idx);
        X_test = X_std(test_idx, :);
        y_test = y_std(test_idx);

        % Ridge regression: beta = (X'X + lambda*I)^(-1) X'y
        beta = (X_train' * X_train + lambda * eye(size(X_train, 2))) \ (X_train' * y_train);

        % Prediction
        y_pred = X_test * beta;
        fold_errors(fold) = mean((y_test - y_pred).^2);
    end

    cv_errors(i) = mean(fold_errors);
end

% Find optimal lambda
[min_error, best_idx] = min(cv_errors);
optimal_lambda = lambda_values(best_idx);

fprintf('Optimal lambda: %.4f (CV error: %.4f)\n', optimal_lambda, min_error);

% Train final ridge model with optimal lambda
ridge_weights = (X_std' * X_std + optimal_lambda * eye(size(X_std, 2))) \ (X_std' * y_std);

%% 5. Correlation-Based Weights
fprintf('\n=== Correlation-Based Weights ===\n');

correlation_weights = zeros(size(X_clean, 2), 1);
for i = 1:size(X_clean, 2)
    % Manual correlation calculation
    x_col = X_clean(:, i);
    valid_idx = ~isnan(x_col) & ~isnan(y);
    if sum(valid_idx) > 2
        x_valid = x_col(valid_idx);
        y_valid = y(valid_idx);
        correlation_weights(i) = sum((x_valid - mean(x_valid)) .* (y_valid - mean(y_valid))) / ...
                               (sqrt(sum((x_valid - mean(x_valid)).^2)) * sqrt(sum((y_valid - mean(y_valid)).^2)));
    else
        correlation_weights(i) = 0;
    end
end

% Normalize to sum to 1 (for positive correlations only)
positive_corr_weights = max(correlation_weights, 0);
if sum(positive_corr_weights) > 0
    positive_corr_weights = positive_corr_weights / sum(positive_corr_weights);
end

%% 6. Combined Weight Calculation
fprintf('\n=== Combined Weight Calculation ===\n');

% Normalize each weight method
regression_weights_norm = abs(regression_weights) / sum(abs(regression_weights));
ridge_weights_norm = abs(ridge_weights) / sum(abs(ridge_weights));
correlation_weights_norm = abs(correlation_weights) / sum(abs(correlation_weights));

% Combined weights (equal contribution from each method)
combined_weights = (regression_weights_norm + ridge_weights_norm + correlation_weights_norm) / 3;

%% 7. Rank and Display Results
fprintf('\n=== Final Weight Rankings ===\n');

% Sort by combined weights
[sorted_weights, sort_idx] = sort(combined_weights, 'descend');

fprintf('Top 15 Competencies for Performance Prediction:\n');
fprintf('Rank | Competency                | Combined | Regression | Ridge   | Correlation\n');
fprintf('-----|---------------------------|----------|------------|---------|------------\n');

for i = 1:min(15, length(sort_idx))
    idx = sort_idx(i);
    fprintf('%4d | %-25s | %8.4f | %10.4f | %7.4f | %11.4f\n', ...
            i, clean_competencies{idx}, ...
            combined_weights(idx), ...
            regression_weights_norm(idx), ...
            ridge_weights_norm(idx), ...
            correlation_weights_norm(idx));
end

%% 8. Validation
fprintf('\n=== Model Validation ===\n');

% Calculate prediction accuracy using combined weights
X_weighted = X_std * combined_weights;
predicted_performance = X_weighted;

% Correlation between predicted and actual performance (manual calculation)
valid_pred_idx = ~isnan(predicted_performance) & ~isnan(y_std);
if sum(valid_pred_idx) > 2
    pred_valid = predicted_performance(valid_pred_idx);
    y_valid = y_std(valid_pred_idx);
    prediction_correlation = sum((pred_valid - mean(pred_valid)) .* (y_valid - mean(y_valid))) / ...
                            (sqrt(sum((pred_valid - mean(pred_valid)).^2)) * sqrt(sum((y_valid - mean(y_valid)).^2)));
else
    prediction_correlation = 0;
end
prediction_rmse = sqrt(mean((predicted_performance - y_std).^2));

fprintf('Validation Results:\n');
fprintf('  Prediction correlation: %.4f\n', prediction_correlation);
fprintf('  RMSE: %.4f\n', prediction_rmse);

%% 9. Create Final Weight Structure
fprintf('\n=== Creating Final Weight Structure ===\n');

% Group by higher-level categories if possible
% Map individual competencies to broader categories
competency_categories = {
    '도전', '전문성';
    '관찰', '관찰력';
    '분석', '분석력';
    '소통', '소통력';
    '창의', '창의성';
    '학습', '학습력';
    '활용', '활용력'
};

% Create category-level weights
category_weights = containers.Map();
category_counts = containers.Map();

for i = 1:length(clean_competencies)
    comp_name = clean_competencies{i};
    weight = combined_weights(i);

    % Find matching category
    category = '기타';  % Default category
    for j = 1:size(competency_categories, 1)
        if contains(comp_name, competency_categories{j, 1})
            category = competency_categories{j, 2};
            break;
        end
    end

    % Accumulate weights by category
    if category_weights.isKey(category)
        category_weights(category) = category_weights(category) + weight;
        category_counts(category) = category_counts(category) + 1;
    else
        category_weights(category) = weight;
        category_counts(category) = 1;
    end
end

fprintf('\nCategory-level weights:\n');
categories = keys(category_weights);
for i = 1:length(categories)
    category = categories{i};
    total_weight = category_weights(category);
    count = category_counts(category);
    avg_weight = total_weight / count;

    fprintf('  %s: %.4f (total), %.4f (average from %d items)\n', ...
            category, total_weight, avg_weight, count);
end

%% 10. Save Final Results
fprintf('\n=== Saving Final Results ===\n');

final_results = struct();
final_results.clean_competencies = clean_competencies;
final_results.combined_weights = combined_weights;
final_results.regression_weights = regression_weights_norm;
final_results.ridge_weights = ridge_weights_norm;
final_results.correlation_weights = correlation_weights_norm;
final_results.weight_rankings = sort_idx;
final_results.prediction_correlation = prediction_correlation;
final_results.prediction_rmse = prediction_rmse;
final_results.optimal_lambda = optimal_lambda;
final_results.R_squared = R_squared;

% Save to file
save('performance_prediction_weights.mat', 'final_results');

% Also save as Excel file for easy viewing
results_table = table();
results_table.Rank = (1:length(sort_idx))';
results_table.Competency = clean_competencies(sort_idx);
results_table.Combined_Weight = combined_weights(sort_idx);
results_table.Regression_Weight = regression_weights_norm(sort_idx);
results_table.Ridge_Weight = ridge_weights_norm(sort_idx);
results_table.Correlation_Weight = correlation_weights_norm(sort_idx);

writetable(results_table, 'competency_weights_ranking.xlsx');

fprintf('Results saved:\n');
fprintf('  - MATLAB file: performance_prediction_weights.mat\n');
fprintf('  - Excel file: competency_weights_ranking.xlsx\n');

fprintf('\n=== Analysis Complete ===\n');
fprintf('성과 예측을 위한 역량 가중치 산출이 완료되었습니다.\n');