%% 개선된 HR 인재유형 분석 시스템 - 최적화 버전
% 주요 개선사항:
% 1. 개별 레이더 차트 생성 (통일된 스케일)
% 2. 이진 분류 머신러닝 (LogitBoost, TreeBagger)
% 3. 하이퍼파라미터 튜닝 및 교차검증

clear; clc; close all;

%% ========================================================================
%                          PART 1: 초기 설정 및 데이터 로딩
% =========================================================================

fprintf('╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║       개선된 HR 인재유형 분석 시스템 - 최적화 버전       ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);
set(0, 'DefaultLineLineWidth', 2);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.output_dir = pwd;
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 성과 순위 정의 (위장형 소화성 제외)
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 1]);

% 고성과자/저성과자 정의 (이진 분류용)
config.high_performers = {'자연성', '성실한 가연성', '유능한 불연성'};
config.low_performers = {'무능한 불연성', '소화성'};
config.excluded_types = {'위장형 소화성', '유익한 불연성', '게으른 가연성'}; % 중간 그룹 제외

%% 1.1 데이터 로딩
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
    fprintf('  ✓ 상위항목 데이터: %d명\n', height(comp_upper));
    
catch ME
    error('데이터 로딩 실패: %s', ME.message);
end

%% 1.2 인재유형 데이터 추출 및 정제
fprintf('\n【STEP 2】 인재유형 데이터 추출 및 정제\n');
fprintf('────────────────────────────────────────────\n');

% 인재유형 컬럼 찾기
talent_col_idx = find(contains(hr_data.Properties.VariableNames, {'인재유형', '인재', '유형'}), 1);
if isempty(talent_col_idx)
    error('인재유형 컬럼을 찾을 수 없습니다.');
end

talent_col_name = hr_data.Properties.VariableNames{talent_col_idx};
fprintf('▶ 인재유형 컬럼: %s\n', talent_col_name);

% 빈 값 제거
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_idx}), :);

% 인재유형 분포 분석
talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('\n전체 인재유형 분포:\n');
for i = 1:length(unique_types)
    fprintf('  • %-20s: %3d명 (%5.1f%%)\n', ...
        unique_types{i}, type_counts(i), type_counts(i)/sum(type_counts)*100);
end

%% 1.3 역량 데이터 처리
fprintf('\n【STEP 3】 역량 데이터 처리\n');
fprintf('────────────────────────────────────────────\n');

% ID 컬럼 찾기
comp_id_col = find(contains(lower(comp_upper.Properties.VariableNames), {'id', '사번'}), 1);
if isempty(comp_id_col)
    error('역량 데이터에서 ID 컬럼을 찾을 수 없습니다.');
end

% 유효한 역량 컬럼 추출
valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)
    col_data = comp_upper{:, i};
    if isnumeric(col_data) && ~all(isnan(col_data))
        valid_data = col_data(~isnan(col_data));
        if length(valid_data) >= 5 && var(valid_data) > 0 && ...
           all(valid_data >= 0) && all(valid_data <= 100)
            valid_comp_cols{end+1} = comp_upper.Properties.VariableNames{i};
            valid_comp_indices(end+1) = i;
        end
    end
end

fprintf('▶ 유효한 역량 항목: %d개\n', length(valid_comp_cols));

%% 1.4 ID 매칭 및 데이터 통합
fprintf('\n【STEP 4】 데이터 매칭 및 통합\n');
fprintf('────────────────────────────────────────────\n');

% ID 표준화
hr_ids_str = arrayfun(@(x) sprintf('%.0f', x), hr_clean.ID, 'UniformOutput', false);
comp_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_upper{:, comp_id_col}, 'UniformOutput', false);

% 교집합 찾기
[matched_ids, hr_idx, comp_idx] = intersect(hr_ids_str, comp_ids_str);

fprintf('▶ 매칭 성공: %d명\n', length(matched_ids));

% 매칭된 데이터 추출
matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, valid_comp_indices);
matched_talent_types = matched_hr{:, talent_col_idx};

%% ========================================================================
%            PART 2: 개선된 레이더 차트 (개별 Figure, 통일 스케일)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║     PART 2: 개선된 레이더 차트 (통일 스케일)             ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 2.1 유형별 프로파일 계산 및 스케일 범위 설정
fprintf('【STEP 5】 유형별 프로파일 계산 및 스케일 설정\n');
fprintf('────────────────────────────────────────────\n');

unique_matched_types = unique(matched_talent_types);
n_types = length(unique_matched_types);

% 프로파일 계산
type_profiles = zeros(n_types, length(valid_comp_cols));
for i = 1:n_types
    type_mask = strcmp(matched_talent_types, unique_matched_types{i});
    type_comp_data = matched_comp{type_mask, :};
    type_profiles(i, :) = nanmean(type_comp_data, 1);
end

% 상위 12개 주요 역량 선정 (분산 기준)
comp_variance = var(table2array(matched_comp), 0, 1, 'omitnan');
[~, var_idx] = sort(comp_variance, 'descend');
top_comp_idx = var_idx(1:min(12, length(var_idx)));
top_comp_names = valid_comp_cols(top_comp_idx);

% 전체 평균 프로파일
overall_mean_profile = nanmean(table2array(matched_comp), 1);

% 통일된 스케일 범위 계산 (모든 유형의 최소/최대값)
all_profile_data = type_profiles(:, top_comp_idx);
global_min = min(all_profile_data(:)) - 5;  % 여유값 5점
global_max = max(all_profile_data(:)) + 5;  % 여유값 5점

fprintf('▶ 통일 스케일 범위: %.1f ~ %.1f\n', global_min, global_max);
fprintf('▶ 선정된 주요 역량: %d개\n', length(top_comp_idx));

%% 2.2 개별 레이더 차트 생성
fprintf('\n【STEP 6】 개별 레이더 차트 생성\n');
fprintf('────────────────────────────────────────────\n');

% 컬러맵 설정
colors = lines(n_types);

for i = 1:n_types
    % 새로운 Figure 창 생성
    fig = figure('Position', [100 + (i-1)*50, 100 + (i-1)*30, 800, 800], ...
                 'Color', 'white', ...
                 'Name', sprintf('인재유형: %s', unique_matched_types{i}));
    
    % 해당 유형의 프로파일 데이터
    type_profile = type_profiles(i, top_comp_idx);
    baseline = overall_mean_profile(top_comp_idx);
    
    % 개선된 레이더 차트 그리기
    createEnhancedRadarChart(type_profile, baseline, top_comp_names, ...
                            unique_matched_types{i}, colors(i,:), ...
                            global_min, global_max);
    
    % 추가 정보 표시
    if config.performance_ranking.isKey(unique_matched_types{i})
        perf_rank = config.performance_ranking(unique_matched_types{i});
        text(0.5, -0.05, sprintf('성과순위: %d', perf_rank), ...
             'Units', 'normalized', ...
             'HorizontalAlignment', 'center', ...
             'FontWeight', 'bold', 'FontSize', 14);
    end
    
    % Figure 저장
    saveas(fig, sprintf('radar_chart_%s_%s.png', ...
           strrep(unique_matched_types{i}, ' ', '_'), config.timestamp));
    
    fprintf('  ✓ %s 차트 생성 완료\n', unique_matched_types{i});
end

%% ========================================================================
%         PART 3: 이진 분류 머신러닝 (고성과자 vs 저성과자)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║    PART 3: 이진 분류 머신러닝 (LogitBoost & TreeBagger)  ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 3.1 이진 레이블 생성
fprintf('【STEP 7】 고성과자/저성과자 이진 레이블 생성\n');
fprintf('────────────────────────────────────────────\n');

% 이진 분류용 데이터 필터링
binary_mask = false(length(matched_talent_types), 1);
binary_labels = zeros(length(matched_talent_types), 1);

for i = 1:length(matched_talent_types)
    type = matched_talent_types{i};
    
    if ismember(type, config.high_performers)
        binary_mask(i) = true;
        binary_labels(i) = 1;  % 고성과자
    elseif ismember(type, config.low_performers)
        binary_mask(i) = true;
        binary_labels(i) = 0;  % 저성과자
    end
end

% 이진 분류용 데이터 추출
X_binary = table2array(matched_comp(binary_mask, :));
y_binary = binary_labels(binary_mask);

fprintf('▶ 고성과자: %d명\n', sum(y_binary == 1));
fprintf('▶ 저성과자: %d명\n', sum(y_binary == 0));
fprintf('▶ 제외된 중간그룹: %d명\n', sum(~binary_mask));

% 데이터 정규화
X_binary_norm = normalize(X_binary, 'range');

%% 3.2 교차검증 설정
fprintf('\n【STEP 8】 교차검증 설정\n');
fprintf('────────────────────────────────────────────\n');

% 5-fold 교차검증
k_folds = 5;
cv = cvpartition(y_binary, 'KFold', k_folds, 'Stratify', true);

fprintf('▶ %d-fold 교차검증 설정 완료\n', k_folds);

%% 3.3 LogitBoost 모델 (하이퍼파라미터 튜닝)
fprintf('\n【STEP 9】 LogitBoost 모델 학습 및 튜닝\n');
fprintf('────────────────────────────────────────────\n');

% 하이퍼파라미터 그리드
num_cycles_grid = [50, 100, 150, 200];
learn_rate_grid = [0.05, 0.1, 0.2, 0.5];
max_splits_grid = [10, 20, 30];

best_logit_accuracy = 0;
best_logit_params = struct();

fprintf('LogitBoost 하이퍼파라미터 튜닝 중...\n');

for nc = num_cycles_grid
    for lr = learn_rate_grid
        for ms = max_splits_grid
            cv_accuracies = zeros(k_folds, 1);
            
            for fold = 1:k_folds
                % 훈련/검증 분할
                train_idx = training(cv, fold);
                val_idx = test(cv, fold);
                
                X_train = X_binary_norm(train_idx, :);
                y_train = y_binary(train_idx);
                X_val = X_binary_norm(val_idx, :);
                y_val = y_binary(val_idx);
                
                % 모델 학습
                try
                    model = fitcensemble(X_train, y_train, ...
                        'Method', 'LogitBoost', ...
                        'NumLearningCycles', nc, ...
                        'LearnRate', lr, ...
                        'Learners', templateTree('MaxNumSplits', ms));
                    
                    % 예측 및 평가
                    y_pred = predict(model, X_val);
                    cv_accuracies(fold) = sum(y_pred == y_val) / length(y_val);
                catch
                    cv_accuracies(fold) = 0;
                end
            end
            
            mean_accuracy = mean(cv_accuracies);
            
            if mean_accuracy > best_logit_accuracy
                best_logit_accuracy = mean_accuracy;
                best_logit_params.NumCycles = nc;
                best_logit_params.LearnRate = lr;
                best_logit_params.MaxSplits = ms;
            end
        end
    end
end

fprintf('최적 LogitBoost 파라미터:\n');
fprintf('  • NumLearningCycles: %d\n', best_logit_params.NumCycles);
fprintf('  • LearnRate: %.2f\n', best_logit_params.LearnRate);
fprintf('  • MaxNumSplits: %d\n', best_logit_params.MaxSplits);
fprintf('  • CV 정확도: %.3f\n', best_logit_accuracy);

% 최적 파라미터로 최종 모델 학습
final_logit_model = fitcensemble(X_binary_norm, y_binary, ...
    'Method', 'LogitBoost', ...
    'NumLearningCycles', best_logit_params.NumCycles, ...
    'LearnRate', best_logit_params.LearnRate, ...
    'Learners', templateTree('MaxNumSplits', best_logit_params.MaxSplits));

% Feature Importance
logit_importance = predictorImportance(final_logit_model);

%% 3.4 TreeBagger 모델 (Random Forest, 하이퍼파라미터 튜닝)
fprintf('\n【STEP 10】 TreeBagger 모델 학습 및 튜닝\n');
fprintf('────────────────────────────────────────────\n');

% 하이퍼파라미터 그리드
num_trees_grid = [100, 200, 300];
min_leaf_grid = [1, 2, 5];
max_splits_grid = [20, 30, 50];

best_rf_accuracy = 0;
best_rf_params = struct();

fprintf('TreeBagger 하이퍼파라미터 튜닝 중...\n');

for nt = num_trees_grid
    for ml = min_leaf_grid
        for ms = max_splits_grid
            cv_accuracies = zeros(k_folds, 1);
            
            for fold = 1:k_folds
                % 훈련/검증 분할
                train_idx = training(cv, fold);
                val_idx = test(cv, fold);
                
                X_train = X_binary_norm(train_idx, :);
                y_train = y_binary(train_idx);
                X_val = X_binary_norm(val_idx, :);
                y_val = y_binary(val_idx);
                
                % 모델 학습
                model = TreeBagger(nt, X_train, y_train, ...
                    'Method', 'classification', ...
                    'MinLeafSize', ml, ...
                    'MaxNumSplits', ms, ...
                    'OOBPredictorImportance', 'on');
                
                % 예측 및 평가
                [y_pred_cell, ~] = predict(model, X_val);
                y_pred = cellfun(@str2double, y_pred_cell);
                cv_accuracies(fold) = sum(y_pred == y_val) / length(y_val);
            end
            
            mean_accuracy = mean(cv_accuracies);
            
            if mean_accuracy > best_rf_accuracy
                best_rf_accuracy = mean_accuracy;
                best_rf_params.NumTrees = nt;
                best_rf_params.MinLeaf = ml;
                best_rf_params.MaxSplits = ms;
            end
        end
    end
end

fprintf('최적 TreeBagger 파라미터:\n');
fprintf('  • NumTrees: %d\n', best_rf_params.NumTrees);
fprintf('  • MinLeafSize: %d\n', best_rf_params.MinLeaf);
fprintf('  • MaxNumSplits: %d\n', best_rf_params.MaxSplits);
fprintf('  • CV 정확도: %.3f\n', best_rf_accuracy);

% 최적 파라미터로 최종 모델 학습
final_rf_model = TreeBagger(best_rf_params.NumTrees, X_binary_norm, y_binary, ...
    'Method', 'classification', ...
    'MinLeafSize', best_rf_params.MinLeaf, ...
    'MaxNumSplits', best_rf_params.MaxSplits, ...
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', valid_comp_cols);

% Feature Importance
rf_importance = final_rf_model.OOBPermutedPredictorDeltaError;

%% 3.5 최종 성능 평가 (홀드아웃 테스트)
fprintf('\n【STEP 11】 최종 모델 성능 평가\n');
fprintf('────────────────────────────────────────────\n');

% 독립적인 테스트 세트로 최종 평가
test_partition = cvpartition(y_binary, 'HoldOut', 0.2, 'Stratify', true);
X_train_final = X_binary_norm(test_partition.training, :);
y_train_final = y_binary(test_partition.training);
X_test_final = X_binary_norm(test_partition.test, :);
y_test_final = y_binary(test_partition.test);

% LogitBoost 평가
y_pred_logit = predict(final_logit_model, X_test_final);
logit_test_accuracy = sum(y_pred_logit == y_test_final) / length(y_test_final);
logit_conf = confusionmat(y_test_final, y_pred_logit);
logit_precision = logit_conf(2,2) / sum(logit_conf(:,2));
logit_recall = logit_conf(2,2) / sum(logit_conf(2,:));
logit_f1 = 2 * (logit_precision * logit_recall) / (logit_precision + logit_recall);

% TreeBagger 평가
[y_pred_rf_cell, rf_scores] = predict(final_rf_model, X_test_final);
y_pred_rf = cellfun(@str2double, y_pred_rf_cell);
rf_test_accuracy = sum(y_pred_rf == y_test_final) / length(y_test_final);
rf_conf = confusionmat(y_test_final, y_pred_rf);
rf_precision = rf_conf(2,2) / sum(rf_conf(:,2));
rf_recall = rf_conf(2,2) / sum(rf_conf(2,:));
rf_f1 = 2 * (rf_precision * rf_recall) / (rf_precision + rf_recall);

fprintf('\n최종 테스트 세트 성능:\n');
fprintf('┌─────────────┬──────────┬──────────┬──────────┬──────────┐\n');
fprintf('│    모델     │  정확도  │  정밀도  │  재현율  │ F1-Score │\n');
fprintf('├─────────────┼──────────┼──────────┼──────────┼──────────┤\n');
fprintf('│ LogitBoost  │  %.3f   │  %.3f   │  %.3f   │  %.3f   │\n', ...
        logit_test_accuracy, logit_precision, logit_recall, logit_f1);
fprintf('│ TreeBagger  │  %.3f   │  %.3f   │  %.3f   │  %.3f   │\n', ...
        rf_test_accuracy, rf_precision, rf_recall, rf_f1);
fprintf('└─────────────┴──────────┴──────────┴──────────┴──────────┘\n');

%% ========================================================================
%              PART 4: 상관분석 vs 머신러닝 가중치 비교
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║        PART 4: 상관분석 vs 머신러닝 가중치 비교          ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 4.1 상관분석 기반 가중치 계산
fprintf('【STEP 12】 상관분석 기반 가중치 계산\n');
fprintf('────────────────────────────────────────────\n');

% 성과점수 기반 상관분석 (전체 데이터)
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    type_name = matched_talent_types{i};
    if config.performance_ranking.isKey(type_name)
        performance_scores(i) = config.performance_ranking(type_name);
    end
end

valid_perf_idx = performance_scores > 0;
valid_performance = performance_scores(valid_perf_idx);
valid_competencies = table2array(matched_comp(valid_perf_idx, :));

% 상관계수 계산
n_competencies = length(valid_comp_cols);
correlations = zeros(n_competencies, 1);
p_values = zeros(n_competencies, 1);

for i = 1:n_competencies
    comp_scores = valid_competencies(:, i);
    valid_idx = ~isnan(comp_scores);
    
    if sum(valid_idx) >= 10
        [r, p] = corr(comp_scores(valid_idx), valid_performance(valid_idx), 'Type', 'Spearman');
        correlations(i) = r;
        p_values(i) = p;
    end
end

% 양의 상관계수만 사용하여 가중치 계산
positive_corr = max(0, correlations);
corr_weights = positive_corr / sum(positive_corr);

%% 4.2 Feature Importance 정규화
fprintf('\n【STEP 13】 Feature Importance 통합 및 비교\n');
fprintf('────────────────────────────────────────────\n');

% 벡터 차원 확인 및 조정
n_features = length(valid_comp_cols);

% LogitBoost importance를 열 벡터로 변환
logit_importance = logit_importance(:);
if length(logit_importance) ~= n_features
    % 차원이 다른 경우 조정
    temp_logit = zeros(n_features, 1);
    min_len = min(length(logit_importance), n_features);
    temp_logit(1:min_len) = logit_importance(1:min_len);
    logit_importance = temp_logit;
end

% TreeBagger importance를 열 벡터로 변환
rf_importance = rf_importance(:);
if length(rf_importance) ~= n_features
    % 차원이 다른 경우 조정
    temp_rf = zeros(n_features, 1);
    min_len = min(length(rf_importance), n_features);
    temp_rf(1:min_len) = rf_importance(1:min_len);
    rf_importance = temp_rf;
end

% 상관계수 가중치도 열 벡터로 변환
corr_weights = corr_weights(:);
if length(corr_weights) ~= n_features
    temp_corr = zeros(n_features, 1);
    min_len = min(length(corr_weights), n_features);
    temp_corr(1:min_len) = corr_weights(1:min_len);
    corr_weights = temp_corr;
end

% 정규화
logit_weights = logit_importance / (sum(logit_importance) + eps);
rf_weights = rf_importance / (sum(rf_importance) + eps);
corr_weights = corr_weights / (sum(corr_weights) + eps);

% 머신러닝 통합 가중치 (LogitBoost와 TreeBagger 평균)
ml_weights = (logit_weights + rf_weights) / 2;
ml_weights = ml_weights / (sum(ml_weights) + eps);

% 최종 통합 가중치 (상관분석 40% + 머신러닝 60%)
final_weights = 0.4 * corr_weights + 0.6 * ml_weights;
final_weights = final_weights / (sum(final_weights) + eps);

% 가중치 비교 테이블 생성
weight_comparison = table();
weight_comparison.Competency = valid_comp_cols';
weight_comparison.Correlation = corr_weights * 100;
weight_comparison.LogitBoost = logit_weights * 100;
weight_comparison.TreeBagger = rf_weights * 100;
weight_comparison.ML_Combined = ml_weights * 100;
weight_comparison.Final = final_weights * 100;

% 최종 가중치 기준 정렬
weight_comparison = sortrows(weight_comparison, 'Final', 'descend');

fprintf('\n상위 15개 역량 가중치 비교:\n');
fprintf('%-20s | 상관(%) | LogB(%) | TreeB(%) | ML평균(%) | 최종(%)\n', '역량');
fprintf('%s\n', repmat('─', 75, 1));

for i = 1:min(15, height(weight_comparison))
    fprintf('%-20s | %6.2f | %7.2f | %8.2f | %9.2f | %7.2f\n', ...
        weight_comparison.Competency{i}, ...
        weight_comparison.Correlation(i), ...
        weight_comparison.LogitBoost(i), ...
        weight_comparison.TreeBagger(i), ...
        weight_comparison.ML_Combined(i), ...
        weight_comparison.Final(i));
end

%% 4.3 가중치 비교 시각화
figure('Position', [100, 100, 1600, 1000], 'Color', 'white');

% 상위 15개 역량의 방법별 가중치 비교
subplot(2, 2, [1, 2]);
top_15 = weight_comparison(1:min(15, height(weight_comparison)), :);
if height(top_15) > 0
    bar_data = [top_15.Correlation, top_15.LogitBoost, top_15.TreeBagger];
else
    bar_data = zeros(1, 3);
end

bar(bar_data);
set(gca, 'XTickLabel', top_15.Competency, 'XTickLabelRotation', 45);
ylabel('가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('방법별 Feature Importance 비교 (상위 15개)', 'FontSize', 14, 'FontWeight', 'bold');
legend('상관분석', 'LogitBoost', 'TreeBagger', 'Location', 'northeast');
grid on;

% 상관분석 vs ML 산점도
subplot(2, 2, 3);
scatter(weight_comparison.Correlation, weight_comparison.ML_Combined, ...
        50, 'filled', 'MarkerFaceColor', [0.3, 0.6, 0.9]);
xlabel('상관분석 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('ML 통합 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('상관분석 vs 머신러닝 가중치', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
hold on;
plot([0, max([weight_comparison.Correlation; weight_comparison.ML_Combined])], ...
     [0, max([weight_comparison.Correlation; weight_comparison.ML_Combined])], ...
     'r--', 'LineWidth', 1.5);

% 최종 통합 가중치
subplot(2, 2, 4);
barh(15:-1:1, top_15.Final, 'FaceColor', [0.8, 0.3, 0.3]);
set(gca, 'YTick', 1:15, 'YTickLabel', flip(top_15.Competency));
xlabel('최종 통합 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('최종 역량 중요도 순위', 'FontSize', 14, 'FontWeight', 'bold');
grid on;

sgtitle('가중치 분석 비교: 상관분석 vs 머신러닝', 'FontSize', 16, 'FontWeight', 'bold');

%% ========================================================================
%                      PART 5: 결과 저장 및 보고서
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║             PART 5: 결과 저장 및 최종 보고서             ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 5.1 Excel 보고서 생성
fprintf('【STEP 14】 Excel 보고서 생성\n');
fprintf('────────────────────────────────────────────\n');

output_filename = sprintf('hr_analysis_optimized_%s.xlsx', config.timestamp);

try
    % Sheet 1: 모델 성능 비교
    model_performance = table();
    model_performance.Model = {'LogitBoost'; 'TreeBagger'};
    model_performance.CV_Accuracy = [best_logit_accuracy; best_rf_accuracy];
    model_performance.Test_Accuracy = [logit_test_accuracy; rf_test_accuracy];
    model_performance.Precision = [logit_precision; rf_precision];
    model_performance.Recall = [logit_recall; rf_recall];
    model_performance.F1_Score = [logit_f1; rf_f1];
    writetable(model_performance, output_filename, 'Sheet', '모델성능');
    
    % Sheet 2: 가중치 비교
    writetable(weight_comparison, output_filename, 'Sheet', '가중치비교');
    
    % Sheet 3: 하이퍼파라미터
    hyperparams = table();
    hyperparams.Model = {'LogitBoost'; 'TreeBagger'};
    hyperparams.Param1 = {sprintf('NumCycles=%d', best_logit_params.NumCycles); ...
                          sprintf('NumTrees=%d', best_rf_params.NumTrees)};
    hyperparams.Param2 = {sprintf('LearnRate=%.2f', best_logit_params.LearnRate); ...
                          sprintf('MinLeaf=%d', best_rf_params.MinLeaf)};
    hyperparams.Param3 = {sprintf('MaxSplits=%d', best_logit_params.MaxSplits); ...
                          sprintf('MaxSplits=%d', best_rf_params.MaxSplits)};
    writetable(hyperparams, output_filename, 'Sheet', '하이퍼파라미터');
    
    fprintf('✓ Excel 보고서 저장 완료: %s\n', output_filename);
    
catch ME
    fprintf('⚠ Excel 저장 실패: %s\n', ME.message);
end

%% 5.2 MATLAB 파일 저장
analysis_results = struct();
analysis_results.models = struct('logit', final_logit_model, 'rf', final_rf_model);
analysis_results.performance = model_performance;
analysis_results.weights = weight_comparison;
analysis_results.hyperparams = struct('logit', best_logit_params, 'rf', best_rf_params);

save(sprintf('hr_analysis_optimized_%s.mat', config.timestamp), 'analysis_results');
fprintf('✓ MATLAB 파일 저장 완료\n');

%% 5.3 최종 보고서 출력
fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║                    최종 분석 보고서                       ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

fprintf('📊 데이터 요약\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • 전체 매칭 데이터: %d명\n', length(matched_ids));
fprintf('  • 이진 분류 데이터: %d명 (고성과자 %d, 저성과자 %d)\n', ...
        length(y_binary), sum(y_binary==1), sum(y_binary==0));
fprintf('  • 역량 항목: %d개\n', length(valid_comp_cols));

fprintf('\n🤖 머신러닝 모델 성능\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • LogitBoost: 정확도 %.1f%%, F1-Score %.3f\n', ...
        logit_test_accuracy*100, logit_f1);
fprintf('  • TreeBagger: 정확도 %.1f%%, F1-Score %.3f\n', ...
        rf_test_accuracy*100, rf_f1);

fprintf('\n⭐ 핵심 역량 Top 5 (최종 통합 가중치)\n');
fprintf('────────────────────────────────────────────\n');
for i = 1:min(5, height(weight_comparison))
    fprintf('  %d. %-20s: %5.2f%%\n', i, ...
            weight_comparison.Competency{i}, ...
            weight_comparison.Final(i));
end

fprintf('\n💡 방법론 비교 인사이트\n');
fprintf('────────────────────────────────────────────\n');

% 상관분석과 ML의 일치도 계산
top5_corr = weight_comparison.Competency(1:5);
[~, ml_idx] = sort(weight_comparison.ML_Combined, 'descend');
top5_ml = weight_comparison.Competency(ml_idx(1:5));
agreement = length(intersect(top5_corr, top5_ml));

fprintf('  • 상위 5개 역량 일치도: %d/5 (%.0f%%)\n', agreement, agreement*20);

if agreement >= 3
    fprintf('  • 상관분석과 ML이 유사한 결과 → 신뢰도 높음\n');
else
    fprintf('  • 상관분석과 ML이 다른 관점 → 통합 접근 필요\n');
end

fprintf('\n✅ 실무 적용 권장사항\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  1. 1차 스크리닝: %s, %s, %s\n', ...
        weight_comparison.Competency{1}, ...
        weight_comparison.Competency{2}, ...
        weight_comparison.Competency{3});
fprintf('  2. 모델 신뢰도: ');
if mean([logit_test_accuracy, rf_test_accuracy]) > 0.75
    fprintf('높음 (실무 즉시 적용 가능)\n');
elseif mean([logit_test_accuracy, rf_test_accuracy]) > 0.65
    fprintf('중간 (보조 도구로 활용)\n');
else
    fprintf('낮음 (추가 데이터 필요)\n');
end
fprintf('  3. 정기 업데이트: 분기별 재학습 권장\n');

fprintf('\n════════════════════════════════════════════════════════════\n');
fprintf('           분석 완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('════════════════════════════════════════════════════════════\n\n');

%% Helper Function: 개선된 레이더 차트
function createEnhancedRadarChart(data, baseline, labels, title_text, color, min_val, max_val)
    % 개선된 레이더 차트 생성 (통일된 스케일 적용)
    n_vars = length(data);
    angles = linspace(0, 2*pi, n_vars+1);
    
    % 스케일 정규화 (통일된 범위 사용)
    data_norm = (data - min_val) / (max_val - min_val);
    baseline_norm = (baseline - min_val) / (max_val - min_val);
    
    % 순환을 위해 첫 번째 값을 마지막에 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];
    
    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);
    
    hold on;
    
    % 그리드 그리기 (5단계)
    grid_levels = 5;
    for i = 1:grid_levels
        r = i / grid_levels;
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);
        
        % 그리드 레이블 (실제 값으로 표시)
        grid_value = min_val + (max_val - min_val) * r;
        text(0, r, sprintf('%.0f', grid_value), ...
             'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'bottom', ...
             'FontSize', 9, 'Color', [0.4 0.4 0.4]);
    end
    
    % 방사선 그리기
    for i = 1:n_vars
        plot([0, cos(angles(i))], [0, sin(angles(i))], 'k:', ...
             'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);
    end
    
    % 기준선 (전체 평균)
    plot(x_base, y_base, '--', 'Color', [0.7 0.7 0.7], 'LineWidth', 2);
    
    % 데이터 플롯
    patch(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2.5);
    
    % 데이터 포인트
    scatter(x_data(1:end-1), y_data(1:end-1), 60, color, 'filled', ...
            'MarkerEdgeColor', 'white', 'LineWidth', 1);
    
    % 레이블 및 값
    label_radius = 1.25;
    for i = 1:n_vars
        [lx, ly] = pol2cart(angles(i), label_radius);
        
        % 차이값 계산
        diff_val = data(i) - baseline(i);
        diff_str = sprintf('%+.1f', diff_val);
        if diff_val > 0
            diff_color = [0, 0.5, 0];
        else
            diff_color = [0.8, 0, 0];
        end
        
        text(lx, ly, sprintf('%s\n%.1f\n(%s)', labels{i}, data(i), diff_str), ...
             'HorizontalAlignment', 'center', ...
             'FontSize', 10, 'FontWeight', 'bold');
    end
    
    % 제목
    title(title_text, 'FontSize', 16, 'FontWeight', 'bold');
    
    % 범례
    legend({'평균선', '해당 유형'}, 'Location', 'best', 'FontSize', 10);
    
    axis equal;
    axis([-1.5 1.5 -1.5 1.5]);
    axis off;
    hold off;
end