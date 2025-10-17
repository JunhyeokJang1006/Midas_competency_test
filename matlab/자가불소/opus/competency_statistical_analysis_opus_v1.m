%% 인재유형 종합 분석 시스템
% HR Talent Type Comprehensive Analysis System
% 목적: 1) 인재 유형별 프로파일링
%      2) 상관 기반 가중치 부여
%      3) 머신러닝 기법을 이용한 예측 분석

clear; clc; close all;

%% ========================================================================
%                          PART 1: 데이터 준비 및 전처리
% =========================================================================

fprintf('═══════════════════════════════════════════════════════════\n');
fprintf('         인재유형 종합 분석 시스템 v2.0\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 11);
set(0, 'DefaultTextFontSize', 11);
set(0, 'DefaultLineLineWidth', 1.5);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.output_dir = pwd;

% 성과 순위 정의 (사용자 제공)
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '위장형 소화성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 2, 1]);

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
    comp_total = readtable(config.comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 상위항목 데이터: %d명\n', height(comp_upper));
    fprintf('  ✓ 종합점수 데이터: %d명\n', height(comp_total));
    
catch ME
    error('데이터 로딩 실패: %s', ME.message);
end

%% 1.2 인재유형 데이터 추출
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
fprintf('▶ 유효한 인재유형 데이터: %d명\n', height(hr_clean));

% 인재유형 분포 분석
talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('\n인재유형 분포:\n');
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

fprintf('▶ 역량 ID 컬럼: %s\n', comp_upper.Properties.VariableNames{comp_id_col});

% 유효한 역량 컬럼 추출 (숫자형이고 변동성이 있는 컬럼)
valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)  % 메타데이터 컬럼 제외
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
fprintf('  역량 목록:\n');
for i = 1:min(5, length(valid_comp_cols))
    fprintf('    %d. %s\n', i, valid_comp_cols{i});
end
if length(valid_comp_cols) > 5
    fprintf('    ... 외 %d개\n', length(valid_comp_cols)-5);
end

%% 1.4 ID 매칭 및 데이터 통합
fprintf('\n【STEP 4】 데이터 매칭 및 통합\n');
fprintf('────────────────────────────────────────────\n');

% ID 표준화
hr_ids_str = arrayfun(@(x) sprintf('%.0f', x), hr_clean.ID, 'UniformOutput', false);
comp_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_upper{:, comp_id_col}, 'UniformOutput', false);

% 교집합 찾기
[matched_ids, hr_idx, comp_idx] = intersect(hr_ids_str, comp_ids_str);

if length(matched_ids) < 10
    error('매칭된 데이터가 부족합니다: %d명', length(matched_ids));
end

fprintf('▶ 매칭 성공: %d명\n', length(matched_ids));

% 매칭된 데이터 추출
matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, valid_comp_indices);
matched_talent_types = matched_hr{:, talent_col_idx};

% 종합점수 매칭
comp_total_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_total{:, 1}, 'UniformOutput', false);
[~, ~, total_idx] = intersect(matched_ids, comp_total_ids_str);

if ~isempty(total_idx)
    total_scores = comp_total{total_idx, end};
    fprintf('▶ 종합점수 통합: %d명\n', length(total_idx));
else
    total_scores = [];
    fprintf('⚠ 종합점수 데이터 없음\n');
end

%% ========================================================================
%                    PART 2: 인재유형별 프로파일링
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 2: 인재유형별 프로파일링\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 2.1 기술통계 분석
fprintf('【STEP 5】 인재유형별 기술통계\n');
fprintf('────────────────────────────────────────────\n');

unique_matched_types = unique(matched_talent_types);
n_types = length(unique_matched_types);

% 통계 테이블 초기화
profile_stats = table();
profile_stats.TalentType = unique_matched_types;
profile_stats.Count = zeros(n_types, 1);
profile_stats.CompetencyMean = zeros(n_types, 1);
profile_stats.CompetencyStd = zeros(n_types, 1);
profile_stats.TotalScoreMean = zeros(n_types, 1);
profile_stats.PerformanceRank = zeros(n_types, 1);

% 각 유형별 역량 프로파일 저장
type_profiles = cell(n_types, 1);

for i = 1:n_types
    type_name = unique_matched_types{i};
    type_mask = strcmp(matched_talent_types, type_name);
    
    % 기본 통계
    profile_stats.Count(i) = sum(type_mask);
    
    % 역량 점수 통계
    type_comp_data = matched_comp{type_mask, :};
    profile_stats.CompetencyMean(i) = nanmean(type_comp_data(:));
    profile_stats.CompetencyStd(i) = nanstd(type_comp_data(:));
    
    % 종합점수 통계
    if ~isempty(total_scores)
        type_total_scores = total_scores(type_mask);
        profile_stats.TotalScoreMean(i) = nanmean(type_total_scores);
    end
    
    % 성과 순위
    if config.performance_ranking.isKey(type_name)
        profile_stats.PerformanceRank(i) = config.performance_ranking(type_name);
    end
    
    % 상세 프로파일 저장
    type_profiles{i} = nanmean(type_comp_data, 1);
end

% 결과 출력
fprintf('\n인재유형별 통계 요약:\n');
fprintf('%-20s | 인원 | 역량평균 | 표준편차 | 종합점수 | 성과순위\n', '유형');
fprintf('%s\n', repmat('-', 70, 1));

for i = 1:height(profile_stats)
    fprintf('%-20s | %4d | %8.2f | %8.2f | %8.2f | %8.0f\n', ...
        profile_stats.TalentType{i}, profile_stats.Count(i), ...
        profile_stats.CompetencyMean(i), profile_stats.CompetencyStd(i), ...
        profile_stats.TotalScoreMean(i), profile_stats.PerformanceRank(i));
end

%% 2.2 프로파일 시각화 - 레이더 차트
fprintf('\n【STEP 6】 인재유형 프로파일 시각화\n');
fprintf('────────────────────────────────────────────\n');

% 시각화할 주요 역량 선정 (분산이 큰 상위 10개)
comp_variance = var(table2array(matched_comp), 0, 1, 'omitnan');
[~, var_idx] = sort(comp_variance, 'descend');
top_comp_idx = var_idx(1:min(10, length(var_idx)));
top_comp_names = valid_comp_cols(top_comp_idx);

% 전체 평균 계산
overall_mean_profile = nanmean(table2array(matched_comp), 1);

% 레이더 차트 생성
figure('Position', [100, 100, 1600, 1000], 'Color', 'white');
colormap_types = lines(n_types);

n_rows = ceil(sqrt(n_types));
n_cols = ceil(n_types / n_rows);

for i = 1:n_types
    subplot(n_rows, n_cols, i);
    
    % 해당 유형의 프로파일 데이터
    type_profile = type_profiles{i}(top_comp_idx);
    baseline_profile = overall_mean_profile(top_comp_idx);
    
    % 레이더 차트 그리기
    createRadarChart(type_profile, baseline_profile, top_comp_names, ...
                    unique_matched_types{i}, colormap_types(i,:));
end

sgtitle('인재유형별 역량 프로파일 (Top 10 주요 역량)', ...
        'FontSize', 16, 'FontWeight', 'bold');

%% ========================================================================
%                    PART 3: 상관 기반 가중치 분석
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 3: 상관 기반 가중치 분석\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 3.1 성과점수 계산
fprintf('【STEP 7】 성과점수 기반 상관분석\n');
fprintf('────────────────────────────────────────────\n');

% 각 개인의 성과점수 할당
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    type_name = matched_talent_types{i};
    if config.performance_ranking.isKey(type_name)
        performance_scores(i) = config.performance_ranking(type_name);
    end
end

% 유효한 데이터만 선택
valid_perf_idx = performance_scores > 0;
valid_performance = performance_scores(valid_perf_idx);
valid_competencies = table2array(matched_comp(valid_perf_idx, :));

fprintf('▶ 성과점수 할당 완료: %d명\n', sum(valid_perf_idx));

%% 3.2 역량별 상관계수 계산
fprintf('\n【STEP 8】 역량-성과 상관분석\n');
fprintf('────────────────────────────────────────────\n');

n_competencies = length(valid_comp_cols);
correlation_results = table();
correlation_results.Competency = valid_comp_cols';
correlation_results.Correlation = zeros(n_competencies, 1);
correlation_results.PValue = zeros(n_competencies, 1);
correlation_results.HighPerf_Mean = zeros(n_competencies, 1);
correlation_results.LowPerf_Mean = zeros(n_competencies, 1);
correlation_results.Difference = zeros(n_competencies, 1);

% 성과 상위/하위 그룹 분류
perf_median = median(valid_performance);
high_perf_idx = valid_performance > perf_median;
low_perf_idx = valid_performance <= perf_median;

for i = 1:n_competencies
    comp_scores = valid_competencies(:, i);
    valid_idx = ~isnan(comp_scores);
    
    if sum(valid_idx) >= 10
        % 상관계수 계산
        [r, p] = corr(comp_scores(valid_idx), valid_performance(valid_idx), ...
                     'Type', 'Spearman');
        correlation_results.Correlation(i) = r;
        correlation_results.PValue(i) = p;
        
        % 그룹별 평균
        correlation_results.HighPerf_Mean(i) = nanmean(comp_scores(high_perf_idx));
        correlation_results.LowPerf_Mean(i) = nanmean(comp_scores(low_perf_idx));
        correlation_results.Difference(i) = correlation_results.HighPerf_Mean(i) - ...
                                           correlation_results.LowPerf_Mean(i);
    end
end

% 상관계수 기반 정렬
correlation_results = sortrows(correlation_results, 'Correlation', 'descend');

fprintf('\n상위 10개 성과 예측 역량:\n');
fprintf('%-30s | 상관계수 | p-value | 차이\n', '역량');
fprintf('%s\n', repmat('-', 65, 1));

for i = 1:min(10, height(correlation_results))
    significance = '';
    if correlation_results.PValue(i) < 0.001
        significance = '***';
    elseif correlation_results.PValue(i) < 0.01
        significance = '**';
    elseif correlation_results.PValue(i) < 0.05
        significance = '*';
    end
    
    fprintf('%-30s | %8.4f%s | %7.4f | %6.2f\n', ...
        correlation_results.Competency{i}, ...
        correlation_results.Correlation(i), significance, ...
        correlation_results.PValue(i), ...
        correlation_results.Difference(i));
end

%% 3.3 가중치 계산 및 정규화
fprintf('\n【STEP 9】 상관 기반 가중치 계산\n');
fprintf('────────────────────────────────────────────\n');

% 양의 상관계수만 사용하여 가중치 계산
positive_corr = max(0, correlation_results.Correlation);
weights_raw = abs(positive_corr);
weights_normalized = weights_raw / sum(weights_raw);

correlation_results.Weight = weights_normalized * 100;

% 상위 가중치 역량
top_weighted = correlation_results(1:min(10, height(correlation_results)), :);

fprintf('\n상관 기반 가중치 (상위 10개):\n');
for i = 1:height(top_weighted)
    fprintf('  %2d. %-30s: %5.2f%%\n', i, ...
        top_weighted.Competency{i}, top_weighted.Weight(i));
end

% 누적 가중치 계산
cumulative_weight = cumsum(correlation_results.Weight);
n_features_80 = find(cumulative_weight >= 80, 1);
fprintf('\n▶ 80%% 설명력을 위한 필요 역량 수: %d개\n', n_features_80);

%% 3.4 가중치 시각화
figure('Position', [100, 100, 1400, 800], 'Color', 'white');

% Subplot 1: 상관계수 분포
subplot(2, 3, 1);
histogram(correlation_results.Correlation, 20, 'FaceColor', [0.3, 0.6, 0.9]);
xlabel('상관계수');
ylabel('빈도');
title('역량-성과 상관계수 분포');
grid on;

% Subplot 2: 상위 10개 가중치
subplot(2, 3, 2);
top_10 = correlation_results(1:min(10, height(correlation_results)), :);
barh(10:-1:1, top_10.Weight(1:min(10, height(top_10))), 'FaceColor', [0.9, 0.3, 0.3]);
set(gca, 'YTick', 1:10, 'YTickLabel', flip(top_10.Competency(1:min(10, height(top_10)))));
xlabel('가중치 (%)');
title('상위 10개 역량 가중치');
grid on;

% Subplot 3: 누적 가중치
subplot(2, 3, 3);
plot(cumulative_weight, 'LineWidth', 2, 'Color', [0.3, 0.7, 0.3]);
hold on;
plot([1, length(cumulative_weight)], [80, 80], 'r--', 'LineWidth', 1.5);
xlabel('역량 개수');
ylabel('누적 가중치 (%)');
title('누적 설명력');
legend('누적 가중치', '80% 기준선', 'Location', 'southeast');
grid on;

% Subplot 4: 성과 상위 vs 하위 비교
subplot(2, 3, 4:6);
x = 1:min(15, height(correlation_results));
bar_data = [correlation_results.HighPerf_Mean(x), correlation_results.LowPerf_Mean(x)];
bar(x, bar_data);
set(gca, 'XTick', x, 'XTickLabel', correlation_results.Competency(x), ...
    'XTickLabelRotation', 45);
ylabel('평균 점수');
legend('성과 상위', '성과 하위', 'Location', 'northwest');
title('성과 그룹별 역량 점수 비교');
grid on;

sgtitle('상관 기반 가중치 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');

%% ========================================================================
%                    PART 4: 머신러닝 예측 분석
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 4: 머신러닝 예측 분석\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 4.1 데이터 준비
fprintf('【STEP 10】 머신러닝 데이터 준비\n');
fprintf('────────────────────────────────────────────\n');

% 특성 및 레이블 준비
X = table2array(matched_comp);
y = matched_talent_types;
[y_unique, ~, y_encoded] = unique(y);

% 데이터 정규화
X_normalized = normalize(X, 'range');

% 훈련/테스트 분할 (80/20)
cv_partition = cvpartition(y_encoded, 'HoldOut', 0.2, 'Stratify', true);
X_train = X_normalized(cv_partition.training, :);
y_train = y_encoded(cv_partition.training);
X_test = X_normalized(cv_partition.test, :);
y_test = y_encoded(cv_partition.test);

fprintf('▶ 훈련 데이터: %d명\n', length(y_train));
fprintf('▶ 테스트 데이터: %d명\n', length(y_test));
fprintf('▶ 클래스 수: %d개\n', length(y_unique));

%% 4.2 Random Forest 모델
fprintf('\n【STEP 11】 Random Forest 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% 하이퍼파라미터 최적화
fprintf('하이퍼파라미터 최적화 중...\n');
rng(42); % 재현성을 위한 시드 설정

% 최적 파라미터 탐색 (간단한 그리드 서치)
n_trees_options = [50, 100, 150];
min_leaf_options = [1, 3, 5];

best_accuracy = 0;
best_params = struct();

for n_trees = n_trees_options
    for min_leaf = min_leaf_options
        % 5-fold 교차검증
        cv_inner = cvpartition(y_train, 'KFold', 5);
        cv_accuracies = zeros(5, 1);
        
        for fold = 1:5
            X_train_fold = X_train(cv_inner.training(fold), :);
            y_train_fold = y_train(cv_inner.training(fold));
            X_val_fold = X_train(cv_inner.test(fold), :);
            y_val_fold = y_train(cv_inner.test(fold));
            
            % 모델 학습
            rf_temp = TreeBagger(n_trees, X_train_fold, y_train_fold, ...
                'Method', 'classification', ...
                'MinLeafSize', min_leaf);
            
            % 예측 및 평가
            y_pred_fold = cellfun(@str2double, predict(rf_temp, X_val_fold));
            cv_accuracies(fold) = sum(y_pred_fold == y_val_fold) / length(y_val_fold);
        end
        
        mean_accuracy = mean(cv_accuracies);
        if mean_accuracy > best_accuracy
            best_accuracy = mean_accuracy;
            best_params.n_trees = n_trees;
            best_params.min_leaf = min_leaf;
        end
    end
end

fprintf('최적 파라미터: Trees=%d, MinLeaf=%d, CV정확도=%.4f\n', ...
    best_params.n_trees, best_params.min_leaf, best_accuracy);

% 최적 파라미터로 최종 모델 학습
rf_model = TreeBagger(best_params.n_trees, X_train, y_train, ...
    'Method', 'classification', ...
    'MinLeafSize', best_params.min_leaf, ...
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', valid_comp_cols);

% 테스트 세트 평가
y_pred_rf = cellfun(@str2double, predict(rf_model, X_test));
rf_accuracy = sum(y_pred_rf == y_test) / length(y_test);

fprintf('Random Forest 테스트 정확도: %.4f\n', rf_accuracy);

% Feature Importance
rf_importance = rf_model.OOBPermutedPredictorDeltaError;
rf_importance_norm = rf_importance / sum(rf_importance);

%% 4.3 Gradient Boosting 모델
fprintf('\n【STEP 12】 Gradient Boosting 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% AdaBoost 사용 (MATLAB의 Gradient Boosting 구현)
gb_model = fitcensemble(X_train, y_train, ...
    'Method', 'AdaBoostM2', ...
    'NumLearningCycles', 100, ...
    'Learners', templateTree('MaxNumSplits', 10));

% 테스트 세트 평가
y_pred_gb = predict(gb_model, X_test);
gb_accuracy = sum(y_pred_gb == y_test) / length(y_test);

fprintf('Gradient Boosting 테스트 정확도: %.4f\n', gb_accuracy);

% Feature Importance
gb_importance = predictorImportance(gb_model);
gb_importance_norm = gb_importance / sum(gb_importance);

%% 4.4 앙상블 모델 (투표)
fprintf('\n【STEP 13】 앙상블 모델 평가\n');
fprintf('────────────────────────────────────────────\n');

% 투표 기반 앙상블
y_pred_ensemble = mode([y_pred_rf, y_pred_gb], 2);
ensemble_accuracy = sum(y_pred_ensemble == y_test) / length(y_test);

fprintf('앙상블 모델 테스트 정확도: %.4f\n', ensemble_accuracy);

%% 4.5 성능 평가 메트릭
fprintf('\n【STEP 14】 상세 성능 평가\n');
fprintf('────────────────────────────────────────────\n');

% Confusion Matrix 계산
conf_rf = confusionmat(y_test, y_pred_rf);
conf_gb = confusionmat(y_test, y_pred_gb);
conf_ensemble = confusionmat(y_test, y_pred_ensemble);

% 클래스별 성능 메트릭
n_classes = length(y_unique);
class_metrics = table();
class_metrics.TalentType = y_unique;
class_metrics.Precision = zeros(n_classes, 1);
class_metrics.Recall = zeros(n_classes, 1);
class_metrics.F1Score = zeros(n_classes, 1);
class_metrics.Support = zeros(n_classes, 1);

for i = 1:n_classes
    tp = conf_ensemble(i, i);
    fp = sum(conf_ensemble(:, i)) - tp;
    fn = sum(conf_ensemble(i, :)) - tp;
    
    class_metrics.Precision(i) = tp / (tp + fp + eps);
    class_metrics.Recall(i) = tp / (tp + fn + eps);
    class_metrics.F1Score(i) = 2 * (class_metrics.Precision(i) * class_metrics.Recall(i)) / ...
                              (class_metrics.Precision(i) + class_metrics.Recall(i) + eps);
    class_metrics.Support(i) = sum(y_test == i);
end

fprintf('\n클래스별 성능 (앙상블 모델):\n');
fprintf('%-20s | Precision | Recall | F1-Score | Support\n', '인재유형');
fprintf('%s\n', repmat('-', 65, 1));

for i = 1:height(class_metrics)
    fprintf('%-20s | %9.4f | %6.4f | %8.4f | %7d\n', ...
        class_metrics.TalentType{i}, ...
        class_metrics.Precision(i), ...
        class_metrics.Recall(i), ...
        class_metrics.F1Score(i), ...
        class_metrics.Support(i));
end

fprintf('\n전체 성능 요약:\n');
fprintf('  • Macro-avg Precision: %.4f\n', mean(class_metrics.Precision));
fprintf('  • Macro-avg Recall: %.4f\n', mean(class_metrics.Recall));
fprintf('  • Macro-avg F1-Score: %.4f\n', mean(class_metrics.F1Score));
fprintf('  • Accuracy: %.4f\n', ensemble_accuracy);

%% 4.6 Feature Importance 종합
fprintf('\n【STEP 15】 종합 Feature Importance\n');
fprintf('────────────────────────────────────────────\n');

% 세 가지 방법의 가중 평균
ensemble_importance = (rf_importance_norm + gb_importance_norm + weights_normalized') / 3;
ensemble_importance = ensemble_importance / sum(ensemble_importance);

% Feature Importance 테이블 생성
importance_table = table();
importance_table.Competency = valid_comp_cols';
importance_table.RF_Importance = rf_importance_norm * 100;
importance_table.GB_Importance = gb_importance_norm * 100;
importance_table.Correlation_Weight = weights_normalized * 100;
importance_table.Ensemble_Importance = ensemble_importance * 100;

importance_table = sortrows(importance_table, 'Ensemble_Importance', 'descend');

fprintf('\n종합 Feature Importance (상위 15개):\n');
fprintf('%-30s | RF(%) | GB(%) | Corr(%) | Final(%%)\n', '역량');
fprintf('%s\n', repmat('-', 70, 1));

for i = 1:min(15, height(importance_table))
    fprintf('%-30s | %5.2f | %5.2f | %7.2f | %7.2f\n', ...
        importance_table.Competency{i}, ...
        importance_table.RF_Importance(i), ...
        importance_table.GB_Importance(i), ...
        importance_table.Correlation_Weight(i), ...
        importance_table.Ensemble_Importance(i));
end

%% 4.7 시각화 - Confusion Matrix
figure('Position', [100, 100, 1400, 400], 'Color', 'white');

subplot(1, 3, 1);
heatmap(y_unique, y_unique, conf_rf);
title('Random Forest');
xlabel('예측'); ylabel('실제');

subplot(1, 3, 2);
heatmap(y_unique, y_unique, conf_gb);
title('Gradient Boosting');
xlabel('예측'); ylabel('실제');

subplot(1, 3, 3);
heatmap(y_unique, y_unique, conf_ensemble);
title('Ensemble Model');
xlabel('예측'); ylabel('실제');

sgtitle('모델별 Confusion Matrix', 'FontSize', 14, 'FontWeight', 'bold');

%% 4.8 시각화 - Feature Importance
figure('Position', [100, 100, 1400, 600], 'Color', 'white');

% 상위 15개 역량만 표시
top_n = min(15, height(importance_table));
x = 1:top_n;

bar_data = [importance_table.RF_Importance(1:top_n), ...
           importance_table.GB_Importance(1:top_n), ...
           importance_table.Correlation_Weight(1:top_n)];

bar(x, bar_data);
set(gca, 'XTick', x, 'XTickLabel', importance_table.Competency(1:top_n), ...
    'XTickLabelRotation', 45);
ylabel('Importance (%)');
xlabel('역량');
title('Feature Importance 비교 (상위 15개)', 'FontSize', 14, 'FontWeight', 'bold');
legend('Random Forest', 'Gradient Boosting', 'Correlation', 'Location', 'northeast');
grid on;

%% ========================================================================
%                          PART 5: 결과 저장 및 보고서
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 5: 결과 저장 및 최종 보고서\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 5.1 결과 저장
fprintf('【STEP 16】 분석 결과 저장\n');
fprintf('────────────────────────────────────────────\n');

% MATLAB 파일 저장
analysis_results = struct();
analysis_results.profile_stats = profile_stats;
analysis_results.correlation_results = correlation_results;
analysis_results.importance_table = importance_table;
analysis_results.class_metrics = class_metrics;
analysis_results.models = struct('rf', rf_model, 'gb', gb_model);
analysis_results.accuracies = struct('rf', rf_accuracy, 'gb', gb_accuracy, ...
                                    'ensemble', ensemble_accuracy);

save('talent_analysis_complete.mat', 'analysis_results');
fprintf('✓ MATLAB 파일 저장: talent_analysis_complete.mat\n');

% Excel 파일 저장
try
    % Sheet 1: 인재유형 프로파일
    writetable(profile_stats, 'talent_analysis_report.xlsx', ...
              'Sheet', 'TalentProfiles');
    
    % Sheet 2: 상관분석 결과
    writetable(correlation_results(1:min(30, height(correlation_results)), :), ...
              'talent_analysis_report.xlsx', 'Sheet', 'CorrelationAnalysis');
    
    % Sheet 3: Feature Importance
    writetable(importance_table(1:min(30, height(importance_table)), :), ...
              'talent_analysis_report.xlsx', 'Sheet', 'FeatureImportance');
    
    % Sheet 4: 모델 성능
    writetable(class_metrics, 'talent_analysis_report.xlsx', ...
              'Sheet', 'ModelPerformance');
    
    fprintf('✓ Excel 파일 저장: talent_analysis_report.xlsx\n');
catch
    fprintf('⚠ Excel 저장 실패 (파일이 열려있을 수 있음)\n');
end

%% 5.2 최종 보고서
fprintf('\n%s\n', repmat('═', 80, 1));
fprintf('                        최종 분석 보고서\n');
fprintf('%s\n', repmat('═', 80, 1));

fprintf('\n📊 데이터 요약:\n');
fprintf('  • 분석 대상: %d명\n', length(matched_ids));
fprintf('  • 인재유형: %d개\n', n_types);
fprintf('  • 역량항목: %d개\n', length(valid_comp_cols));

fprintf('\n🎯 주요 발견사항:\n');
fprintf('  1. 최고 성과 인재유형: %s (평균 %.2f점)\n', ...
    profile_stats.TalentType{profile_stats.PerformanceRank == max(profile_stats.PerformanceRank)}, ...
    profile_stats.CompetencyMean(profile_stats.PerformanceRank == max(profile_stats.PerformanceRank)));

fprintf('  2. 최대 인원 인재유형: %s (%d명)\n', ...
    profile_stats.TalentType{profile_stats.Count == max(profile_stats.Count)}, ...
    max(profile_stats.Count));

fprintf('  3. 핵심 예측 역량 Top 3:\n');
for i = 1:min(3, height(importance_table))
    fprintf('     - %s (%.2f%%)\n', importance_table.Competency{i}, ...
            importance_table.Ensemble_Importance(i));
end

fprintf('\n🤖 머신러닝 모델 성능:\n');
fprintf('  • Random Forest: %.2f%%\n', rf_accuracy * 100);
fprintf('  • Gradient Boosting: %.2f%%\n', gb_accuracy * 100);
fprintf('  • Ensemble Model: %.2f%%\n', ensemble_accuracy * 100);

fprintf('\n✨ 권장사항:\n');
fprintf('  1. %s 역량을 중점적으로 평가\n', importance_table.Competency{1});
fprintf('  2. 상위 %d개 역량으로 80%% 예측력 달성 가능\n', n_features_80);
fprintf('  3. %s 인재유형 육성 프로그램 우선 개발 권장\n', ...
    profile_stats.TalentType{profile_stats.PerformanceRank == max(profile_stats.PerformanceRank)});

fprintf('\n%s\n', repmat('═', 80, 1));
fprintf('               분석 완료 - %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('%s\n', repmat('═', 80, 1));

%% Helper Functions

function createRadarChart(data, baseline, labels, title_text, color)
    % 레이더 차트 생성 함수
    n_vars = length(data);
    angles = linspace(0, 2*pi, n_vars+1);
    
    % 데이터 정규화 (0-1 스케일)
    data_norm = data / 100;
    baseline_norm = baseline / 100;
    
    % 순환을 위해 첫 번째 값을 마지막에 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];
    
    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);
    
    hold on;
    
    % 그리드 그리기
    for r = 0.2:0.2:1
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
    end
    
    % 방사선 그리기
    for i = 1:n_vars
        plot([0, cos(angles(i))], [0, sin(angles(i))], 'k:', ...
             'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
    end
    
    % 기준선 (전체 평균)
    plot(x_base, y_base, '--', 'Color', [0.5 0.5 0.5], 'LineWidth', 1.5);
    
    % 데이터 플롯
    patch(x_data, y_data, color, 'FaceAlpha', 0.3, 'EdgeColor', color, 'LineWidth', 2);
    
    % 데이터 포인트
    scatter(x_data(1:end-1), y_data(1:end-1), 30, color, 'filled');
    
    % 레이블
    label_radius = 1.15;
    for i = 1:n_vars
        [lx, ly] = pol2cart(angles(i), label_radius);
        text(lx, ly, sprintf('%s\n%.0f', labels{i}, data(i)), ...
             'HorizontalAlignment', 'center', 'FontSize', 8);
    end
    
    % 제목
    title(title_text, 'FontSize', 11, 'FontWeight', 'bold');
    
    axis equal;
    axis([-1.3 1.3 -1.3 1.3]);
    axis off;
    hold off;
end