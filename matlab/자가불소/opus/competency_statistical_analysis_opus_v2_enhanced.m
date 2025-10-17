%% 인재유형 종합 분석 시스템 v2.0 Enhanced
% HR Talent Type Comprehensive Analysis System
% 목적: 1) 인재 유형별 프로파일링
%      2) 상관 기반 가중치 부여
%      3) 머신러닝 기법을 이용한 예측 분석
%      4) 고성과자/저성과자 이진 분류 추가 분석

clear; clc; close all;

%% ========================================================================
%                          PART 1: 데이터 준비 및 전처리
% =========================================================================

fprintf('═══════════════════════════════════════════════════════════\n');
fprintf('         인재유형 종합 분석 시스템 v2.0 Enhanced\n');
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

% 성과 순위 정의 (사용자 제공) - 위장형 소화성 제외
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 1]);

% 제외할 인재유형 설정
config.excluded_types = {'위장형 소화성'};

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

%% 1.2 인재유형 데이터 추출 및 필터링
fprintf('\n【STEP 2】 인재유형 데이터 추출 및 정제 (위장형 소화성 제외)\n');
fprintf('────────────────────────────────────────────\n');

% 인재유형 컬럼 찾기
talent_col_idx = find(contains(hr_data.Properties.VariableNames, {'인재유형', '인재', '유형'}), 1);
if isempty(talent_col_idx)
    error('인재유형 컬럼을 찾을 수 없습니다.');
end

talent_col_name = hr_data.Properties.VariableNames{talent_col_idx};
fprintf('▶ 인재유형 컬럼: %s\n', talent_col_name);

% 빈 값 및 제외 유형 제거
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_idx}), :);

% 위장형 소화성 제외
excluded_mask = false(height(hr_clean), 1);
for i = 1:length(config.excluded_types)
    excluded_mask = excluded_mask | strcmp(hr_clean{:, talent_col_idx}, config.excluded_types{i});
end
hr_clean = hr_clean(~excluded_mask, :);

fprintf('▶ 유효한 인재유형 데이터: %d명 (위장형 소화성 제외)\n', height(hr_clean));

% 인재유형 분포 분석
talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('\n인재유형 분포 (위장형 소화성 제외):\n');
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

%% ========================================================================
%                    PART 3: 상관 기반 가중치 분석
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 3: 상관 기반 가중치 분석\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 3.1 성과점수 계산
fprintf('【STEP 6】 성과점수 기반 상관분석\n');
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
fprintf('\n【STEP 7】 역량-성과 상관분석\n');
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

% 양의 상관계수만 사용하여 가중치 계산
positive_corr = max(0, correlation_results.Correlation);
weights_raw = abs(positive_corr);
weights_normalized = weights_raw / sum(weights_raw);

correlation_results.Weight = weights_normalized * 100;

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

%% ========================================================================
%          PART 4: 고성과자/저성과자 이진 분류 고급 머신러닝 분석
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('         PART 4: 고성과자/저성과자 이진 분류 머신러닝 분석\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 4.1 고성과자/저성과자 레이블 생성
fprintf('【STEP 8】 고성과자/저성과자 레이블 생성\n');
fprintf('────────────────────────────────────────────\n');

% 성과 순위 기반 이진 분류 (상위 50% vs 하위 50%)
performance_threshold = median(valid_performance);

% 고성과자(1) vs 저성과자(0) 레이블 생성
high_perf_labels = double(valid_performance > performance_threshold);

fprintf('성과 기준 분류:\n');
fprintf('  • 고성과자 (상위 50%%): %d명\n', sum(high_perf_labels == 1));
fprintf('  • 저성과자 (하위 50%%): %d명\n', sum(high_perf_labels == 0));
fprintf('  • 임계값: %.2f\n', performance_threshold);

% 데이터 정규화
X_perf = normalize(valid_competencies, 'range');
y_perf = high_perf_labels;

% 훈련/테스트 분할 (80/20) - 수동 구현
n_samples = length(y_perf);
test_ratio = 0.2;

% 계층화 분할
high_indices = find(y_perf == 1);
low_indices = find(y_perf == 0);

n_high_test = max(1, round(length(high_indices) * test_ratio));
n_low_test = max(1, round(length(low_indices) * test_ratio));

% 랜덤 선택
rng(42);
high_test_idx = high_indices(randperm(length(high_indices), n_high_test));
low_test_idx = low_indices(randperm(length(low_indices), n_low_test));

test_indices = [high_test_idx; low_test_idx];
train_indices = setdiff(1:n_samples, test_indices);

X_train_perf = X_perf(train_indices, :);
y_train_perf = y_perf(train_indices);
X_test_perf = X_perf(test_indices, :);
y_test_perf = y_perf(test_indices);

fprintf('\n데이터 분할 완료:\n');
fprintf('  • 훈련 데이터: %d명 (고성과자 %d명, 저성과자 %d명)\n', ...
        length(y_train_perf), sum(y_train_perf == 1), sum(y_train_perf == 0));
fprintf('  • 테스트 데이터: %d명 (고성과자 %d명, 저성과자 %d명)\n', ...
        length(y_test_perf), sum(y_test_perf == 1), sum(y_test_perf == 0));

%% 4.2 트리 기반 배깅 모델 (Random Forest + Extra Trees)
fprintf('\n【STEP 9】 트리 기반 배깅 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% 1. Random Forest
fprintf('1. Random Forest 학습 중...\n');
rf_perf_model = TreeBagger(200, X_train_perf, y_train_perf, ...
    'Method', 'classification', ...
    'MinLeafSize', 3, ...
    'MaxNumSplits', 30, ...
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', valid_comp_cols);

[y_pred_rf_perf_cell, rf_perf_scores] = predict(rf_perf_model, X_test_perf);
y_pred_rf_perf = cellfun(@str2double, y_pred_rf_perf_cell);

% TreeBagger 결과 처리
if iscell(rf_perf_scores)
    rf_perf_probs = cell2mat(rf_perf_scores);
else
    rf_perf_probs = rf_perf_scores;
end

rf_perf_accuracy = sum(y_pred_rf_perf == y_test_perf) / length(y_test_perf);
rf_perf_importance = rf_perf_model.OOBPermutedPredictorDeltaError;

fprintf('   Random Forest 정확도: %.4f\n', rf_perf_accuracy);

% 2. Extra Trees (Extremely Randomized Trees) - TreeBagger로 구현
fprintf('2. Extra Trees 학습 중...\n');
et_perf_model = TreeBagger(200, X_train_perf, y_train_perf, ...
    'Method', 'classification', ...
    'MinLeafSize', 1, ...
    'MaxNumSplits', size(X_train_perf, 2), ...  % 모든 특성 사용
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', valid_comp_cols);

[y_pred_et_perf_cell, et_perf_scores] = predict(et_perf_model, X_test_perf);
y_pred_et_perf = cellfun(@str2double, y_pred_et_perf_cell);

% TreeBagger 결과 처리
if iscell(et_perf_scores)
    et_perf_probs = cell2mat(et_perf_scores);
else
    et_perf_probs = et_perf_scores;
end

et_perf_accuracy = sum(y_pred_et_perf == y_test_perf) / length(y_test_perf);
et_perf_importance = et_perf_model.OOBPermutedPredictorDeltaError;

fprintf('   Extra Trees 정확도: %.4f\n', et_perf_accuracy);

%% 4.3 뉴럴네트워크 모델 (다층 퍼셉트론)
fprintf('\n【STEP 10】 뉴럴네트워크 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% MATLAB의 Neural Network Toolbox가 있는 경우 사용
if exist('patternnet', 'file')
    fprintf('Neural Network Toolbox를 이용한 MLP 학습 중...\n');

    % 네트워크 구성
    hidden_sizes = [20, 10];  % 은닉층 크기
    nn_model = patternnet(hidden_sizes);

    % 훈련 설정
    nn_model.trainParam.epochs = 200;
    nn_model.trainParam.goal = 1e-5;
    nn_model.trainParam.showWindow = false;

    % 데이터 전치 (Neural Network Toolbox 형식)
    X_train_nn = X_train_perf';
    y_train_nn = full(ind2vec(y_train_perf' + 1));  % 1-based 인덱스로 변환

    % 모델 학습
    [nn_model, tr] = train(nn_model, X_train_nn, y_train_nn);

    % 예측
    X_test_nn = X_test_perf';
    nn_outputs = nn_model(X_test_nn);
    [~, y_pred_nn_perf] = max(nn_outputs);
    y_pred_nn_perf = y_pred_nn_perf' - 1;  % 0-based로 다시 변환

    nn_perf_accuracy = sum(y_pred_nn_perf == y_test_perf) / length(y_test_perf);

    % Feature Importance (가중치 분석)
    weights1 = nn_model.IW{1,1};  % 입력층 -> 첫 번째 은닉층
    nn_perf_importance = mean(abs(weights1), 1)';  % 절댓값 평균

    fprintf('   Neural Network 정확도: %.4f\n', nn_perf_accuracy);
else
    fprintf('⚠ Neural Network Toolbox가 없어 간단한 로지스틱 회귀로 대체\n');

    % 로지스틱 회귀 대체
    [B, ~, stats] = glmfit(X_train_perf, y_train_perf, 'binomial', 'link', 'logit');

    % 예측
    nn_pred_scores = glmval(B, X_test_perf, 'logit');
    y_pred_nn_perf = double(nn_pred_scores > 0.5);

    nn_perf_accuracy = sum(y_pred_nn_perf == y_test_perf) / length(y_test_perf);
    nn_perf_importance = abs(B(2:end));  % 절편 제외

    fprintf('   Logistic Regression 정확도: %.4f\n', nn_perf_accuracy);
end

%% 4.4 그레디언트 부스팅 모델
fprintf('\n【STEP 11】 그레디언트 부스팅 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

fprintf('Gradient Boosting (AdaBoost) 학습 중...\n');

% 적응적 부스팅
gb_perf_model = fitcensemble(X_train_perf, y_train_perf, ...
    'Method', 'AdaBoostM1', ...
    'NumLearningCycles', 100, ...
    'Learners', templateTree('MaxNumSplits', 20, 'MinLeafSize', 5), ...
    'LearnRate', 0.1);

% 예측
[y_pred_gb_perf, gb_perf_scores] = predict(gb_perf_model, X_test_perf);
gb_perf_accuracy = sum(y_pred_gb_perf == y_test_perf) / length(y_test_perf);

% Feature Importance
gb_perf_importance = predictorImportance(gb_perf_model);

fprintf('   Gradient Boosting 정확도: %.4f\n', gb_perf_accuracy);

%% 4.5 모델 성능 비교 및 통합 Feature Importance
fprintf('\n【STEP 12】 모델 성능 비교 및 통합 Feature Importance\n');
fprintf('────────────────────────────────────────────\n');

% 성능 요약
model_performance = table();
model_performance.Model = {'Random Forest'; 'Extra Trees'; 'Neural Network'; 'Gradient Boosting'};
model_performance.Accuracy = [rf_perf_accuracy; et_perf_accuracy; nn_perf_accuracy; gb_perf_accuracy];

fprintf('\n이진 분류 모델 성능 비교:\n');
fprintf('%-20s | 정확도\n', '모델');
fprintf('─────────────────────────────────\n');
for i = 1:height(model_performance)
    fprintf('%-20s | %6.2f%%\n', model_performance.Model{i}, model_performance.Accuracy(i) * 100);
end

% Feature Importance 정규화
rf_perf_importance_norm = rf_perf_importance / sum(rf_perf_importance);
et_perf_importance_norm = et_perf_importance / sum(et_perf_importance);
nn_perf_importance_norm = nn_perf_importance / sum(nn_perf_importance);
gb_perf_importance_norm = gb_perf_importance / sum(gb_perf_importance);

% 성능 기반 가중 평균 (정확도에 비례한 가중치)
total_accuracy = sum(model_performance.Accuracy);
rf_weight = rf_perf_accuracy / total_accuracy;
et_weight = et_perf_accuracy / total_accuracy;
nn_weight = nn_perf_accuracy / total_accuracy;
gb_weight = gb_perf_accuracy / total_accuracy;

% 통합 Feature Importance 계산
integrated_importance = rf_weight * rf_perf_importance_norm + ...
                       et_weight * et_perf_importance_norm + ...
                       nn_weight * nn_perf_importance_norm + ...
                       gb_weight * gb_perf_importance_norm;

% Feature Importance 테이블 생성 (차원 호환성 개선)
n_comp = length(valid_comp_cols);

% 모든 importance 벡터를 열 벡터로 변환하고 길이 맞춤
rf_perf_importance_norm = rf_perf_importance_norm(:);
if length(rf_perf_importance_norm) ~= n_comp
    rf_perf_importance_norm = rf_perf_importance_norm(1:min(n_comp, length(rf_perf_importance_norm)));
    if length(rf_perf_importance_norm) < n_comp
        rf_perf_importance_norm = [rf_perf_importance_norm; zeros(n_comp - length(rf_perf_importance_norm), 1)];
    end
end

et_perf_importance_norm = et_perf_importance_norm(:);
if length(et_perf_importance_norm) ~= n_comp
    et_perf_importance_norm = et_perf_importance_norm(1:min(n_comp, length(et_perf_importance_norm)));
    if length(et_perf_importance_norm) < n_comp
        et_perf_importance_norm = [et_perf_importance_norm; zeros(n_comp - length(et_perf_importance_norm), 1)];
    end
end

nn_perf_importance_norm = nn_perf_importance_norm(:);
if length(nn_perf_importance_norm) ~= n_comp
    nn_perf_importance_norm = nn_perf_importance_norm(1:min(n_comp, length(nn_perf_importance_norm)));
    if length(nn_perf_importance_norm) < n_comp
        nn_perf_importance_norm = [nn_perf_importance_norm; zeros(n_comp - length(nn_perf_importance_norm), 1)];
    end
end

gb_perf_importance_norm = gb_perf_importance_norm(:);
if length(gb_perf_importance_norm) ~= n_comp
    gb_perf_importance_norm = gb_perf_importance_norm(1:min(n_comp, length(gb_perf_importance_norm)));
    if length(gb_perf_importance_norm) < n_comp
        gb_perf_importance_norm = [gb_perf_importance_norm; zeros(n_comp - length(gb_perf_importance_norm), 1)];
    end
end

integrated_importance = integrated_importance(:);
if length(integrated_importance) ~= n_comp
    integrated_importance = integrated_importance(1:min(n_comp, length(integrated_importance)));
    if length(integrated_importance) < n_comp
        integrated_importance = [integrated_importance; zeros(n_comp - length(integrated_importance), 1)];
    end
end

perf_importance_table = table();
perf_importance_table.Competency = valid_comp_cols';
perf_importance_table.RF_Importance = rf_perf_importance_norm * 100;
perf_importance_table.ET_Importance = et_perf_importance_norm * 100;
perf_importance_table.NN_Importance = nn_perf_importance_norm * 100;
perf_importance_table.GB_Importance = gb_perf_importance_norm * 100;
perf_importance_table.Integrated_Importance = integrated_importance * 100;

% 통합 중요도 기준 정렬
perf_importance_table = sortrows(perf_importance_table, 'Integrated_Importance', 'descend');

fprintf('\n통합 Feature Importance (상위 10개):\n');
fprintf('%-20s | RF(%%) | ET(%%) | NN(%%) | GB(%%) | 통합(%%)\n', '역량');
fprintf('───────────────────────────────────────────────────────────────\n');

for i = 1:min(10, height(perf_importance_table))
    fprintf('%-20s | %5.1f | %5.1f | %5.1f | %5.1f | %6.1f\n', ...
        perf_importance_table.Competency{i}, ...
        perf_importance_table.RF_Importance(i), ...
        perf_importance_table.ET_Importance(i), ...
        perf_importance_table.NN_Importance(i), ...
        perf_importance_table.GB_Importance(i), ...
        perf_importance_table.Integrated_Importance(i));
end

%% 4.6 최종 가중치 계산 (상관분석 + 머신러닝)
fprintf('\n【STEP 13】 최종 통합 가중치 계산\n');
fprintf('────────────────────────────────────────────\n');

% 상관분석 가중치와 머신러닝 가중치 결합
% 길이 맞춤
n_comp = length(valid_comp_cols);
if length(weights_normalized) ~= n_comp
    weights_normalized = weights_normalized(1:min(n_comp, length(weights_normalized)));
    if length(weights_normalized) < n_comp
        weights_normalized = [weights_normalized; zeros(n_comp - length(weights_normalized), 1)];
    end
end

% 최종 가중치 = 50% 상관분석 + 50% 머신러닝
final_weights = 0.5 * weights_normalized + 0.5 * integrated_importance;
final_weights = final_weights / sum(final_weights);

% 최종 가중치 테이블
final_weight_table = table();
final_weight_table.Competency = valid_comp_cols';
final_weight_table.Correlation_Weight = weights_normalized * 100;
final_weight_table.ML_Weight = integrated_importance * 100;
final_weight_table.Final_Weight = final_weights * 100;

final_weight_table = sortrows(final_weight_table, 'Final_Weight', 'descend');

fprintf('\n최종 통합 가중치 (상위 10개):\n');
fprintf('%-20s | 상관(%%) | ML(%%) | 최종(%%)\n', '역량');
fprintf('──────────────────────────────────────────────\n');

for i = 1:min(10, height(final_weight_table))
    fprintf('%-20s | %6.1f | %5.1f | %6.1f\n', ...
        final_weight_table.Competency{i}, ...
        final_weight_table.Correlation_Weight(i), ...
        final_weight_table.ML_Weight(i), ...
        final_weight_table.Final_Weight(i));
end

%% 4.7 시각화
fprintf('\n【STEP 14】 결과 시각화\n');
fprintf('────────────────────────────────────────────\n');

% Figure 1: 모델 성능 비교
figure('Position', [100, 100, 1400, 1000], 'Color', 'white');

% Subplot 1: 모델 정확도 비교
subplot(2, 3, 1);
bar(model_performance.Accuracy * 100, 'FaceColor', [0.3, 0.6, 0.9]);
set(gca, 'XTickLabel', model_performance.Model, 'XTickLabelRotation', 45);
ylabel('정확도 (%)');
title('이진 분류 모델 성능 비교');
ylim([0, 100]);
grid on;

% 각 막대에 정확도 표시
for i = 1:length(model_performance.Accuracy)
    text(i, model_performance.Accuracy(i) * 100 + 2, ...
         sprintf('%.1f%%', model_performance.Accuracy(i) * 100), ...
         'HorizontalAlignment', 'center', 'FontWeight', 'bold');
end

% Subplot 2: Feature Importance 히트맵 (상위 10개)
subplot(2, 3, 2:3);
top_10_idx = 1:min(10, height(perf_importance_table));
heatmap_data = [perf_importance_table.RF_Importance(top_10_idx), ...
                perf_importance_table.ET_Importance(top_10_idx), ...
                perf_importance_table.NN_Importance(top_10_idx), ...
                perf_importance_table.GB_Importance(top_10_idx)]';

imagesc(heatmap_data);
colormap(hot);
colorbar;
set(gca, 'XTick', 1:length(top_10_idx), ...
    'XTickLabel', perf_importance_table.Competency(top_10_idx), ...
    'XTickLabelRotation', 45, ...
    'YTick', 1:4, ...
    'YTickLabel', {'RF', 'ET', 'NN', 'GB'});
title('모델별 Feature Importance (상위 10개)', 'FontSize', 12);

% Subplot 3: 최종 통합 가중치
subplot(2, 3, 4:6);
top_15_final = 1:min(15, height(final_weight_table));
bar_data = [final_weight_table.Correlation_Weight(top_15_final), ...
           final_weight_table.ML_Weight(top_15_final)];

bar(bar_data);
set(gca, 'XTickLabel', final_weight_table.Competency(top_15_final), ...
    'XTickLabelRotation', 45);
ylabel('가중치 (%)');
title('최종 통합 가중치 (상위 15개)', 'FontSize', 12);
legend('상관분석', '머신러닝', 'Location', 'northeast');
grid on;

sgtitle('고성과자/저성과자 예측 모델 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');

%% ========================================================================
%                          PART 5: 종합 결과 및 보고서
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 5: 종합 결과 및 최종 보고서\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 5.1 결과 저장
fprintf('【STEP 15】 분석 결과 저장\n');
fprintf('────────────────────────────────────────────\n');

% MATLAB 파일 저장
analysis_results = struct();
analysis_results.profile_stats = profile_stats;
analysis_results.correlation_results = correlation_results;
analysis_results.performance_models = struct('rf', rf_perf_model, 'et', et_perf_model, 'gb', gb_perf_model);
analysis_results.performance_accuracy = model_performance;
analysis_results.feature_importance = perf_importance_table;
analysis_results.final_weights = final_weight_table;
analysis_results.config = config;

save('talent_analysis_enhanced_complete.mat', 'analysis_results');
fprintf('✓ MATLAB 파일 저장: talent_analysis_enhanced_complete.mat\n');

% Excel 파일 저장
try
    % Sheet 1: 인재유형 프로파일
    writetable(profile_stats, 'talent_analysis_enhanced_report.xlsx', ...
              'Sheet', '인재유형프로파일');

    % Sheet 2: 상관분석 결과
    writetable(correlation_results(1:min(20, height(correlation_results)), :), ...
              'talent_analysis_enhanced_report.xlsx', 'Sheet', '상관분석결과');

    % Sheet 3: 이진분류 모델 성능
    writetable(model_performance, 'talent_analysis_enhanced_report.xlsx', ...
              'Sheet', '이진분류성능');

    % Sheet 4: Feature Importance
    writetable(perf_importance_table(1:min(20, height(perf_importance_table)), :), ...
              'talent_analysis_enhanced_report.xlsx', 'Sheet', 'Feature중요도');

    % Sheet 5: 최종 통합 가중치
    writetable(final_weight_table, 'talent_analysis_enhanced_report.xlsx', ...
              'Sheet', '최종통합가중치');

    fprintf('✓ Excel 파일 저장: talent_analysis_enhanced_report.xlsx\n');
catch
    fprintf('⚠ Excel 저장 실패 (파일이 열려있을 수 있음)\n');
end

%% 5.2 최종 종합 보고서
fprintf('\n%s\n', repmat('═', 80, 1));
fprintf('                     최종 종합 분석 보고서 Enhanced\n');
fprintf('%s\n', repmat('═', 80, 1));

fprintf('\n📊 데이터 요약:\n');
fprintf('  • 분석 대상: %d명 (위장형 소화성 제외)\n', length(matched_ids));
fprintf('  • 인재유형: %d개\n', n_types);
fprintf('  • 역량항목: %d개\n', length(valid_comp_cols));

fprintf('\n🎯 인재유형별 주요 발견:\n');
[~, best_perf_idx] = max(profile_stats.PerformanceRank);
[~, most_common_idx] = max(profile_stats.Count);
fprintf('  1. 최고 성과 인재유형: %s (성과순위 %.0f)\n', ...
        profile_stats.TalentType{best_perf_idx}, profile_stats.PerformanceRank(best_perf_idx));
fprintf('  2. 최다 인원 인재유형: %s (%d명)\n', ...
        profile_stats.TalentType{most_common_idx}, profile_stats.Count(most_common_idx));

fprintf('\n🤖 이진 분류 모델 성능:\n');
for i = 1:height(model_performance)
    fprintf('  • %-18s: %5.1f%%\n', model_performance.Model{i}, model_performance.Accuracy(i) * 100);
end

fprintf('\n⭐ 핵심 예측 역량 Top 5 (최종 통합 가중치):\n');
for i = 1:min(5, height(final_weight_table))
    fprintf('  %d. %-15s: %5.1f%% (상관: %4.1f%%, ML: %4.1f%%)\n', i, ...
            final_weight_table.Competency{i}, ...
            final_weight_table.Final_Weight(i), ...
            final_weight_table.Correlation_Weight(i), ...
            final_weight_table.ML_Weight(i));
end

fprintf('\n📈 방법론별 기여도:\n');
avg_ml_accuracy = mean(model_performance.Accuracy);
if avg_ml_accuracy > 0.7
    fprintf('  • 머신러닝 모델: 우수 (평균 %.1f%%)\n', avg_ml_accuracy * 100);
    fprintf('  • 권장: 머신러닝 기반 실시간 평가 시스템 구축\n');
elseif avg_ml_accuracy > 0.6
    fprintf('  • 머신러닝 모델: 양호 (평균 %.1f%%)\n', avg_ml_accuracy * 100);
    fprintf('  • 권장: 상관분석과 머신러닝 조합 활용\n');
else
    fprintf('  • 머신러닝 모델: 보통 (평균 %.1f%%)\n', avg_ml_accuracy * 100);
    fprintf('  • 권장: 상관분석 결과 우선 활용\n');
end

fprintf('\n✨ 실무 적용 권장사항:\n');
fprintf('  1. 1차 스크리닝: 상위 3개 역량 (%s, %s, %s)\n', ...
        final_weight_table.Competency{1}, final_weight_table.Competency{2}, final_weight_table.Competency{3});
fprintf('  2. 정밀 평가: 상위 5개 역량 통합 점수 활용\n');
fprintf('  3. 모델 업데이트: 분기별 새로운 데이터로 재학습\n');
fprintf('  4. 예측 신뢰도: %.0f%% 이상일 때 높은 신뢰\n', avg_ml_accuracy * 100);

fprintf('\n📋 기술적 세부사항:\n');
fprintf('  • 데이터 전처리: 위장형 소화성 제외, 정규화 적용\n');
fprintf('  • 모델 구성: 4개 알고리즘 앙상블\n');
fprintf('  • 가중치 방법: 상관분석(50%%) + 머신러닝(50%%)\n');
fprintf('  • 검증 방법: 80/20 계층화 분할\n');

fprintf('\n%s\n', repmat('═', 80, 1));
fprintf('               Enhanced 분석 완료 - %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('%s\n', repmat('═', 80, 1));

fprintf('\n📈 시각화 완료: 모델 성능, Feature Importance, 통합 가중치 차트 생성\n');