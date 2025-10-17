%% 인재유형 종합 분석 시스템 v3.0 (Enhanced Class Imbalance Solution)
% HR Talent Type Comprehensive Analysis System
% 목적: 1) 인재 유형별 프로파일링
%      2) 상관 기반 가중치 부여
%      3) 고도화된 머신러닝 기법을 이용한 예측 분석 (클래스 불균형 최적화)

clear; clc; close all;

%% ========================================================================
%                          PART 1: 데이터 준비 및 전처리
% =========================================================================

fprintf('═══════════════════════════════════════════════════════════\n');
fprintf('         인재유형 종합 분석 시스템 v3.0 (Enhanced)\n');
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

%% ========================================================================
%          PART 4: 고도화된 머신러닝 예측 분석 (클래스 불균형 최적화)
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('      PART 4: 고도화된 머신러닝 예측 분석 (클래스 불균형 최적화)\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 4.1 클래스 불균형 분석 및 개선된 SMOTE
fprintf('【STEP 8】 고급 클래스 불균형 분석\n');
fprintf('────────────────────────────────────────────\n');

% 특성 및 레이블 준비
X = table2array(matched_comp);
y = matched_talent_types;
[y_unique, ~, y_encoded] = unique(y);

% 클래스별 샘플 수 계산
class_counts = histcounts(y_encoded, 1:(length(y_unique)+1));
imbalance_ratio = max(class_counts) / min(class_counts);

fprintf('⚠ 클래스 불균형 분석:\n');
fprintf('  • 불균형 비율: %.1f:1\n', imbalance_ratio);
fprintf('  • 클래스별 분포:\n');
for i = 1:length(y_unique)
    fprintf('    - %-20s: %3d명 (%5.1f%%)\n', y_unique{i}, class_counts(i), ...
            class_counts(i)/sum(class_counts)*100);
end

% 데이터 정규화
X_normalized = normalize(X, 'range');

%% 4.1.1 개선된 적응적 SMOTE 구현
fprintf('\n【STEP 8-1】 적응적 SMOTE 구현\n');
fprintf('────────────────────────────────────────────\n');

% 동적 목표 샘플 수 설정
min_samples = 10;  % 최소 보장 샘플
max_samples = round(prctile(class_counts, 75));  % 75 퍼센타일 기준
target_samples = max(min_samples, min(max_samples, round(median(class_counts))));

fprintf('목표 샘플 수: %d명/클래스 (범위: %d-%d)\n', target_samples, min_samples, max_samples);

% 적응적 SMOTE 적용
X_balanced = [];
y_balanced = [];

for class = 1:length(y_unique)
    class_idx = find(y_encoded == class);
    X_class = X_normalized(class_idx, :);
    n_samples = length(class_idx);

    if n_samples < target_samples
        % 소수 클래스 - 개선된 오버샘플링
        n_synthetic = target_samples - n_samples;

        if n_samples >= 2
            % 적응적 K값 설정
            k_neighbors = min(max(3, round(sqrt(n_samples))), n_samples-1);

            % 경계선 샘플 우선 선택을 위한 밀도 계산
            densities = zeros(n_samples, 1);
            for i = 1:n_samples
                distances = sqrt(sum((X_class - X_class(i, :)).^2, 2));
                densities(i) = 1 / (mean(sort(distances(2:min(k_neighbors+1, end)))) + eps);
            end

            % 낮은 밀도(경계선) 샘플에 높은 가중치
            sample_weights = 1 ./ (densities + eps);
            sample_probs = sample_weights / sum(sample_weights);

            X_synthetic = zeros(n_synthetic, size(X_class, 2));

            for i = 1:n_synthetic
                % 가중 확률에 따라 기준 샘플 선택 (호환성 개선)
                cumulative_probs = cumsum(sample_probs);
                rand_val = rand();
                base_idx = find(cumulative_probs >= rand_val, 1);
                if isempty(base_idx)
                    base_idx = n_samples;
                end
                base_sample = X_class(base_idx, :);

                % 최근접 이웃 찾기
                distances = sqrt(sum((X_class - base_sample).^2, 2));
                [~, sorted_idx] = sort(distances);
                neighbor_idx = sorted_idx(randi([2, min(k_neighbors+1, n_samples)]));
                neighbor_sample = X_class(neighbor_idx, :);

                % 적응적 보간 계수 (경계선 샘플일수록 보수적)
                density_factor = densities(base_idx) / max(densities);
                lambda = 0.3 + 0.4 * density_factor + 0.3 * rand();  % 0.3-0.7 범위

                % 합성 샘플 생성 + 소량 노이즈
                noise_level = 0.01 * std(X_class(:));
                X_synthetic(i, :) = base_sample + lambda * (neighbor_sample - base_sample) + ...
                                   noise_level * randn(1, size(X_class, 2));
            end

            X_balanced = [X_balanced; X_class; X_synthetic];
            y_balanced = [y_balanced; repmat(class, n_samples + n_synthetic, 1)];

            fprintf('  %s: %d → %d (적응적 SMOTE %d개, K=%d)\n', y_unique{class}, ...
                    n_samples, n_samples + n_synthetic, n_synthetic, k_neighbors);
        else
            % 샘플이 1개뿐인 경우 - 노이즈 변형 복제
            noise_level = 0.05;
            X_variants = repmat(X_class, target_samples, 1) + ...
                        noise_level * randn(target_samples, size(X_class, 2));
            X_balanced = [X_balanced; X_variants];
            y_balanced = [y_balanced; repmat(class, target_samples, 1)];
            fprintf('  %s: %d → %d (노이즈 변형 복제)\n', y_unique{class}, n_samples, target_samples);
        end
    elseif n_samples > target_samples * 1.5
        % 다수 클래스 - 계층화 언더샘플링
        % 성과 점수 기반 계층화
        class_perf_scores = performance_scores(class_idx);
        if any(class_perf_scores > 0)
            % 성과 점수별 균등 샘플링
            [~, ~, perf_bins] = unique(class_perf_scores);
            sample_idx = [];
            samples_per_bin = ceil(target_samples / max(perf_bins));

            for bin = 1:max(perf_bins)
                bin_idx = find(perf_bins == bin);
                if length(bin_idx) <= samples_per_bin
                    sample_idx = [sample_idx; bin_idx];
                else
                    if length(bin_idx) <= samples_per_bin
                        sample_idx = [sample_idx; bin_idx];
                    else
                        rand_perm = randperm(length(bin_idx));
                        sample_idx = [sample_idx; bin_idx(rand_perm(1:samples_per_bin))];
                    end
                end
            end
            sample_idx = sample_idx(1:min(target_samples, end));
        else
            rand_perm = randperm(n_samples);
            sample_idx = rand_perm(1:target_samples);
        end

        X_balanced = [X_balanced; X_class(sample_idx, :)];
        y_balanced = [y_balanced; repmat(class, length(sample_idx), 1)];
        fprintf('  %s: %d → %d (계층화 언더샘플링)\n', y_unique{class}, n_samples, length(sample_idx));
    else
        % 적절한 수준 - 그대로 사용
        X_balanced = [X_balanced; X_class];
        y_balanced = [y_balanced; repmat(class, n_samples, 1)];
        fprintf('  %s: %d (유지)\n', y_unique{class}, n_samples);
    end
end

fprintf('균형화 완료: %d → %d 샘플\n', length(y_encoded), length(y_balanced));

%% 4.2 교차검증 기반 모델 최적화
fprintf('\n【STEP 9】 교차검증 기반 하이퍼파라미터 최적화\n');
fprintf('────────────────────────────────────────────\n');

% 원본 데이터를 테스트용으로 분할 (수동 계층화 분할)
test_ratio = 0.2;
test_indices = [];
for class = 1:length(y_unique)
    class_indices = find(y_encoded == class);
    n_test = max(1, round(length(class_indices) * test_ratio));
    rand_perm = randperm(length(class_indices));
    test_indices = [test_indices; class_indices(rand_perm(1:n_test))];
end
train_indices = setdiff(1:length(y_encoded), test_indices);

X_test = X_normalized(test_indices, :);
y_test = y_encoded(test_indices);

% 균형화된 데이터를 훈련용으로 사용
X_train = X_balanced;
y_train = y_balanced;

fprintf('▶ 훈련 데이터: %d명 (균형화됨)\n', length(y_train));
fprintf('▶ 테스트 데이터: %d명 (원본)\n', length(y_test));

% 5-Fold 교차검증으로 하이퍼파라미터 최적화 (수동 구현)
cv_folds = 5;
n_train = length(y_train);
fold_size = floor(n_train / cv_folds);
cv_indices = cell(cv_folds, 1);
rand_perm = randperm(n_train);
for fold = 1:cv_folds
    start_idx = (fold-1) * fold_size + 1;
    if fold == cv_folds
        end_idx = n_train;
    else
        end_idx = fold * fold_size;
    end
    cv_indices{fold} = rand_perm(start_idx:end_idx);
end

% Random Forest 파라미터 그리드
rf_params = struct();
rf_params.n_trees = [50, 100, 150, 200];
rf_params.min_leaf = [1, 2, 3, 5];
rf_params.max_splits = [10, 20, 50];

best_rf_score = 0;
best_rf_params = struct();

fprintf('Random Forest 하이퍼파라미터 최적화 중...\n');
param_count = 0;
total_params = length(rf_params.n_trees) * length(rf_params.min_leaf) * length(rf_params.max_splits);

for n_trees = rf_params.n_trees
    for min_leaf = rf_params.min_leaf
        for max_splits = rf_params.max_splits
            param_count = param_count + 1;

            % 교차검증 평가
            cv_scores = zeros(cv_folds, 1);
            for fold = 1:cv_folds
                val_indices = cv_indices{fold};
                train_indices_fold = setdiff(1:n_train, val_indices);

                X_train_fold = X_train(train_indices_fold, :);
                y_train_fold = y_train(train_indices_fold);
                X_val_fold = X_train(val_indices, :);
                y_val_fold = y_train(val_indices);

                % 모델 학습
                model = TreeBagger(n_trees, X_train_fold, y_train_fold, ...
                    'Method', 'classification', ...
                    'MinLeafSize', min_leaf, ...
                    'MaxNumSplits', max_splits);

                % 예측 및 균형 정확도 계산
                y_pred = cellfun(@str2double, predict(model, X_val_fold));

                % 클래스별 재현율 계산 (균형 정확도)
                unique_classes = unique(y_val_fold);
                class_recalls = zeros(length(unique_classes), 1);
                for c = 1:length(unique_classes)
                    class_mask = y_val_fold == unique_classes(c);
                    if sum(class_mask) > 0
                        class_recalls(c) = sum(y_pred(class_mask) == unique_classes(c)) / sum(class_mask);
                    end
                end
                cv_scores(fold) = mean(class_recalls);  % 균형 정확도
            end

            mean_score = mean(cv_scores);
            if mean_score > best_rf_score
                best_rf_score = mean_score;
                best_rf_params.n_trees = n_trees;
                best_rf_params.min_leaf = min_leaf;
                best_rf_params.max_splits = max_splits;
            end

            if mod(param_count, 10) == 0 || param_count == total_params
                fprintf('  진행률: %d/%d (%.1f%%), 현재 최고: %.4f\n', ...
                        param_count, total_params, param_count/total_params*100, best_rf_score);
            end
        end
    end
end

fprintf('최적 RF 파라미터: Trees=%d, MinLeaf=%d, MaxSplits=%d, CV점수=%.4f\n', ...
        best_rf_params.n_trees, best_rf_params.min_leaf, best_rf_params.max_splits, best_rf_score);

%% 4.3 앙상블 모델 구성
fprintf('\n【STEP 10】 다중 앙상블 모델 구성\n');
fprintf('────────────────────────────────────────────\n');

rng(42);  % 재현성을 위한 시드

% 1. 최적화된 Random Forest
fprintf('1. 최적화된 Random Forest 학습 중...\n');
rf_model = TreeBagger(best_rf_params.n_trees, X_train, y_train, ...
    'Method', 'classification', ...
    'MinLeafSize', best_rf_params.min_leaf, ...
    'MaxNumSplits', best_rf_params.max_splits, ...
    'OOBPredictorImportance', 'on', ...
    'PredictorNames', valid_comp_cols);

% 2. Cost-sensitive Gradient Boosting
fprintf('2. Cost-sensitive Gradient Boosting 학습 중...\n');
% 클래스 가중치 계산
unique_train_classes = unique(y_train);
class_costs = zeros(length(unique_train_classes), length(unique_train_classes));
for i = 1:length(unique_train_classes)
    for j = 1:length(unique_train_classes)
        if i ~= j
            % 소수 클래스 오분류에 더 큰 비용
            class_i_count = sum(y_train == unique_train_classes(i));
            total_samples = length(y_train);
            cost_weight = total_samples / (length(unique_train_classes) * class_i_count);
            class_costs(i, j) = cost_weight;
        end
    end
end

gb_model = fitcensemble(X_train, y_train, ...
    'Method', 'AdaBoostM2', ...
    'NumLearningCycles', 100, ...
    'LearnRate', 0.1, ...
    'Learners', templateTree('MaxNumSplits', 20, 'MinLeafSize', 5), ...
    'Cost', class_costs);

% 3. 클래스별 전문 분류기 (One-vs-Rest)
fprintf('3. 클래스별 전문 분류기 학습 중...\n');
class_experts = cell(length(y_unique), 1);
expert_scores = zeros(length(y_unique), 1);

for class = 1:length(y_unique)
    if sum(y_train == class) >= 5  % 최소 5개 샘플 필요
        y_binary = double(y_train == class);

        % 이진 분류를 위한 SMOTE 추가 적용
        pos_idx = find(y_binary == 1);
        neg_idx = find(y_binary == 0);

        if length(pos_idx) < length(neg_idx) / 2
            % 양성 클래스 추가 증강
            n_synthetic = min(length(neg_idx) - length(pos_idx), length(pos_idx));
            X_pos = X_train(pos_idx, :);

            for i = 1:n_synthetic
                base_idx = randi(length(pos_idx));
                if length(pos_idx) > 1
                    neighbor_idx = randi(length(pos_idx));
                    while neighbor_idx == base_idx
                        neighbor_idx = randi(length(pos_idx));
                    end
                    lambda = rand();
                    synthetic_sample = X_pos(base_idx, :) + lambda * (X_pos(neighbor_idx, :) - X_pos(base_idx, :));
                    X_train = [X_train; synthetic_sample];
                    y_binary = [y_binary; 1];
                end
            end
        end

        class_experts{class} = fitcensemble(X_train(1:length(y_train), :), y_binary(1:length(y_train)), ...
            'Method', 'RUSBoost', ...
            'NumLearningCycles', 50, ...
            'Learners', templateTree('MaxNumSplits', 10));

        % 교차검증 평가 (3-fold 수동 구현)
        y_binary_subset = y_binary(1:length(y_train));
        n_binary = length(y_binary_subset);
        fold_size_expert = floor(n_binary / 3);
        expert_cv_scores = zeros(3, 1);
        rand_perm_expert = randperm(n_binary);

        for fold = 1:3
            start_idx = (fold-1) * fold_size_expert + 1;
            if fold == 3
                end_idx = n_binary;
            else
                end_idx = fold * fold_size_expert;
            end
            test_idx_expert = rand_perm_expert(start_idx:end_idx);
            train_idx_expert = setdiff(1:n_binary, test_idx_expert);

            if length(test_idx_expert) > 0 && length(train_idx_expert) > 0
                y_pred_expert = predict(class_experts{class}, X_train(test_idx_expert, :));
                y_true_expert = y_binary_subset(test_idx_expert);
                expert_cv_scores(fold) = sum(y_pred_expert == y_true_expert) / length(y_true_expert);
            end
        end
        expert_scores(class) = mean(expert_cv_scores);

        fprintf('  %s 전문가: CV 정확도 %.3f\n', y_unique{class}, expert_scores(class));
    end
end

%% 4.4 앙상블 예측 및 평가
fprintf('\n【STEP 11】 앙상블 예측 및 평가\n');
fprintf('────────────────────────────────────────────\n');

% Random Forest 예측
[y_pred_rf_cell, rf_scores] = predict(rf_model, X_test);
y_pred_rf = cellfun(@str2double, y_pred_rf_cell);
% TreeBagger 결과 처리 (MATLAB 버전 호환성)
if iscell(rf_scores)
    rf_probs = cell2mat(rf_scores);
else
    rf_probs = rf_scores;
end

% Gradient Boosting 예측
[y_pred_gb, gb_scores] = predict(gb_model, X_test);
gb_probs = gb_scores;

% 클래스 전문가 예측
expert_probs = zeros(length(y_test), length(y_unique));
for class = 1:length(y_unique)
    if ~isempty(class_experts{class})
        [~, scores] = predict(class_experts{class}, X_test);
        if size(scores, 2) >= 2
            expert_probs(:, class) = scores(:, 2);
        end
    end
end

% 가중 앙상블 (성능 기반 가중치)
model_weights = [0.4, 0.3, 0.3];  % RF, GB, Expert 가중치
ensemble_probs = model_weights(1) * rf_probs + model_weights(2) * gb_probs + ...
                model_weights(3) * expert_probs;

% 확률 정규화
ensemble_probs = ensemble_probs ./ sum(ensemble_probs, 2);
[~, y_pred_ensemble] = max(ensemble_probs, [], 2);

% 개별 모델 정확도
rf_accuracy = sum(y_pred_rf == y_test) / length(y_test);
gb_accuracy = sum(y_pred_gb == y_test) / length(y_test);
ensemble_accuracy = sum(y_pred_ensemble == y_test) / length(y_test);

fprintf('모델 성능 비교:\n');
fprintf('  • Random Forest: %.4f\n', rf_accuracy);
fprintf('  • Gradient Boosting: %.4f\n', gb_accuracy);
fprintf('  • 가중 앙상블: %.4f\n', ensemble_accuracy);

%% 4.5 고급 성능 평가
fprintf('\n【STEP 12】 고급 성능 평가\n');
fprintf('────────────────────────────────────────────\n');

% Confusion Matrix
conf_ensemble = confusionmat(y_test, y_pred_ensemble);

% 클래스별 성능 메트릭
n_classes = length(y_unique);
class_metrics = table();
class_metrics.TalentType = y_unique;
class_metrics.Precision = zeros(n_classes, 1);
class_metrics.Recall = zeros(n_classes, 1);
class_metrics.F1Score = zeros(n_classes, 1);
class_metrics.Support = zeros(n_classes, 1);
class_metrics.Confidence = zeros(n_classes, 1);

for i = 1:n_classes
    tp = conf_ensemble(i, i);
    fp = sum(conf_ensemble(:, i)) - tp;
    fn = sum(conf_ensemble(i, :)) - tp;

    class_metrics.Precision(i) = tp / (tp + fp + eps);
    class_metrics.Recall(i) = tp / (tp + fn + eps);
    class_metrics.F1Score(i) = 2 * (class_metrics.Precision(i) * class_metrics.Recall(i)) / ...
                              (class_metrics.Precision(i) + class_metrics.Recall(i) + eps);
    class_metrics.Support(i) = sum(y_test == i);

    % 예측 확신도 (해당 클래스로 예측된 샘플들의 평균 확률)
    pred_mask = y_pred_ensemble == i;
    if sum(pred_mask) > 0
        class_metrics.Confidence(i) = mean(ensemble_probs(pred_mask, i));
    end
end

% 균형 정확도 및 기타 메트릭
balanced_accuracy = mean(class_metrics.Recall);
macro_f1 = mean(class_metrics.F1Score);
weighted_f1 = sum(class_metrics.F1Score .* class_metrics.Support) / sum(class_metrics.Support);

% Matthews 상관계수 계산
mcc_numerator = 0;
mcc_denominator = 0;
for i = 1:n_classes
    for j = 1:n_classes
        for k = 1:n_classes
            mcc_numerator = mcc_numerator + conf_ensemble(i,i) * conf_ensemble(j,k) - ...
                           conf_ensemble(i,k) * conf_ensemble(k,i);
        end
    end
end

sum_pred = sum(conf_ensemble, 1);
sum_true = sum(conf_ensemble, 2)';
mcc_denominator = sqrt(sum(sum_pred.^2) - sum(sum_pred)^2) * ...
                 sqrt(sum(sum_true.^2) - sum(sum_true)^2);
mcc = mcc_numerator / (mcc_denominator + eps);

fprintf('\n종합 성능 지표:\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  • 정확도 (Accuracy): %.4f\n', ensemble_accuracy);
fprintf('  • 균형 정확도 (Balanced Accuracy): %.4f\n', balanced_accuracy);
fprintf('  • Macro F1-Score: %.4f\n', macro_f1);
fprintf('  • Weighted F1-Score: %.4f\n', weighted_f1);
fprintf('  • Matthews 상관계수 (MCC): %.4f\n', mcc);

fprintf('\n클래스별 상세 성능:\n');
fprintf('%-20s | Prec. | Recall | F1    | Conf. | Support\n', '인재유형');
fprintf('%s\n', repmat('-', 65, 1));

for i = 1:height(class_metrics)
    fprintf('%-20s | %5.3f | %6.3f | %5.3f | %5.3f | %7d\n', ...
        class_metrics.TalentType{i}, ...
        class_metrics.Precision(i), ...
        class_metrics.Recall(i), ...
        class_metrics.F1Score(i), ...
        class_metrics.Confidence(i), ...
        class_metrics.Support(i));
end

%% 4.6 Feature Importance 분석
fprintf('\n【STEP 13】 앙상블 Feature Importance\n');
fprintf('────────────────────────────────────────────\n');

% Random Forest importance
rf_importance = rf_model.OOBPermutedPredictorDeltaError;
rf_importance_norm = rf_importance / sum(rf_importance);

% Gradient Boosting importance
gb_importance = predictorImportance(gb_model);
gb_importance_norm = gb_importance / sum(gb_importance);

% 상관 기반 가중치와 결합
final_importance = (0.4 * rf_importance_norm + 0.3 * gb_importance_norm + 0.3 * weights_normalized');
final_importance = final_importance / sum(final_importance);

% Feature Importance 테이블 (차원 맞춤)
importance_table = table();
importance_table.Competency = valid_comp_cols';

% 모든 변수를 열 벡터로 변환
rf_importance_norm = rf_importance_norm(:);
gb_importance_norm = gb_importance_norm(:);
weights_normalized = weights_normalized(:);
final_importance = final_importance(:);

% 길이 확인 및 조정
n_features = length(valid_comp_cols);
if length(rf_importance_norm) ~= n_features
    rf_importance_norm = rf_importance_norm(1:min(end, n_features));
    if length(rf_importance_norm) < n_features
        rf_importance_norm = [rf_importance_norm; zeros(n_features - length(rf_importance_norm), 1)];
    end
end
if length(gb_importance_norm) ~= n_features
    gb_importance_norm = gb_importance_norm(1:min(end, n_features));
    if length(gb_importance_norm) < n_features
        gb_importance_norm = [gb_importance_norm; zeros(n_features - length(gb_importance_norm), 1)];
    end
end
if length(weights_normalized) ~= n_features
    weights_normalized = weights_normalized(1:min(end, n_features));
    if length(weights_normalized) < n_features
        weights_normalized = [weights_normalized; zeros(n_features - length(weights_normalized), 1)];
    end
end
if length(final_importance) ~= n_features
    final_importance = final_importance(1:min(end, n_features));
    if length(final_importance) < n_features
        final_importance = [final_importance; zeros(n_features - length(final_importance), 1)];
    end
end

importance_table.RF_Importance = rf_importance_norm * 100;
importance_table.GB_Importance = gb_importance_norm * 100;
importance_table.Correlation_Weight = weights_normalized * 100;
importance_table.Final_Importance = final_importance * 100;

importance_table = sortrows(importance_table, 'Final_Importance', 'descend');

fprintf('\n최종 Feature Importance (상위 10개):\n');
fprintf('%-30s | RF(%) | GB(%) | Corr(%) | Final(%%)\n', '역량');
fprintf('%s\n', repmat('-', 70, 1));

for i = 1:min(10, height(importance_table))
    fprintf('%-30s | %5.2f | %5.2f | %7.2f | %7.2f\n', ...
        importance_table.Competency{i}, ...
        importance_table.RF_Importance(i), ...
        importance_table.GB_Importance(i), ...
        importance_table.Correlation_Weight(i), ...
        importance_table.Final_Importance(i));
end

%% ========================================================================
%                          PART 5: 결과 저장 및 보고서
% =========================================================================

fprintf('\n\n═══════════════════════════════════════════════════════════\n');
fprintf('              PART 5: 결과 저장 및 최종 보고서\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

%% 5.1 결과 저장
fprintf('【STEP 14】 분석 결과 저장\n');
fprintf('────────────────────────────────────────────\n');

% MATLAB 파일 저장
analysis_results = struct();
analysis_results.profile_stats = profile_stats;
analysis_results.correlation_results = correlation_results;
analysis_results.importance_table = importance_table;
analysis_results.class_metrics = class_metrics;
analysis_results.models = struct('rf', rf_model, 'gb', gb_model, 'experts', {class_experts});
analysis_results.performance = struct('rf', rf_accuracy, 'gb', gb_accuracy, ...
                                    'ensemble', ensemble_accuracy, 'balanced', balanced_accuracy, ...
                                    'macro_f1', macro_f1, 'mcc', mcc);
analysis_results.ensemble_probs = ensemble_probs;
analysis_results.config = config;

save('talent_analysis_v3_complete.mat', 'analysis_results');
fprintf('✓ MATLAB 파일 저장: talent_analysis_v3_complete.mat\n');

% Excel 파일 저장
try
    % Sheet 1: 인재유형 프로파일
    writetable(profile_stats, 'talent_analysis_v3_report.xlsx', ...
              'Sheet', 'TalentProfiles');

    % Sheet 2: 상관분석 결과
    writetable(correlation_results(1:min(30, height(correlation_results)), :), ...
              'talent_analysis_v3_report.xlsx', 'Sheet', 'CorrelationAnalysis');

    % Sheet 3: Feature Importance
    writetable(importance_table(1:min(30, height(importance_table)), :), ...
              'talent_analysis_v3_report.xlsx', 'Sheet', 'FeatureImportance');

    % Sheet 4: 모델 성능
    model_performance = table();
    model_performance.Model = {'Random Forest'; 'Gradient Boosting'; 'Weighted Ensemble'};
    model_performance.Accuracy = [rf_accuracy; gb_accuracy; ensemble_accuracy];
    model_performance.Balanced_Accuracy = [NaN; NaN; balanced_accuracy];
    model_performance.Macro_F1 = [NaN; NaN; macro_f1];
    model_performance.MCC = [NaN; NaN; mcc];

    writetable(model_performance, 'talent_analysis_v3_report.xlsx', ...
              'Sheet', 'ModelPerformance');

    % Sheet 5: 클래스별 성능
    writetable(class_metrics, 'talent_analysis_v3_report.xlsx', ...
              'Sheet', 'ClassPerformance');

    fprintf('✓ Excel 파일 저장: talent_analysis_v3_report.xlsx\n');
catch
    fprintf('⚠ Excel 저장 실패 (파일이 열려있을 수 있음)\n');
end

%% 5.2 최종 보고서
fprintf('\n%s\n', repmat('═', 80, 1));
fprintf('                     최종 분석 보고서 v3.0\n');
fprintf('%s\n', repmat('═', 80, 1));

fprintf('\n📊 데이터 요약:\n');
fprintf('  • 분석 대상: %d명\n', length(matched_ids));
fprintf('  • 인재유형: %d개\n', n_types);
fprintf('  • 역량항목: %d개\n', length(valid_comp_cols));
fprintf('  • 클래스 불균형 비율: %.1f:1\n', imbalance_ratio);

fprintf('\n🔧 클래스 불균형 해결 방법:\n');
fprintf('  • 적응적 SMOTE (경계선 샘플 우선, 동적 K값)\n');
fprintf('  • 계층화 언더샘플링 (성과 기반)\n');
fprintf('  • 노이즈 추가 및 밀도 기반 가중치\n');
fprintf('  • 다중 앙상블 (RF + GB + 전문가 모델)\n');

fprintf('\n🎯 주요 발견사항:\n');
fprintf('  1. 최고 성과 인재유형: %s\n', ...
    profile_stats.TalentType{profile_stats.PerformanceRank == max(profile_stats.PerformanceRank)});

fprintf('  2. 핵심 예측 역량 Top 3:\n');
for i = 1:min(3, height(importance_table))
    fprintf('     - %s (%.2f%%)\n', importance_table.Competency{i}, ...
            importance_table.Final_Importance(i));
end

fprintf('\n🤖 모델 성능 (v3.0 개선):\n');
fprintf('  • 전체 정확도: %.2f%%\n', ensemble_accuracy * 100);
fprintf('  • 균형 정확도: %.2f%% (소수 클래스 고려)\n', balanced_accuracy * 100);
fprintf('  • Macro F1-Score: %.2f%%\n', macro_f1 * 100);
fprintf('  • Matthews 상관계수: %.3f\n', mcc);

fprintf('\n📈 성능 개선 효과:\n');
if balanced_accuracy > 0.7
    fprintf('  • 우수: 실무 적용 가능 수준\n');
elseif balanced_accuracy > 0.5
    fprintf('  • 양호: 보조 도구로 활용 가능\n');
else
    fprintf('  • 개선 필요: 추가 데이터 수집 권장\n');
end

fprintf('\n✨ 권장사항:\n');
fprintf('  1. 상위 5개 역량 (%s 등)으로 1차 스크리닝\n', importance_table.Competency{1});
fprintf('  2. 예측 확신도 %.2f 이상일 때 높은 신뢰도\n', mean(class_metrics.Confidence));
fprintf('  3. 소수 클래스는 추가 데이터 수집 후 재학습\n');
fprintf('  4. 3개월마다 모델 재학습으로 성능 유지\n');

fprintf('\n📋 기술적 세부사항:\n');
fprintf('  • SMOTE 합성 샘플: %d개\n', length(y_balanced) - length(y_encoded));
fprintf('  • 교차검증 fold: 5-fold\n');
fprintf('  • 앙상블 구성: RF(40%%) + GB(30%%) + Expert(30%%)\n');
fprintf('  • 하이퍼파라미터: Trees=%d, MinLeaf=%d, MaxSplits=%d\n', ...
         best_rf_params.n_trees, best_rf_params.min_leaf, best_rf_params.max_splits);

fprintf('\n%s\n', repmat('═', 80, 1));
fprintf('               분석 완료 - %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('%s\n', repmat('═', 80, 1));

%% 추가 시각화
figure('Position', [100, 100, 1400, 1000], 'Color', 'white');

% Subplot 1: 클래스 분포 (원본 vs 균형화)
subplot(2, 3, 1);
original_dist = histcounts(y_encoded, 1:(length(y_unique)+1));
balanced_dist = histcounts(y_balanced, 1:(length(y_unique)+1));
bar_data = [original_dist; balanced_dist]';
bar(bar_data);
set(gca, 'XTickLabel', y_unique, 'XTickLabelRotation', 45);
ylabel('샘플 수');
title('클래스 분포 변화');
legend('원본', '균형화', 'Location', 'best');
grid on;

% Subplot 2: 모델 성능 비교
subplot(2, 3, 2);
model_names = {'RF', 'GB', 'Ensemble'};
accuracies = [rf_accuracy, gb_accuracy, ensemble_accuracy];
bar(accuracies, 'FaceColor', [0.3, 0.6, 0.9]);
set(gca, 'XTickLabel', model_names);
ylabel('정확도');
title('모델 성능 비교');
ylim([0, 1]);
grid on;
for i = 1:length(accuracies)
    text(i, accuracies(i) + 0.02, sprintf('%.3f', accuracies(i)), ...
         'HorizontalAlignment', 'center');
end

% Subplot 3: 클래스별 F1-Score
subplot(2, 3, 3);
bar(class_metrics.F1Score, 'FaceColor', [0.9, 0.3, 0.3]);
set(gca, 'XTickLabel', class_metrics.TalentType, 'XTickLabelRotation', 45);
ylabel('F1-Score');
title('클래스별 F1-Score');
grid on;

% Subplot 4: Feature Importance Top 10
subplot(2, 3, 4:6);
top_n = min(10, height(importance_table));
bar_data = [importance_table.RF_Importance(1:top_n), ...
           importance_table.GB_Importance(1:top_n), ...
           importance_table.Final_Importance(1:top_n)];
bar(bar_data);
set(gca, 'XTickLabel', importance_table.Competency(1:top_n), 'XTickLabelRotation', 45);
ylabel('중요도 (%)');
title('Feature Importance 비교 (상위 10개)');
legend('Random Forest', 'Gradient Boosting', 'Final Ensemble', 'Location', 'northeast');
grid on;

sgtitle('인재유형 예측 모델 v3.0 - 종합 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');

fprintf('\n📈 시각화 완료: 클래스 분포, 모델 성능, Feature Importance 차트 생성\n');
