%% HR 데이터 분석 시스템 - 메인 스크립트
% =========================================================================
% 주요 기능:
% 1. 인재유형별 역량 프로파일 분석
% 2. 고성과자 예측 모델 개발 (Cost-Sensitive Learning)
% 3. 다양한 가중치 방법론 비교 (상관, 로지스틱, Bootstrap, T-test, 앙상블)
% 4. 통계적 유의성 검증 (퍼뮤테이션 테스트)
% 5. 성별 효과 및 인구통계 변수 분석
% =========================================================================

clear; clc; close all;
rng(42);

%% ========================================================================
%                          메인 실행 부분 (간결하게 정리)
% =========================================================================


try
    % 1. 초기 설정
    fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
    fprintf('║            HR 데이터 분석 시스템 시작                      ║\n');
    fprintf('╚════════════════════════════════════════════════════════════╝\n\n');
    
    config = initializeConfig();
    setupGraphicsDefaults();
    
    % 2. 데이터 로딩 및 전처리
    fprintf('\n【Phase 1】 데이터 로딩 및 전처리\n');
    fprintf('════════════════════════════════════════════\n');
    
    [hr_data, comp_data, comp_total] = loadData(config);
    comp_data = applyDataFilters(comp_data, config);
    
    % 3. 데이터 통합 및 매칭
    [matched_data, talent_info] = integrateAndMatchData(hr_data, comp_data, comp_total, config);
    
    % 4. 이상치 제거 (선택적)
    if config.outlier_removal.enabled
        matched_data = removeOutliers(matched_data, config);
    end
    
    % 5. 탐색적 데이터 분석
    fprintf('\n【Phase 2】 탐색적 데이터 분석\n');
    fprintf('════════════════════════════════════════════\n');
    
    eda_results = performEDA(matched_data, talent_info, config);
    
    % 6. 유형별 프로파일 분석 및 시각화
    fprintf('\n【Phase 3】 인재유형 프로파일 분석\n');
    fprintf('════════════════════════════════════════════\n');
    
    profile_results = analyzeTypeProfiles(matched_data, talent_info, config);
    createRadarCharts(profile_results, config);
    
    % 7. 예측 모델링
    fprintf('\n【Phase 4】 고성과자 예측 모델링\n');
    fprintf('════════════════════════════════════════════\n');
    
    % 데이터 준비
    [X, y, feature_names] = preparePredictionData(matched_data, talent_info, config);
    
    % 다양한 가중치 방법 적용
    weight_results = calculateAllWeights(X, y, feature_names, config);
    
    % 모델 학습 및 평가
    model_results = trainAndEvaluateModels(X, y, weight_results, config);
    
    % 8. 통계적 검증
    fprintf('\n【Phase 5】 통계적 유의성 검증\n');
    fprintf('════════════════════════════════════════════\n');
    
    statistical_results = performStatisticalTests(X, y, matched_data, talent_info, config);
    
    % Bootstrap 안정성 검증
    if config.run_bootstrap
        bootstrap_results = performBootstrapValidation(X, y, feature_names, config);
    else
        bootstrap_results = [];
    end
    
    % 퍼뮤테이션 테스트
    if config.run_permutation
        permutation_results = performPermutationTests(X, y, model_results, config);
    else
        permutation_results = [];
    end
    
    % 9. 성별 및 인구통계 효과 분석
    fprintf('\n【Phase 6】 인구통계 변수 효과 분석\n');
    fprintf('════════════════════════════════════════════\n');
    
    demo_results = analyzeDemographicEffects(matched_data, y, config);
    
    % 10. 결과 통합 및 시각화
    fprintf('\n【Phase 7】 결과 통합 및 시각화\n');
    fprintf('════════════════════════════════════════════\n');
    
    all_results = consolidateResults(profile_results, weight_results, model_results, ...
                                    statistical_results, bootstrap_results, ...
                                    permutation_results, demo_results);
    
    createFinalVisualizations(all_results, config);
    
    % 11. 결과 저장
    fprintf('\n【Phase 8】 결과 저장\n');
    fprintf('════════════════════════════════════════════\n');
    
    saveResults(all_results, config);
    exportToExcel(all_results, config);
    
    % 12. 최종 요약
    printFinalSummary(all_results, config);
    
    fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
    fprintf('║            ✅ HR 데이터 분석 완료!                         ║\n');
    fprintf('╚════════════════════════════════════════════════════════════╝\n\n');
    
catch ME
    fprintf('\n❌ 오류 발생: %s\n', ME.message);
    fprintf('   위치: %s (줄 %d)\n', ME.stack(1).file, ME.stack(1).line);
    rethrow(ME);
end

%% ========================================================================
%                          함수 정의 부분
% =========================================================================

%% === 초기화 및 설정 함수들 ===

function config = initializeConfig()
    % 전체 시스템 설정 구조체 생성
    
    config = struct();
    
    % 파일 경로
    config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
    config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
    config.output_dir = 'D:\project\HR데이터\결과\자가불소_simple_mean';
    
    % 타임스탬프
    config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
    
    % 모델 설정
    config.model_file = fullfile(config.output_dir, 'trained_logitboost_model.mat');
    config.use_saved_model = true;
    config.baseline_type = 'weighted';  % 'simple' or 'weighted'
    
    % 분석 옵션
    config.extreme_group_method = 'all';  % 'extreme' or 'all'
    config.force_recalc_permutation = false;
    config.run_bootstrap = true;
    config.run_permutation = true;
    
    % 파일 관리
    config.create_backup = true;
    config.backup_folder = 'backup';
    config.use_timestamp = false;
    
    % 이상치 제거 설정
    config.outlier_removal = createOutlierConfig();
    
    % 성과 순위 정의
    config.performance_ranking = createPerformanceRanking();
    
    % 고성과자/저성과자 정의
    config.high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
    config.low_performers = {'무능한 불연성', '소화성', '게으른 가연성'};
    config.excluded_from_analysis = {'유능한 불연성'};
    config.excluded_types = {'위장형 소화성'};
    
    % 결과 디렉토리 생성
    if ~exist(config.output_dir, 'dir')
        mkdir(config.output_dir);
    end
end

function outlier_config = createOutlierConfig()
    % 이상치 제거 설정 생성
    
    outlier_config = struct();
    outlier_config.enabled = true;
    outlier_config.method = 'zscore';  % 'iqr', 'zscore', 'percentile', 'none'
    outlier_config.iqr_multiplier = 1.5;
    outlier_config.zscore_threshold = 3;
    outlier_config.percentile_bounds = [5, 95];
    outlier_config.apply_to_competencies = true;
    outlier_config.report_outliers = true;
end

function ranking = createPerformanceRanking()
    % 성과 순위 매핑 생성
    
    ranking = containers.Map(...
        {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
         '게으른 가연성', '무능한 불연성', '소화성'}, ...
        [8, 7, 6, 5, 4, 3, 1]);
end

function setupGraphicsDefaults()
    % 그래픽 기본 설정
    
    set(0, 'DefaultAxesFontName', 'Malgun Gothic');
    set(0, 'DefaultTextFontName', 'Malgun Gothic');
    set(0, 'DefaultAxesFontSize', 12);
    set(0, 'DefaultTextFontSize', 12);
    set(0, 'DefaultLineLineWidth', 2);
end

%% === 데이터 로딩 및 전처리 함수들 ===

function [hr_data, comp_data, comp_total] = loadData(config)
    % 데이터 파일 로딩
    
    fprintf('▶ 데이터 로딩 중...\n');
    
    % 파일 존재 확인
    if ~exist(config.hr_file, 'file')
        error('HR 데이터 파일을 찾을 수 없습니다: %s', config.hr_file);
    end
    if ~exist(config.comp_file, 'file')
        error('역량검사 데이터 파일을 찾을 수 없습니다: %s', config.comp_file);
    end
    
    % 데이터 로딩
    hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
    comp_data = readtable(config.comp_file, 'Sheet', '역량검사_상위항목_단순평균', 'VariableNamingRule', 'preserve');
    comp_total = readtable(config.comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    
    fprintf('  ✓ HR 데이터: %d명\n', height(hr_data));
    fprintf('  ✓ 역량 데이터: %d명\n', height(comp_data));
    fprintf('  ✓ 종합점수 데이터: %d명\n', height(comp_total));
end

function comp_data = applyDataFilters(comp_data, config)
    % 신뢰성 필터 및 개발자 필터 적용
    
    % 신뢰성 필터
    comp_data = applyReliabilityFilter(comp_data);
    
    % 개발자 필터 (선택적)
    comp_data = applyDeveloperFilter(comp_data, config);
end

function comp_data = applyReliabilityFilter(comp_data)
    % 신뢰불가 데이터 제외
    
    fprintf('▶ 신뢰성 필터링...\n');
    
    reliability_col_idx = find(contains(comp_data.Properties.VariableNames, '신뢰가능성'), 1);
    
    if ~isempty(reliability_col_idx)
        reliability_data = comp_data{:, reliability_col_idx};
        
        if iscell(reliability_data)
            unreliable_idx = strcmp(reliability_data, '신뢰불가');
        else
            unreliable_idx = false(height(comp_data), 1);
        end
        
        fprintf('  신뢰불가 데이터: %d명 제거\n', sum(unreliable_idx));
        comp_data = comp_data(~unreliable_idx, :);
        fprintf('  ✓ 신뢰가능한 데이터: %d명\n', height(comp_data));
    else
        fprintf('  ⚠ 신뢰가능성 컬럼 없음\n');
    end
end

function comp_data = applyDeveloperFilter(comp_data, config)
    % 개발자 데이터 필터링
    
    fprintf('▶ 개발자 필터링 옵션...\n');
    
    developer_col_idx = find(contains(comp_data.Properties.VariableNames, ...
                            {'개발자여부', '개발자', 'developer'}), 1);
    
    if isempty(developer_col_idx)
        fprintf('  ⚠ 개발자 컬럼 없음\n');
        return;
    end
    
    % 개발자 판별
    developer_data = comp_data{:, developer_col_idx};
    
    if iscell(developer_data)
        developer_idx = contains(lower(string(developer_data)), {'개발자', 'y', 'yes', '예'});
    else
        developer_idx = (developer_data == 1) | (developer_data == true);
    end
    
    n_dev = sum(developer_idx);
    n_non_dev = sum(~developer_idx);
    
    fprintf('  개발자: %d명, 비개발자: %d명\n', n_dev, n_non_dev);
    
    % 자동으로 모든 데이터 사용 (필요시 수정 가능)
    config.developer_filter = 'all';
    fprintf('  ✓ 모든 데이터 사용\n');
end

%% === 데이터 통합 및 매칭 함수들 ===

function [matched_data, talent_info] = integrateAndMatchData(hr_data, comp_data, comp_total, config)
    % 데이터 통합 및 ID 매칭
    
    fprintf('▶ 데이터 통합 및 매칭...\n');
    
    % 인재유형 추출
    [hr_clean, talent_types] = extractTalentTypes(hr_data, config);
    
    % 역량 데이터 처리
    [valid_comp_cols, valid_comp_indices] = processCompetencyColumns(comp_data);
    
    % ID 매칭
    [matched_hr, matched_comp, matched_talent_types, total_scores] = ...
        matchDataByID(hr_clean, comp_data, comp_total, valid_comp_indices, talent_types);
    
    % 구조체로 정리
    matched_data = struct();
    matched_data.hr = matched_hr;
    matched_data.competencies = matched_comp;
    matched_data.talent_types = matched_talent_types;
    matched_data.total_scores = total_scores;
    matched_data.comp_names = valid_comp_cols;
    
    talent_info = struct();
    talent_info.unique_types = unique(matched_talent_types);
    talent_info.type_counts = countTalentTypes(matched_talent_types);
    
    fprintf('  ✓ 매칭 완료: %d명\n', height(matched_hr));
end

function [hr_clean, talent_types] = extractTalentTypes(hr_data, config)
    % 인재유형 데이터 추출 및 정제
    
    talent_col_idx = find(contains(hr_data.Properties.VariableNames, {'인재유형', '인재', '유형'}), 1);
    
    if isempty(talent_col_idx)
        error('인재유형 컬럼을 찾을 수 없습니다.');
    end
    
    % 빈 값 및 제외 유형 제거
    valid_idx = ~cellfun(@isempty, hr_data{:, talent_col_idx});
    hr_clean = hr_data(valid_idx, :);
    
    excluded_mask = ismember(hr_clean{:, talent_col_idx}, config.excluded_types);
    hr_clean = hr_clean(~excluded_mask, :);
    
    talent_types = hr_clean{:, talent_col_idx};
end

function [valid_comp_cols, valid_comp_indices] = processCompetencyColumns(comp_data)
    % 유효한 역량 컬럼 추출
    
    valid_comp_cols = {};
    valid_comp_indices = [];
    
    for i = 6:width(comp_data)  % 일반적으로 6번째 컬럼부터 역량 데이터
        col_name = comp_data.Properties.VariableNames{i};
        col_data = comp_data{:, i};
        
        if isnumeric(col_data) && ~all(isnan(col_data))
            valid_data = col_data(~isnan(col_data));
            
            if length(valid_data) >= 5
                data_var = var(valid_data);
                if (data_var > 0 || length(unique(valid_data)) > 1) && ...
                   all(valid_data >= 0) && all(valid_data <= 100)
                    valid_comp_cols{end+1} = col_name;
                    valid_comp_indices(end+1) = i;
                end
            end
        end
    end
    
    fprintf('  유효 역량: %d개\n', length(valid_comp_cols));
end

function [matched_hr, matched_comp, matched_talent_types, total_scores] = ...
         matchDataByID(hr_clean, comp_data, comp_total, valid_comp_indices, talent_types)
    % ID 기준 데이터 매칭
    
    % ID 컬럼 찾기
    comp_id_col = find(contains(lower(comp_data.Properties.VariableNames), {'id', '사번'}), 1);
    
    if isempty(comp_id_col)
        error('역량 데이터에서 ID 컬럼을 찾을 수 없습니다.');
    end
    
    % ID 표준화 및 매칭
    hr_ids_str = arrayfun(@(x) sprintf('%.0f', x), hr_clean.ID, 'UniformOutput', false);
    comp_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_data{:, comp_id_col}, 'UniformOutput', false);
    
    [matched_ids, hr_idx, comp_idx] = intersect(hr_ids_str, comp_ids_str);
    
    matched_hr = hr_clean(hr_idx, :);
    matched_comp = comp_data(comp_idx, valid_comp_indices);
    matched_talent_types = talent_types(hr_idx);
    
    % 종합점수 매칭
    total_scores = matchTotalScores(comp_total, matched_ids);
end

function total_scores = matchTotalScores(comp_total, matched_ids)
    % 종합점수 매칭
    
    comp_total_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_total{:, 1}, 'UniformOutput', false);
    [~, ~, total_idx] = intersect(matched_ids, comp_total_ids_str);
    
    if ~isempty(total_idx)
        total_scores = comp_total{total_idx, end};
    else
        total_scores = [];
    end
end

function type_counts = countTalentTypes(talent_types)
    % 인재유형별 개수 계산
    
    [unique_types, ~, type_indices] = unique(talent_types);
    type_counts = accumarray(type_indices, 1);
end

%% === 이상치 제거 함수 ===

function matched_data = removeOutliers(matched_data, config)
    % 이상치 제거
    
    fprintf('▶ 이상치 제거 중...\n');
    
    comp_data = table2array(matched_data.competencies);
    outlier_mask = detectOutliers(comp_data, config.outlier_removal);
    
    % 행별 이상치 판별
    row_outliers = any(outlier_mask, 2);
    clean_indices = ~row_outliers;
    
    fprintf('  제거된 이상치: %d명 (%.1f%%)\n', ...
           sum(row_outliers), sum(row_outliers)/size(comp_data, 1)*100);
    
    % 모든 데이터 필터링
    matched_data.hr = matched_data.hr(clean_indices, :);
    matched_data.competencies = matched_data.competencies(clean_indices, :);
    matched_data.talent_types = matched_data.talent_types(clean_indices);
    
    if ~isempty(matched_data.total_scores)
        matched_data.total_scores = matched_data.total_scores(clean_indices);
    end
end

function outlier_mask = detectOutliers(data, outlier_config)
    % 이상치 탐지
    
    outlier_mask = false(size(data));
    
    switch outlier_config.method
        case 'zscore'
            z_scores = abs(zscore(data, 0, 1));
            outlier_mask = z_scores > outlier_config.zscore_threshold;
            
        case 'iqr'
            Q1 = prctile(data, 25, 1);
            Q3 = prctile(data, 75, 1);
            IQR = Q3 - Q1;
            lower = Q1 - outlier_config.iqr_multiplier * IQR;
            upper = Q3 + outlier_config.iqr_multiplier * IQR;
            
            for col = 1:size(data, 2)
                outlier_mask(:, col) = (data(:, col) < lower(col)) | ...
                                       (data(:, col) > upper(col));
            end
            
        case 'percentile'
            lower = prctile(data, outlier_config.percentile_bounds(1), 1);
            upper = prctile(data, outlier_config.percentile_bounds(2), 1);
            
            for col = 1:size(data, 2)
                outlier_mask(:, col) = (data(:, col) < lower(col)) | ...
                                       (data(:, col) > upper(col));
            end
    end
end

%% === 탐색적 데이터 분석 함수들 ===

function eda_results = performEDA(matched_data, talent_info, config)
    % 탐색적 데이터 분석
    
    fprintf('▶ 탐색적 데이터 분석 수행 중...\n');
    
    eda_results = struct();
    
    % 기술통계
    eda_results.descriptive_stats = calculateDescriptiveStats(matched_data.competencies, ...
                                                             matched_data.comp_names);
    
    % 분포 분석
    eda_results.distribution_stats = analyzeDistributions(matched_data.competencies, ...
                                                         matched_data.comp_names);
    
    % 상관관계
    eda_results.correlation_matrix = corr(table2array(matched_data.competencies), ...
                                         'rows', 'complete');
    
    % 인재유형별 통계
    eda_results.type_stats = calculateTypeStatistics(matched_data, talent_info);
    
    fprintf('  ✓ EDA 완료\n');
end

function stats = calculateDescriptiveStats(comp_data, comp_names)
    % 기술통계 계산
    
    data = table2array(comp_data);
    
    stats = table();
    stats.Competency = comp_names';
    stats.Mean = mean(data, 1, 'omitnan')';
    stats.Std = std(data, 0, 1, 'omitnan')';
    stats.Min = min(data, [], 1, 'omitnan')';
    stats.Q1 = prctile(data, 25, 1)';
    stats.Median = median(data, 1, 'omitnan')';
    stats.Q3 = prctile(data, 75, 1)';
    stats.Max = max(data, [], 1, 'omitnan')';
    stats.Range = stats.Max - stats.Min;
    stats.IQR = stats.Q3 - stats.Q1;
    stats.CV = stats.Std ./ stats.Mean;
end

function dist_stats = analyzeDistributions(comp_data, comp_names)
    % 분포 특성 분석
    
    data = table2array(comp_data);
    n_comps = length(comp_names);
    
    dist_stats = table();
    dist_stats.Competency = comp_names';
    dist_stats.Skewness = zeros(n_comps, 1);
    dist_stats.Kurtosis = zeros(n_comps, 1);
    dist_stats.Normality_pValue = zeros(n_comps, 1);
    
    for i = 1:n_comps
        col_data = data(:, i);
        valid_data = col_data(~isnan(col_data));
        
        if length(valid_data) > 3
            dist_stats.Skewness(i) = skewness(valid_data);
            dist_stats.Kurtosis(i) = kurtosis(valid_data);
            
            % Shapiro-Wilk test (if available) or Lilliefors test
            try
                [~, p] = lillietest(valid_data);
                dist_stats.Normality_pValue(i) = p;
            catch
                dist_stats.Normality_pValue(i) = NaN;
            end
        end
    end
end

function type_stats = calculateTypeStatistics(matched_data, talent_info)
    % 인재유형별 통계 계산 (원본과 동일한 정보 포함)

    unique_types = talent_info.unique_types;
    n_types = length(unique_types);

    type_stats = table();
    type_stats.TalentType = unique_types;
    type_stats.Count = zeros(n_types, 1);
    type_stats.CompetencyMean = zeros(n_types, 1);
    type_stats.CompetencyStd = zeros(n_types, 1);
    type_stats.TotalScoreMean = zeros(n_types, 1);
    type_stats.PerformanceRank = zeros(n_types, 1);

    % performance_ranking이 config에 있다고 가정
    performance_ranking = containers.Map(...
        {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
         '게으른 가연성', '무능한 불연성', '소화성'}, ...
        [8, 7, 6, 5, 4, 3, 1]);

    for i = 1:n_types
        type_name = unique_types{i};
        type_mask = strcmp(matched_data.talent_types, type_name);
        type_comp_data = table2array(matched_data.competencies(type_mask, :));

        type_stats.Count(i) = sum(type_mask);
        type_stats.CompetencyMean(i) = mean(type_comp_data(:), 'omitnan');
        type_stats.CompetencyStd(i) = std(type_comp_data(:), 'omitnan');

        % 종합점수 평균
        if ~isempty(matched_data.total_scores)
            type_total_scores = matched_data.total_scores(type_mask);
            type_stats.TotalScoreMean(i) = mean(type_total_scores, 'omitnan');
        end

        % 성과 순위
        if performance_ranking.isKey(type_name)
            type_stats.PerformanceRank(i) = performance_ranking(type_name);
        end
    end
end

%% === 프로파일 분석 함수들 ===

function profile_results = analyzeTypeProfiles(matched_data, talent_info, config)
    % 인재유형별 프로파일 분석
    
    fprintf('▶ 유형별 프로파일 분석 중...\n');
    
    unique_types = talent_info.unique_types;
    n_types = length(unique_types);
    n_comps = length(matched_data.comp_names);
    
    % 프로파일 매트릭스 초기화
    type_profiles = zeros(n_types, n_comps);
    
    % 유형별 평균 프로파일 계산
    for i = 1:n_types
        type_mask = strcmp(matched_data.talent_types, unique_types{i});
        type_comp_data = table2array(matched_data.competencies(type_mask, :));
        type_profiles(i, :) = mean(type_comp_data, 1, 'omitnan');
    end
    
    % 상위 역량 선정 (분산 기준)
    comp_variance = var(table2array(matched_data.competencies), 0, 1, 'omitnan');
    [~, var_idx] = sort(comp_variance, 'descend');
    top_comp_idx = var_idx(1:min(12, length(var_idx)));
    top_comp_names = matched_data.comp_names(top_comp_idx);
    
    % 전체 평균 계산
    overall_mean = mean(table2array(matched_data.competencies), 1, 'omitnan');
    
    % 가중평균 계산
    weighted_mean = calculateWeightedMean(type_profiles, talent_info.type_counts);
    
    % 결과 구조체
    profile_results = struct();
    profile_results.type_profiles = type_profiles;
    profile_results.unique_types = unique_types;
    profile_results.top_comp_idx = top_comp_idx;
    profile_results.top_comp_names = top_comp_names;
    profile_results.overall_mean = overall_mean;
    profile_results.weighted_mean = weighted_mean;
    profile_results.baseline_type = config.baseline_type;
    
    fprintf('  ✓ 프로파일 분석 완료\n');
end

function weighted_mean = calculateWeightedMean(profiles, counts)
    % 샘플 수 기반 가중평균 계산
    
    weights = counts / sum(counts);
    weighted_mean = weights' * profiles;
end

%% === 시각화 함수들 ===

function createRadarCharts(profile_results, config)
    % 레이더 차트 생성
    
    fprintf('▶ 레이더 차트 생성 중...\n');
    
    n_types = length(profile_results.unique_types);
    colors = lines(n_types);
    
    % 스케일 범위 설정
    all_data = profile_results.type_profiles(:, profile_results.top_comp_idx);
    global_min = min(all_data(:)) - 5;
    global_max = max(all_data(:)) + 5;
    
    for i = 1:n_types
        fig = figure('Position', [100+(i-1)*50, 100+(i-1)*30, 800, 800], ...
                    'Color', 'white', ...
                    'Name', sprintf('인재유형: %s', profile_results.unique_types{i}));
        
        drawRadarChart(profile_results.type_profiles(i, profile_results.top_comp_idx), ...
                      profile_results.top_comp_names, ...
                      profile_results.unique_types{i}, ...
                      getBaseline(profile_results), ...
                      [global_min, global_max], ...
                      colors(i, :));
        
        % 저장
        if config.create_backup
            saveFigure(fig, sprintf('radar_%s.png', profile_results.unique_types{i}), config);
        end
    end
    
    fprintf('  ✓ %d개 차트 생성 완료\n', n_types);
end

function baseline = getBaseline(profile_results)
    % 기준선 선택
    
    if strcmp(profile_results.baseline_type, 'weighted')
        baseline = profile_results.weighted_mean(profile_results.top_comp_idx);
    else
        baseline = profile_results.overall_mean(profile_results.top_comp_idx);
    end
end

function drawRadarChart(data, labels, title_text, baseline, scale_range, color)
    % 실제 레이더 차트 그리기
    
    n_vars = length(data);
    angles = linspace(0, 2*pi, n_vars+1);
    
    % 정규화
    data_norm = (data - scale_range(1)) / (scale_range(2) - scale_range(1));
    baseline_norm = (baseline - scale_range(1)) / (scale_range(2) - scale_range(1));
    
    % 순환을 위해 첫 번째 값 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];
    
    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);
    
    hold on;
    
    % 그리드 그리기
    drawRadarGrid(angles, 5, scale_range);
    
    % 데이터 플롯
    plot(x_base, y_base, '--', 'Color', [0.7 0.7 0.7], 'LineWidth', 2);
    patch(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2.5);
    scatter(x_data(1:end-1), y_data(1:end-1), 60, color, 'filled', ...
           'MarkerEdgeColor', 'white', 'LineWidth', 1);
    
    % 레이블 추가
    addRadarLabels(angles, labels, data, baseline, 1.25);
    
    title(title_text, 'FontSize', 16, 'FontWeight', 'bold');
    legend({'평균선', '해당 유형'}, 'Location', 'best');
    
    axis equal;
    axis([-1.5 1.5 -1.5 1.5]);
    axis off;
    hold off;
end

function drawRadarGrid(angles, n_levels, scale_range)
    % 레이더 차트 그리드 그리기
    
    for j = 1:n_levels
        r = j / n_levels;
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);
        
        % 그리드 레이블
        grid_value = scale_range(1) + (scale_range(2) - scale_range(1)) * r;
        text(0, r, sprintf('%.0f', grid_value), ...
             'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'bottom', ...
             'FontSize', 9, 'Color', [0.4 0.4 0.4]);
    end
    
    % 방사선
    n_vars = length(angles) - 1;
    for j = 1:n_vars
        plot([0, cos(angles(j))], [0, sin(angles(j))], 'k:', ...
             'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);
    end
end

function addRadarLabels(angles, labels, data, baseline, radius)
    % 레이더 차트 레이블 추가
    
    n_vars = length(labels);
    
    for j = 1:n_vars
        [lx, ly] = pol2cart(angles(j), radius);
        
        % 차이값 계산
        diff_val = data(j) - baseline(j);
        diff_str = sprintf('%+.1f', diff_val);
        
        if diff_val > 0
            diff_color = [0, 0.5, 0];
        else
            diff_color = [0.8, 0, 0];
        end
        
        text(lx, ly, sprintf('%s\n%.1f\n(%s)', labels{j}, data(j), diff_str), ...
             'HorizontalAlignment', 'center', ...
             'FontSize', 10, 'FontWeight', 'bold', ...
             'Interpreter', 'none');
    end
end

%% === 예측 모델링 함수들 ===

function [X, y, feature_names] = preparePredictionData(matched_data, talent_info, config)
    % 예측을 위한 데이터 준비
    
    fprintf('▶ 예측 데이터 준비 중...\n');
    
    % 레이블 생성
    y = createBinaryLabels(matched_data.talent_types, config);
    
    % 유효한 샘플만 선택
    valid_idx = ~isnan(y);
    
    X = table2array(matched_data.competencies(valid_idx, :));
    y = y(valid_idx);
    feature_names = matched_data.comp_names;
    
    % 결측값 처리
    [X, y] = handleMissingValues(X, y);
    
    fprintf('  ✓ 최종 데이터: %d샘플 × %d특성\n', size(X, 1), size(X, 2));
    fprintf('    고성과자: %d명, 저성과자: %d명\n', sum(y==1), sum(y==0));
end

function y_binary = createBinaryLabels(talent_types, config)
    % 성과점수 기반 레이블 생성 (원본 파일과 동일한 방식)

    % 먼저 성과점수 할당
    performance_scores = zeros(length(talent_types), 1);
    for i = 1:length(talent_types)
        type_name = talent_types{i};
        if config.performance_ranking.isKey(type_name)
            performance_scores(i) = config.performance_ranking(type_name);
        end
    end

    % 유효한 성과점수가 있는 데이터만 사용
    valid_perf_idx = performance_scores > 0;
    valid_performance = performance_scores(valid_perf_idx);

    % Quartile 기반 극단 그룹 비교 (원본 방식)
    perf_q25 = prctile(valid_performance, 25);
    perf_q75 = prctile(valid_performance, 75);

    y_binary = NaN(length(talent_types), 1);

    for i = 1:length(talent_types)
        perf_score = performance_scores(i);
        if perf_score > 0
            if perf_score >= perf_q75
                y_binary(i) = 1;  % 고성과자 (상위 25%)
            elseif perf_score <= perf_q25
                y_binary(i) = 0;  % 저성과자 (하위 25%)
            else
                y_binary(i) = NaN;  % 중간 성과자 제외
            end
        end
    end
end

function [X_clean, y_clean] = handleMissingValues(X, y)
    % 결측값 처리
    
    % 완전한 케이스만 사용
    complete_cases = ~any(isnan(X), 2);
    
    X_clean = X(complete_cases, :);
    y_clean = y(complete_cases);
    
    fprintf('  결측값 제거: %d샘플 → %d샘플\n', length(y), length(y_clean));
end

%% === 가중치 계산 함수들 ===

function weight_results = calculateAllWeights(X, y, feature_names, config)
    % 모든 가중치 방법 계산
    
    fprintf('▶ 가중치 계산 중...\n');
    
    % 정규화
    X_norm = zscore(X);
    
    weight_results = struct();
    
    % 1. 상관 기반 가중치
    weight_results.correlation = calculateCorrelationWeights(X_norm, y);
    
    % 2. 로지스틱 회귀 가중치
    weight_results.logistic = calculateLogisticWeights(X_norm, y, config);
    
    % 3. T-test 효과크기 가중치
    weight_results.ttest = calculateTtestWeights(X_norm, y);
    
    % 4. Bootstrap 평균 (로지스틱 기반)
    if config.run_bootstrap
        weight_results.bootstrap = calculateBootstrapWeights(X_norm, y, config);
    else
        weight_results.bootstrap = weight_results.logistic;
    end
    
    % 5. 앙상블 가중치
    weight_results.ensemble = calculateEnsembleWeights(weight_results);
    
    % 특성 이름 저장
    weight_results.feature_names = feature_names;
    
    fprintf('  ✓ 5가지 가중치 방법 계산 완료\n');
end

function weights = calculateCorrelationWeights(X, y)
    % 상관 기반 가중치 계산 (원본 파일과 동일한 방식)

    n_features = size(X, 2);
    correlations = zeros(n_features, 1);
    p_values = zeros(n_features, 1);

    % 유효한 데이터만 사용
    valid_idx = ~isnan(y);
    X_valid = X(valid_idx, :);
    y_valid = y(valid_idx);

    for i = 1:n_features
        comp_scores = X_valid(:, i);
        valid_comp_idx = ~isnan(comp_scores);

        if sum(valid_comp_idx) >= 10  % 최소 샘플 수 확인
            [r, p] = corr(comp_scores(valid_comp_idx), y_valid(valid_comp_idx), ...
                         'Type', 'Spearman', 'rows', 'complete');
            correlations(i) = abs(r);
            p_values(i) = p;
        else
            correlations(i) = 0;
            p_values(i) = 1;
        end
    end

    % 유의한 상관계수만 사용 (p < 0.05)
    significant_mask = p_values < 0.05;
    correlations(~significant_mask) = 0;

    % 가중치 정규화
    if sum(correlations) > 0
        weights = correlations / sum(correlations) * 100;
    else
        weights = ones(n_features, 1) / n_features * 100;  % 모든 특성에 동일 가중치
    end
end

function weights = calculateLogisticWeights(X, y, config)
    % 로지스틱 회귀 가중치 계산 (LOOCV 최적화)
    
    % Lambda 최적화
    optimal_lambda = optimizeLambda(X, y);
    
    % 최종 모델 학습
    sample_weights = calculateClassWeights(y);
    
    mdl = fitclinear(X, y, ...
        'Learner', 'logistic', ...
        'Regularization', 'ridge', ...
        'Lambda', optimal_lambda, ...
        'Weights', sample_weights);
    
    % 가중치 추출
    coefficients = mdl.Beta;
    positive_coefs = max(0, coefficients);
    weights = positive_coefs / sum(positive_coefs) * 100;
end

function optimal_lambda = optimizeLambda(X, y)
    % Leave-One-Out CV로 최적 Lambda 찾기
    
    lambda_range = logspace(-3, 0, 10);
    cv_scores = zeros(length(lambda_range), 1);
    
    for lambda_idx = 1:length(lambda_range)
        current_lambda = lambda_range(lambda_idx);
        predictions = performLOOCV(X, y, current_lambda);
        cv_scores(lambda_idx) = mean(predictions == y);
    end
    
    [~, best_idx] = max(cv_scores);
    optimal_lambda = lambda_range(best_idx);
end

function predictions = performLOOCV(X, y, lambda)
    % Leave-One-Out 교차검증
    
    n = length(y);
    predictions = zeros(n, 1);
    
    for i = 1:n
        train_idx = true(n, 1);
        train_idx(i) = false;
        
        X_train = X(train_idx, :);
        y_train = y(train_idx);
        X_test = X(i, :);
        
        try
            mdl = fitclinear(X_train, y_train, ...
                'Learner', 'logistic', ...
                'Regularization', 'ridge', ...
                'Lambda', lambda, ...
                'Verbose', 0);
            
            predictions(i) = predict(mdl, X_test);
        catch
            predictions(i) = mode(y_train);
        end
    end
end

function sample_weights = calculateClassWeights(y)
    % 클래스 불균형을 위한 샘플 가중치 계산
    
    n0 = sum(y == 0);
    n1 = sum(y == 1);
    n_total = length(y);
    
    sample_weights = zeros(size(y));
    sample_weights(y == 0) = n_total / (2 * n0);
    sample_weights(y == 1) = n_total / (2 * n1);
end

function weights = calculateTtestWeights(X, y)
    % T-test 효과크기 기반 가중치
    
    n_features = size(X, 2);
    effect_sizes = zeros(n_features, 1);
    
    high_idx = y == 1;
    low_idx = y == 0;
    
    for i = 1:n_features
        high_scores = X(high_idx, i);
        low_scores = X(low_idx, i);
        
        pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                          (length(low_scores)-1)*var(low_scores)) / ...
                         (length(high_scores) + length(low_scores) - 2));
        
        if pooled_std > 0
            effect_sizes(i) = abs(mean(high_scores) - mean(low_scores)) / pooled_std;
        end
    end
    
    weights = effect_sizes / sum(effect_sizes) * 100;
end

function weights = calculateBootstrapWeights(X, y, config)
    % Bootstrap 평균 가중치
    
    n_bootstrap = 1000;  % 간소화
    n_features = size(X, 2);
    bootstrap_weights = zeros(n_features, n_bootstrap);
    
    for b = 1:n_bootstrap
        % 복원추출
        bootstrap_idx = randsample(length(y), length(y), true);
        X_boot = X(bootstrap_idx, :);
        y_boot = y(bootstrap_idx);
        
        % 상관계수 계산
        correlations = zeros(n_features, 1);
        for i = 1:n_features
            r = corr(X_boot(:, i), y_boot, 'rows', 'complete');
            correlations(i) = abs(r);
        end
        
        bootstrap_weights(:, b) = correlations / sum(correlations) * 100;
    end
    
    weights = mean(bootstrap_weights, 2);
end

function weights = calculateEnsembleWeights(weight_results)
    % 앙상블 가중치 (평균)
    
    all_weights = [weight_results.correlation, ...
                   weight_results.logistic, ...
                   weight_results.ttest, ...
                   weight_results.bootstrap];
    
    weights = mean(all_weights, 2);
end

%% === 모델 학습 및 평가 함수들 ===

function model_results = trainAndEvaluateModels(X, y, weight_results, config)
    % 모든 가중치 방법으로 모델 학습 및 평가
    
    fprintf('▶ 모델 학습 및 평가 중...\n');
    
    X_norm = zscore(X);
    
    methods = {'correlation', 'logistic', 'ttest', 'bootstrap', 'ensemble'};
    model_results = struct();
    
    for i = 1:length(methods)
        method = methods{i};
        weights = weight_results.(method) / 100;  % 백분율을 비율로
        
        % 가중 점수 계산
        weighted_scores = X_norm * weights;
        
        % 성능 평가
        [auc, accuracy, precision, recall, f1] = ...
            evaluateWeightedModel(weighted_scores, y);
        
        % 결과 저장
        model_results.(method) = struct(...
            'AUC', auc, ...
            'accuracy', accuracy, ...
            'precision', precision, ...
            'recall', recall, ...
            'f1_score', f1, ...
            'weights', weights);
        
        fprintf('  %s: F1=%.3f, AUC=%.3f\n', method, f1, auc);
    end
    
    fprintf('  ✓ 모델 평가 완료\n');
end

function [auc, accuracy, precision, recall, f1] = evaluateWeightedModel(scores, y)
    % 가중 점수 기반 모델 평가
    
    % ROC 분석
    [X_roc, Y_roc, T_roc, auc] = perfcurve(y, scores, 1);
    
    % 최적 임계값 (Youden's J)
    J = Y_roc - X_roc;
    [~, opt_idx] = max(J);
    optimal_threshold = T_roc(opt_idx);
    
    % 예측
    predictions = scores > optimal_threshold;
    
    % 성능 지표
    TP = sum(predictions == 1 & y == 1);
    TN = sum(predictions == 0 & y == 0);
    FP = sum(predictions == 1 & y == 0);
    FN = sum(predictions == 0 & y == 1);
    
    accuracy = (TP + TN) / length(y);
    precision = TP / (TP + FP + eps);
    recall = TP / (TP + FN + eps);
    f1 = 2 * (precision * recall) / (precision + recall + eps);
end

%% === 통계적 검증 함수들 ===

function statistical_results = performStatisticalTests(X, y, matched_data, talent_info, config)
    % 통계적 검증 수행
    
    fprintf('▶ 통계적 검증 수행 중...\n');
    
    statistical_results = struct();
    
    % T-test 분석
    statistical_results.ttest = performExtremGroupTtest(X, y, matched_data, config);
    
    % 마할라노비스 거리 분석
    statistical_results.mahalanobis = calculateGroupSeparation(X, y);
    
    % 성별 분석 추가
    statistical_results.gender_analysis = performGenderAnalysis(matched_data, y, config);
    
    % 나이 효과 분석 추가
    statistical_results.age_analysis = performAgeAnalysis(matched_data, y, config);
    
    fprintf('  ✓ 통계 검증 완료\n');
end

function ttest_results = performExtremGroupTtest(X, y, matched_data, config)
    % 극단 그룹 T-test
    
    n_features = size(X, 2);
    
    ttest_results = table();
    ttest_results.Feature = matched_data.comp_names';
    ttest_results.Mean_Diff = zeros(n_features, 1);
    ttest_results.t_statistic = zeros(n_features, 1);
    ttest_results.p_value = zeros(n_features, 1);
    ttest_results.Cohen_d = zeros(n_features, 1);
    
    for i = 1:n_features
        high_scores = X(y == 1, i);
        low_scores = X(y == 0, i);
        
        [h, p, ci, stats] = ttest2(high_scores, low_scores);
        
        ttest_results.Mean_Diff(i) = mean(high_scores) - mean(low_scores);
        ttest_results.t_statistic(i) = stats.tstat;
        ttest_results.p_value(i) = p;
        
        % Cohen's d
        pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                          (length(low_scores)-1)*var(low_scores)) / ...
                         (length(high_scores) + length(low_scores) - 2));
        
        if pooled_std > 0
            ttest_results.Cohen_d(i) = ttest_results.Mean_Diff(i) / pooled_std;
        end
    end
    
    ttest_results = sortrows(ttest_results, 'Cohen_d', 'descend');
end

function mahal_results = calculateGroupSeparation(X, y)
    % 마할라노비스 거리 계산
    
    X_high = X(y == 1, :);
    X_low = X(y == 0, :);
    
    % 평균 벡터
    mean_high = mean(X_high);
    mean_low = mean(X_low);
    
    % Pooled covariance
    n_high = size(X_high, 1);
    n_low = size(X_low, 1);
    
    cov_high = cov(X_high);
    cov_low = cov(X_low);
    
    pooled_cov = ((n_high - 1) * cov_high + (n_low - 1) * cov_low) / ...
                 (n_high + n_low - 2);
    
    % 정규화
    pooled_cov_reg = pooled_cov + eye(size(pooled_cov)) * 1e-6;
    
    % 마할라노비스 거리
    mean_diff = mean_high - mean_low;
    
    try
        L = chol(pooled_cov_reg, 'lower');
        v = L \ mean_diff';
        mahal_distance = sqrt(v' * v);
    catch
        mahal_distance = sqrt(mean_diff * pinv(pooled_cov_reg) * mean_diff');
    end
    
    mahal_results = struct();
    mahal_results.distance = mahal_distance;
    mahal_results.n_high = n_high;
    mahal_results.n_low = n_low;
    
    % 해석
    if mahal_distance >= 2.0
        mahal_results.interpretation = '우수한 그룹 분리';
    elseif mahal_distance >= 1.5
        mahal_results.interpretation = '보통 그룹 분리';
    else
        mahal_results.interpretation = '약한 그룹 분리';
    end
end

%% === Bootstrap 검증 함수 ===

function bootstrap_results = performBootstrapValidation(X, y, feature_names, config)
    % Bootstrap 안정성 검증
    
    fprintf('▶ Bootstrap 검증 수행 중...\n');
    
    n_bootstrap = 5000;
    n_features = size(X, 2);
    
    bootstrap_weights = zeros(n_features, n_bootstrap);
    
    for b = 1:n_bootstrap
        % 복원추출
        idx = randsample(length(y), length(y), true);
        X_boot = X(idx, :);
        y_boot = y(idx);
        
        % 가중치 계산
        weights = calculateCorrelationWeights(zscore(X_boot), y_boot);
        bootstrap_weights(:, b) = weights;
        
        if mod(b, 500) == 0
            fprintf('  진행: %d/%d\n', b, n_bootstrap);
        end
    end
    
    % 통계 계산
    bootstrap_results = table();
    bootstrap_results.Feature = feature_names';
    bootstrap_results.Mean_Weight = mean(bootstrap_weights, 2);
    bootstrap_results.Std_Weight = std(bootstrap_weights, 0, 2);
    bootstrap_results.CI_Lower = prctile(bootstrap_weights, 2.5, 2);
    bootstrap_results.CI_Upper = prctile(bootstrap_weights, 97.5, 2);
    bootstrap_results.CV = bootstrap_results.Std_Weight ./ bootstrap_results.Mean_Weight;
    
    fprintf('  ✓ Bootstrap 완료\n');
end

%% === 퍼뮤테이션 테스트 함수 ===

function permutation_results = performPermutationTests(X, y, model_results, config)
    % 퍼뮤테이션 테스트
    
    fprintf('▶ 퍼뮤테이션 테스트 수행 중...\n');
    
    n_permutations = 5000;
    methods = fieldnames(model_results);
    
    permutation_results = struct();
    
    for i = 1:length(methods)
        method = methods{i};
        observed_f1 = model_results.(method).f1_score;
        observed_auc = model_results.(method).AUC;
        
        null_f1 = zeros(n_permutations, 1);
        null_auc = zeros(n_permutations, 1);
        
        for p = 1:n_permutations
            % 레이블 셔플
            y_shuffled = y(randperm(length(y)));
            
            % 가중치 재계산 (간단한 상관계수만)
            weights = calculateCorrelationWeights(zscore(X), y_shuffled);
            weighted_scores = zscore(X) * (weights / 100);
            
            % 평가
            [auc, ~, ~, ~, f1] = evaluateWeightedModel(weighted_scores, y_shuffled);
            
            null_auc(p) = auc;
            null_f1(p) = f1;
        end
        
        % p-value 계산
        p_value_f1 = sum(null_f1 >= observed_f1) / n_permutations;
        p_value_auc = sum(null_auc >= observed_auc) / n_permutations;
        
        permutation_results.(method) = struct(...
            'observed_f1', observed_f1, ...
            'observed_auc', observed_auc, ...
            'p_value_f1', p_value_f1, ...
            'p_value_auc', p_value_auc, ...
            'null_mean_f1', mean(null_f1), ...
            'null_mean_auc', mean(null_auc));
        
        fprintf('  %s: F1 p=%.4f, AUC p=%.4f\n', method, p_value_f1, p_value_auc);
    end
    
    fprintf('  ✓ 퍼뮤테이션 완료\n');
end

%% === 인구통계 분석 함수 ===

function demo_results = analyzeDemographicEffects(matched_data, y, config)
    % 인구통계 변수 효과 분석
    
    fprintf('▶ 인구통계 효과 분석 중...\n');
    
    demo_results = struct();
    
    % 나이 효과
    age_col_idx = find(strcmp(matched_data.hr.Properties.VariableNames, '만 나이'), 1);
    if ~isempty(age_col_idx)
        ages = matched_data.hr{:, age_col_idx};
        % y와 ages의 길이가 다를 수 있으므로 적절한 길이로 맞춤
        min_length = min(length(ages), length(y));
        ages_subset = ages(1:min_length);
        y_subset = y(1:min_length);

        valid_idx = ages_subset > 0 & ages_subset < 100 & ~isnan(ages_subset) & ~isnan(y_subset);

        if sum(valid_idx) > 5
            demo_results.age_correlation = corr(ages_subset(valid_idx), y_subset(valid_idx), 'Type', 'Spearman');
            fprintf('  나이-성과 상관: %.3f\n', demo_results.age_correlation);
        end
    end
    
    % 성별 효과
    gender_col_idx = find(strcmp(matched_data.hr.Properties.VariableNames, '성별'), 1);
    if ~isempty(gender_col_idx)
        genders = matched_data.hr{:, gender_col_idx};
        % y와 genders의 길이가 다를 수 있으므로 적절한 길이로 맞춤
        min_length = min(length(genders), length(y));
        genders_subset = genders(1:min_length);
        y_subset = y(1:min_length);

        if iscell(genders_subset)
            male_idx = contains(lower(string(genders_subset)), {'남', 'male', 'm'});
            female_idx = contains(lower(string(genders_subset)), {'여', 'female', 'f'});

            if sum(male_idx) > 0 && sum(female_idx) > 0
                [h, p] = ttest2(y_subset(male_idx), y_subset(female_idx));
                demo_results.gender_pvalue = p;
                fprintf('  성별 효과 p-value: %.4f\n', p);
            end
        end
    end
    
    fprintf('  ✓ 인구통계 분석 완료\n');
end

function gender_results = performGenderAnalysis(matched_data, y, config)
    % 성별 분석 수행
    
    fprintf('▶ 성별 분석 수행 중...\n');
    
    gender_results = struct();
    
    % 성별 컬럼 찾기
    gender_col_idx = find(strcmp(matched_data.hr.Properties.VariableNames, '성별'), 1);
    if isempty(gender_col_idx)
        fprintf('  ⚠ 성별 컬럼 없음\n');
        return;
    end
    
    genders = matched_data.hr{:, gender_col_idx};
    min_length = min(length(genders), length(y));
    genders_subset = genders(1:min_length);
    y_subset = y(1:min_length);
    
    if iscell(genders_subset)
        male_idx = contains(lower(string(genders_subset)), {'남', 'male', 'm'});
        female_idx = contains(lower(string(genders_subset)), {'여', 'female', 'f'});
    else
        male_idx = (genders_subset == 1);
        female_idx = (genders_subset == 2);
    end
    
    n_male = sum(male_idx);
    n_female = sum(female_idx);
    
    if n_male < 3 || n_female < 3
        fprintf('  ⚠ 성별 그룹 샘플 수 부족 (남성: %d, 여성: %d)\n', n_male, n_female);
        return;
    end
    
    % 성별별 성과 비교
    male_performance = y_subset(male_idx);
    female_performance = y_subset(female_idx);
    
    [h, p] = ttest2(male_performance, female_performance);
    
    % Cohen's d 계산
    pooled_std = sqrt(((n_male-1)*var(male_performance) + (n_female-1)*var(female_performance)) / ...
                     (n_male + n_female - 2));
    cohens_d = (mean(male_performance) - mean(female_performance)) / pooled_std;
    
    gender_results.sample_sizes = struct('male', n_male, 'female', n_female);
    gender_results.performance = struct('male_mean', mean(male_performance), ...
                                      'female_mean', mean(female_performance), ...
                                      'p_value', p, 'cohens_d', cohens_d);
    
    % 성별별 역량 차이 분석
    comp_data = table2array(matched_data.competencies);
    n_comps = size(comp_data, 2);
    
    gender_diff_results = table();
    gender_diff_results.Competency = matched_data.comp_names';
    gender_diff_results.Male_Mean = zeros(n_comps, 1);
    gender_diff_results.Female_Mean = zeros(n_comps, 1);
    gender_diff_results.Mean_Diff = zeros(n_comps, 1);
    gender_diff_results.Cohen_d = zeros(n_comps, 1);
    gender_diff_results.P_Value = ones(n_comps, 1);
    gender_diff_results.Significant = false(n_comps, 1);
    
    for i = 1:n_comps
        comp_data_subset = comp_data(1:min_length, i);
        male_scores = comp_data_subset(male_idx);
        female_scores = comp_data_subset(female_idx);
        
        % 결측치 제거
        male_scores = male_scores(~isnan(male_scores));
        female_scores = female_scores(~isnan(female_scores));
        
        if length(male_scores) >= 3 && length(female_scores) >= 3
            gender_diff_results.Male_Mean(i) = mean(male_scores);
            gender_diff_results.Female_Mean(i) = mean(female_scores);
            gender_diff_results.Mean_Diff(i) = gender_diff_results.Male_Mean(i) - gender_diff_results.Female_Mean(i);
            
            try
                [h, p] = ttest2(male_scores, female_scores);
                gender_diff_results.P_Value(i) = p;
                gender_diff_results.Significant(i) = p < 0.05;
                
                % Cohen's d
                pooled_std = sqrt(((length(male_scores)-1)*var(male_scores) + ...
                                  (length(female_scores)-1)*var(female_scores)) / ...
                                 (length(male_scores) + length(female_scores) - 2));
                if pooled_std > 0
                    gender_diff_results.Cohen_d(i) = gender_diff_results.Mean_Diff(i) / pooled_std;
                end
            catch
                % t-test 실패 시 기본값 유지
            end
        end
    end
    
    gender_results.competency_differences = gender_diff_results;
    
    fprintf('  ✓ 성별 분석 완료 (남성: %d명, 여성: %d명)\n', n_male, n_female);
end

function age_results = performAgeAnalysis(matched_data, y, config)
    % 나이 효과 분석 수행
    
    fprintf('▶ 나이 효과 분석 수행 중...\n');
    
    age_results = struct();
    
    % 나이 컬럼 찾기
    age_col_idx = find(strcmp(matched_data.hr.Properties.VariableNames, '만 나이'), 1);
    if isempty(age_col_idx)
        fprintf('  ⚠ 나이 컬럼 없음\n');
        return;
    end
    
    ages = matched_data.hr{:, age_col_idx};
    min_length = min(length(ages), length(y));
    ages_subset = ages(1:min_length);
    y_subset = y(1:min_length);
    
    valid_idx = ages_subset > 0 & ages_subset < 100 & ~isnan(ages_subset) & ~isnan(y_subset);
    
    if sum(valid_idx) < 10
        fprintf('  ⚠ 유효한 나이 데이터 부족\n');
        return;
    end
    
    valid_ages = ages_subset(valid_idx);
    valid_y = y_subset(valid_idx);
    
    % 나이-성과 상관관계
    age_perf_corr = corr(valid_ages, valid_y, 'Type', 'Spearman');
    
    % 나이 구간별 분석
    age_quartiles = prctile(valid_ages, [25, 50, 75]);
    age_groups = zeros(length(valid_ages), 1);
    age_groups(valid_ages <= age_quartiles(1)) = 1; % 젊은 그룹
    age_groups(valid_ages > age_quartiles(1) & valid_ages <= age_quartiles(3)) = 2; % 중간 그룹
    age_groups(valid_ages > age_quartiles(3)) = 3; % 나이 많은 그룹
    
    group_means = zeros(3, 1);
    group_sizes = zeros(3, 1);
    for i = 1:3
        group_mask = age_groups == i;
        group_means(i) = mean(valid_y(group_mask));
        group_sizes(i) = sum(group_mask);
    end
    
    age_results.correlation = age_perf_corr;
    age_results.quartiles = age_quartiles;
    age_results.group_means = group_means;
    age_results.group_sizes = group_sizes;
    
    fprintf('  ✓ 나이 효과 분석 완료 (상관: %.3f)\n', age_perf_corr);
end

%% === 결과 통합 함수 ===

function all_results = consolidateResults(profile_results, weight_results, model_results, ...
                                         statistical_results, bootstrap_results, ...
                                         permutation_results, demo_results)
    % 모든 결과 통합
    
    all_results = struct();
    all_results.profiles = profile_results;
    all_results.weights = weight_results;
    all_results.models = model_results;
    all_results.statistics = statistical_results;
    all_results.bootstrap = bootstrap_results;
    all_results.permutation = permutation_results;
    all_results.demographics = demo_results;
    all_results.timestamp = datestr(now);
end

%% === 최종 시각화 함수 ===
%% === 최종 시각화 함수 ===

function createFinalVisualizations(all_results, config)
    % 종합 시각화 생성
    
    fprintf('▶ 최종 시각화 생성 중...\n');
    
    % 1. 가중치 비교 차트
    createWeightComparisonChart(all_results.weights);
    
    % 2. 모델 성능 비교
    createModelPerformanceChart(all_results.models);
    
    % 3. 퍼뮤테이션 테스트 결과
    if ~isempty(all_results.permutation)
        createPermutationChart(all_results.permutation);
    end
    
    % 4. Bootstrap 안정성
    if ~isempty(all_results.bootstrap)
        createBootstrapChart(all_results.bootstrap);
    end
    
    % 5. 성별 분석 시각화
    if isfield(all_results.statistics, 'gender_analysis')
        createGenderAnalysisCharts(all_results.statistics.gender_analysis, config);
    end
    
    % 6. 나이 효과 시각화
    if isfield(all_results.statistics, 'age_analysis')
        createAgeAnalysisCharts(all_results.statistics.age_analysis, config);
    end
    
    % 7. 통계 검증 결과 시각화
    createStatisticalValidationCharts(all_results.statistics, config);
    
    fprintf('  ✓ 시각화 완료\n');
end

function createWeightComparisonChart(weight_results)
    % 가중치 비교 시각화
    
    figure('Position', [100, 100, 1200, 600], 'Color', 'white');
    
    % 상위 15개 특성만
    n_display = min(15, length(weight_results.feature_names));
    
    % 데이터 준비
    all_weights = [weight_results.correlation(1:n_display), ...
                   weight_results.logistic(1:n_display), ...
                   weight_results.ttest(1:n_display), ...
                   weight_results.bootstrap(1:n_display), ...
                   weight_results.ensemble(1:n_display)];
    
    % 막대 그래프
    bar(all_weights);
    
    set(gca, 'XTickLabel', weight_results.feature_names(1:n_display), ...
             'XTickLabelRotation', 45);
    
    legend({'Correlation', 'Logistic', 'T-test', 'Bootstrap', 'Ensemble'}, ...
           'Location', 'best');
    
    ylabel('가중치 (%)');
    title('역량별 가중치 비교 (Top 15)');
    grid on;
end

function createModelPerformanceChart(model_results)
    % 모델 성능 비교
    
    figure('Position', [200, 200, 800, 600], 'Color', 'white');
    
    methods = fieldnames(model_results);
    n_methods = length(methods);
    
    % 성능 지표 추출
    metrics = zeros(n_methods, 4);
    for i = 1:n_methods
        method = methods{i};
        metrics(i, :) = [model_results.(method).accuracy, ...
                        model_results.(method).precision, ...
                        model_results.(method).recall, ...
                        model_results.(method).f1_score];
    end
    
    % 그룹 막대 그래프
    bar(metrics);
    
    set(gca, 'XTickLabel', methods, 'XTickLabelRotation', 45);
    legend({'Accuracy', 'Precision', 'Recall', 'F1-Score'}, 'Location', 'best');
    ylabel('Score');
    ylim([0, 1]);
    title('모델 성능 비교');
    grid on;
end

function createPermutationChart(permutation_results)
    % 퍼뮤테이션 테스트 시각화
    
    figure('Position', [300, 300, 1000, 600], 'Color', 'white');
    
    methods = fieldnames(permutation_results);
    n_methods = length(methods);
    
    % p-value 추출
    p_values_f1 = zeros(n_methods, 1);
    p_values_auc = zeros(n_methods, 1);
    
    for i = 1:n_methods
        method = methods{i};
        p_values_f1(i) = permutation_results.(method).p_value_f1;
        p_values_auc(i) = permutation_results.(method).p_value_auc;
    end
    
    % 서브플롯
    subplot(1, 2, 1);
    bar(p_values_f1);
    yline(0.05, 'r--', 'LineWidth', 2);
    set(gca, 'XTickLabel', methods, 'XTickLabelRotation', 45);
    ylabel('p-value');
    title('F1 Score 유의성');
    
    subplot(1, 2, 2);
    bar(p_values_auc);
    yline(0.05, 'r--', 'LineWidth', 2);
    set(gca, 'XTickLabel', methods, 'XTickLabelRotation', 45);
    ylabel('p-value');
    title('AUC 유의성');
    
    sgtitle('퍼뮤테이션 테스트 결과');
end

function createBootstrapChart(bootstrap_results)
    % Bootstrap 안정성 시각화
    
    figure('Position', [400, 400, 1200, 600], 'Color', 'white');
    
    % CV가 낮은(안정적인) 상위 10개
    [~, stable_idx] = sort(bootstrap_results.CV);
    top_stable = stable_idx(1:min(10, length(stable_idx)));
    
    % 에러바 차트
    errorbar(1:length(top_stable), ...
             bootstrap_results.Mean_Weight(top_stable), ...
             bootstrap_results.Std_Weight(top_stable), ...
             'o', 'MarkerSize', 8, 'LineWidth', 2);
    
    set(gca, 'XTick', 1:length(top_stable), ...
             'XTickLabel', bootstrap_results.Feature(top_stable), ...
             'XTickLabelRotation', 45);
    
    ylabel('가중치 (%)');
    title('Bootstrap 가중치 안정성 (Top 10 안정적 역량)');
    grid on;
end

function createGenderAnalysisCharts(gender_results, config)
    % 성별 분석 시각화
    
    if isempty(gender_results) || ~isfield(gender_results, 'competency_differences')
        return;
    end
    
    fprintf('▶ 성별 분석 시각화 생성 중...\n');
    
    % 1. 성별별 역량 차이 막대그래프
    fig_gender = figure('Position', [100, 100, 1200, 800], 'Color', 'white');
    
    % 상위 15개 역량만 표시
    n_display = min(15, height(gender_results.competency_differences));
    [~, sort_idx] = sort(abs(gender_results.competency_differences.Cohen_d), 'descend');
    top_effects = sort_idx(1:n_display);
    
    effect_sizes = gender_results.competency_differences.Cohen_d(top_effects);
    comp_names = gender_results.competency_differences.Competency(top_effects);
    is_significant = gender_results.competency_differences.Significant(top_effects);
    
    % 색상 설정
    bar_colors = zeros(length(effect_sizes), 3);
    for i = 1:length(effect_sizes)
        if is_significant(i)
            if effect_sizes(i) > 0
                bar_colors(i, :) = [0.2, 0.4, 0.8]; % 진한 파랑 (남성 > 여성)
            else
                bar_colors(i, :) = [0.8, 0.2, 0.4]; % 진한 빨강 (여성 > 남성)
            end
        else
            if effect_sizes(i) > 0
                bar_colors(i, :) = [0.6, 0.7, 0.9]; % 연한 파랑
            else
                bar_colors(i, :) = [0.9, 0.6, 0.7]; % 연한 빨강
            end
        end
    end
    
    b = bar(effect_sizes, 'FaceColor', 'flat');
    b.CData = bar_colors;
    
    % 기준선 추가
    hold on;
    plot([0, length(effect_sizes)+1], [0, 0], 'k-', 'LineWidth', 1);
    plot([0, length(effect_sizes)+1], [0.5, 0.5], 'k--', 'LineWidth', 0.5);
    plot([0, length(effect_sizes)+1], [-0.5, -0.5], 'k--', 'LineWidth', 0.5);
    
    % 유의성 표시
    for i = 1:length(effect_sizes)
        if is_significant(i)
            y_pos = effect_sizes(i) + sign(effect_sizes(i)) * 0.05;
            text(i, y_pos, '*', 'HorizontalAlignment', 'center', ...
                 'FontSize', 16, 'FontWeight', 'bold', 'Color', 'red');
        end
    end
    
    set(gca, 'XTick', 1:length(comp_names), 'XTickLabel', comp_names);
    xtickangle(45);
    ylabel('Cohen''s d (효과크기)');
    title('역량별 성별 차이 (상위 15개, *: p < 0.05)', 'FontSize', 12);
    grid on;
    
    % 범례 추가
    text(0.02, 0.98, sprintf('파랑: 남성 > 여성\n빨강: 여성 > 남성\n진한색: 통계적 유의\n연한색: 비유의'), ...
         'Units', 'normalized', 'VerticalAlignment', 'top', ...
         'BackgroundColor', 'white', 'EdgeColor', 'black', 'Interpreter', 'none');
    
    % 저장
    if config.create_backup
        saveFigure(fig_gender, 'gender_competency_differences.png', config);
    end
    
    fprintf('  ✓ 성별 분석 차트 생성 완료\n');
end

function createAgeAnalysisCharts(age_results, config)
    % 나이 효과 시각화
    
    if isempty(age_results)
        return;
    end
    
    fprintf('▶ 나이 효과 시각화 생성 중...\n');
    
    fig_age = figure('Position', [200, 200, 1000, 600], 'Color', 'white');
    
    % 나이 구간별 성과 비교
    subplot(1, 2, 1);
    bar(age_results.group_means, 'FaceColor', [0.3, 0.6, 0.9]);
    set(gca, 'XTickLabel', {'젊은 그룹', '중간 그룹', '나이 많은 그룹'});
    ylabel('평균 성과 점수');
    title(sprintf('나이 구간별 성과 (상관: %.3f)', age_results.correlation));
    grid on;
    
    % 나이 구간별 샘플 수
    subplot(1, 2, 2);
    bar(age_results.group_sizes, 'FaceColor', [0.9, 0.6, 0.3]);
    set(gca, 'XTickLabel', {'젊은 그룹', '중간 그룹', '나이 많은 그룹'});
    ylabel('샘플 수');
    title('나이 구간별 샘플 수');
    grid on;
    
    sgtitle('나이 효과 분석', 'FontSize', 14, 'FontWeight', 'bold');
    
    % 저장
    if config.create_backup
        saveFigure(fig_age, 'age_effect_analysis.png', config);
    end
    
    fprintf('  ✓ 나이 효과 차트 생성 완료\n');
end

function createStatisticalValidationCharts(statistical_results, config)
    % 통계 검증 결과 시각화
    
    fprintf('▶ 통계 검증 시각화 생성 중...\n');
    
    fig_stats = figure('Position', [300, 300, 1200, 800], 'Color', 'white');
    
    % T-test 결과 시각화
    if isfield(statistical_results, 'ttest')
        subplot(2, 2, 1);
        ttest_data = statistical_results.ttest;
        [~, sort_idx] = sort(abs(ttest_data.Cohen_d), 'descend');
        top_ttest = sort_idx(1:min(15, length(sort_idx)));
        
        bar(ttest_data.Cohen_d(top_ttest), 'FaceColor', [0.2, 0.6, 0.8]);
        set(gca, 'XTick', 1:length(top_ttest), ...
                 'XTickLabel', ttest_data.Feature(top_ttest), ...
                 'XTickLabelRotation', 45);
        ylabel('Cohen''s d');
        title('T-test 효과크기 (Top 15)');
        grid on;
    end
    
    % 마할라노비스 거리 결과
    if isfield(statistical_results, 'mahalanobis')
        subplot(2, 2, 2);
        mahal_data = statistical_results.mahalanobis;
        bar(1, mahal_data.distance, 'FaceColor', [0.8, 0.4, 0.2]);
        ylabel('마할라노비스 거리');
        title(sprintf('그룹 분리도: %.2f\n%s', mahal_data.distance, mahal_data.interpretation));
        set(gca, 'XTick', []);
        grid on;
    end
    
    % 성별 차이 요약
    if isfield(statistical_results, 'gender_analysis')
        subplot(2, 2, 3);
        gender_data = statistical_results.gender_analysis;
        if isfield(gender_data, 'performance')
            perf_data = gender_data.performance;
            bar([perf_data.male_mean, perf_data.female_mean], 'FaceColor', [0.6, 0.8, 0.4]);
            set(gca, 'XTickLabel', {'남성', '여성'});
            ylabel('평균 성과 점수');
            title(sprintf('성별 성과 비교\np=%.4f, d=%.3f', perf_data.p_value, perf_data.cohens_d));
            grid on;
        end
    end
    
    % 나이 효과 요약
    if isfield(statistical_results, 'age_analysis')
        subplot(2, 2, 4);
        age_data = statistical_results.age_analysis;
        if isfield(age_data, 'correlation')
            scatter(1, age_data.correlation, 100, 'filled', 'MarkerFaceColor', [0.4, 0.6, 0.8]);
            ylabel('나이-성과 상관계수');
            title(sprintf('나이 효과\nr=%.3f', age_data.correlation));
            set(gca, 'XTick', []);
            grid on;
        end
    end
    
    sgtitle('통계 검증 결과 종합', 'FontSize', 14, 'FontWeight', 'bold');
    
    % 저장
    if config.create_backup
        saveFigure(fig_stats, 'statistical_validation.png', config);
    end
    
    fprintf('  ✓ 통계 검증 차트 생성 완료\n');
end

%% === 결과 저장 함수들 ===

function saveResults(all_results, config)
    % MAT 파일로 결과 저장
    
    fprintf('▶ 결과 저장 중...\n');
    
    filename = getOutputFilename('analysis_results.mat', config);
    
    % 백업 처리
    if config.create_backup
        backupExistingFile(filename, config);
    end
    
    % 저장
    save(filename, 'all_results');
    fprintf('  ✓ 결과 저장 완료: %s\n', filename);
end

function exportToExcel(all_results, config)
    % Excel 파일로 내보내기
    
    fprintf('▶ Excel 파일 생성 중...\n');
    
    filename = getOutputFilename('HR_Analysis_Results.xlsx', config);
    
    if config.create_backup
        backupExistingFile(filename, config);
    end
    
    try
        % 요약 시트
        writeSummarySheet(filename, all_results);
        
        % 가중치 시트
        writeWeightsSheet(filename, all_results);
        
        % 모델 성능 시트
        writeModelPerformanceSheet(filename, all_results);
        
        % 통계 시트
        writeStatisticsSheet(filename, all_results);
        
        % 성별 분석 시트
        writeGenderAnalysisSheet(filename, all_results);
        
        % 나이 효과 시트
        writeAgeAnalysisSheet(filename, all_results);
        
        % Bootstrap 결과 시트
        writeBootstrapSheet(filename, all_results);
        
        % 퍼뮤테이션 테스트 시트
        writePermutationSheet(filename, all_results);
        
        fprintf('  ✓ Excel 파일 생성 완료: %s\n', filename);
        
    catch ME
        fprintf('  ⚠ Excel 생성 실패: %s\n', ME.message);
    end
end

function writeSummarySheet(filename, all_results)
    % 요약 시트 작성
    
    % 최고 성능 방법 찾기
    methods = fieldnames(all_results.models);
    best_f1 = 0;
    best_method = '';
    
    for i = 1:length(methods)
        if all_results.models.(methods{i}).f1_score > best_f1
            best_f1 = all_results.models.(methods{i}).f1_score;
            best_method = methods{i};
        end
    end
    
    summary = table();
    summary.Item = {'분석일시'; '최적방법'; '최고F1'; '최고AUC'};
    summary.Value = {all_results.timestamp; 
                    best_method;
                    sprintf('%.3f', best_f1);
                    sprintf('%.3f', all_results.models.(best_method).AUC)};
    
    writetable(summary, filename, 'Sheet', '요약');
end

function writeWeightsSheet(filename, all_results)
    % 가중치 시트 작성
    
    weights_table = table();
    weights_table.Feature = all_results.weights.feature_names';
    weights_table.Correlation = all_results.weights.correlation;
    weights_table.Logistic = all_results.weights.logistic;
    weights_table.Ttest = all_results.weights.ttest;
    weights_table.Bootstrap = all_results.weights.bootstrap;
    weights_table.Ensemble = all_results.weights.ensemble;
    
    writetable(weights_table, filename, 'Sheet', '가중치');
end

function writeModelPerformanceSheet(filename, all_results)
    % 모델 성능 시트 작성
    
    methods = fieldnames(all_results.models);
    performance = table();
    
    performance.Method = methods;
    
    for i = 1:length(methods)
        method = methods{i};
        performance.AUC(i) = all_results.models.(method).AUC;
        performance.F1_Score(i) = all_results.models.(method).f1_score;
        performance.Accuracy(i) = all_results.models.(method).accuracy;
        performance.Precision(i) = all_results.models.(method).precision;
        performance.Recall(i) = all_results.models.(method).recall;
    end
    
    writetable(performance, filename, 'Sheet', '모델성능');
end

function writeStatisticsSheet(filename, all_results)
    % 통계 시트 작성
    
    if isfield(all_results.statistics, 'ttest')
        writetable(all_results.statistics.ttest, filename, 'Sheet', 'T-test결과');
    end
    
    if isfield(all_results.statistics, 'mahalanobis')
        mahal_table = table();
        mahal_table.Metric = {'Distance'; 'N_High'; 'N_Low'; 'Interpretation'};
        mahal_table.Value = {all_results.statistics.mahalanobis.distance;
                            all_results.statistics.mahalanobis.n_high;
                            all_results.statistics.mahalanobis.n_low;
                            all_results.statistics.mahalanobis.interpretation};
        
        writetable(mahal_table, filename, 'Sheet', '마할라노비스');
    end
end

function writeGenderAnalysisSheet(filename, all_results)
    % 성별 분석 시트 작성
    
    if isfield(all_results.statistics, 'gender_analysis')
        gender_data = all_results.statistics.gender_analysis;
        
        % 성별 성과 비교
        if isfield(gender_data, 'performance')
            perf_table = table();
            perf_table.Gender = {'Male'; 'Female'};
            perf_table.Mean_Performance = [gender_data.performance.male_mean; gender_data.performance.female_mean];
            perf_table.Sample_Size = [gender_data.sample_sizes.male; gender_data.sample_sizes.female];
            perf_table.P_Value = [gender_data.performance.p_value; NaN];
            perf_table.Cohens_d = [gender_data.performance.cohens_d; NaN];
            
            writetable(perf_table, filename, 'Sheet', '성별성과비교');
        end
        
        % 성별 역량 차이
        if isfield(gender_data, 'competency_differences')
            writetable(gender_data.competency_differences, filename, 'Sheet', '성별역량차이');
        end
    end
end

function writeAgeAnalysisSheet(filename, all_results)
    % 나이 효과 시트 작성
    
    if isfield(all_results.statistics, 'age_analysis')
        age_data = all_results.statistics.age_analysis;
        
        age_table = table();
        age_table.Metric = {'Correlation'; 'Q1'; 'Q2'; 'Q3'; 'Young_Group_Mean'; 'Middle_Group_Mean'; 'Old_Group_Mean'};
        age_table.Value = {age_data.correlation;
                          age_data.quartiles(1);
                          age_data.quartiles(2);
                          age_data.quartiles(3);
                          age_data.group_means(1);
                          age_data.group_means(2);
                          age_data.group_means(3)};
        
        writetable(age_table, filename, 'Sheet', '나이효과분석');
    end
end

function writeBootstrapSheet(filename, all_results)
    % Bootstrap 결과 시트 작성
    
    if isfield(all_results, 'bootstrap') && ~isempty(all_results.bootstrap)
        writetable(all_results.bootstrap, filename, 'Sheet', 'Bootstrap결과');
    end
end

function writePermutationSheet(filename, all_results)
    % 퍼뮤테이션 테스트 시트 작성
    
    if isfield(all_results, 'permutation') && ~isempty(all_results.permutation)
        perm_data = all_results.permutation;
        
        % 각 방법별 결과를 하나의 테이블로 통합
        methods = fieldnames(perm_data);
        n_methods = length(methods);
        
        perm_table = table();
        perm_table.Method = methods;
        perm_table.Observed_F1 = zeros(n_methods, 1);
        perm_table.Observed_AUC = zeros(n_methods, 1);
        perm_table.P_Value_F1 = zeros(n_methods, 1);
        perm_table.P_Value_AUC = zeros(n_methods, 1);
        perm_table.Null_Mean_F1 = zeros(n_methods, 1);
        perm_table.Null_Mean_AUC = zeros(n_methods, 1);
        perm_table.Z_Score_F1 = zeros(n_methods, 1);
        perm_table.Z_Score_AUC = zeros(n_methods, 1);
        
        for i = 1:n_methods
            method = methods{i};
            if isfield(perm_data.(method), 'f1')
                perm_table.Observed_F1(i) = perm_data.(method).f1.observed;
                perm_table.P_Value_F1(i) = perm_data.(method).f1.p_value;
                perm_table.Null_Mean_F1(i) = perm_data.(method).f1.mean_null;
                perm_table.Z_Score_F1(i) = perm_data.(method).f1.z_score;
            end
            
            if isfield(perm_data.(method), 'auc')
                perm_table.Observed_AUC(i) = perm_data.(method).auc.observed;
                perm_table.P_Value_AUC(i) = perm_data.(method).auc.p_value;
                perm_table.Null_Mean_AUC(i) = perm_data.(method).auc.mean_null;
                perm_table.Z_Score_AUC(i) = perm_data.(method).auc.z_score;
            end
        end
        
        writetable(perm_table, filename, 'Sheet', '퍼뮤테이션테스트');
    end
end

%% === 유틸리티 함수들 ===

function filename = getOutputFilename(base_name, config)
    % 출력 파일명 생성
    
    if config.use_timestamp
        [~, name, ext] = fileparts(base_name);
        filename = fullfile(config.output_dir, sprintf('%s_%s%s', name, config.timestamp, ext));
    else
        filename = fullfile(config.output_dir, base_name);
    end
end

function backupExistingFile(filepath, config)
    % 기존 파일 백업
    
    if ~exist(filepath, 'file')
        return;
    end
    
    backup_dir = fullfile(fileparts(filepath), config.backup_folder);
    
    if ~exist(backup_dir, 'dir')
        mkdir(backup_dir);
    end
    
    [~, filename, ext] = fileparts(filepath);
    backup_timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
    backup_filename = sprintf('%s_%s%s', filename, backup_timestamp, ext);
    backup_filepath = fullfile(backup_dir, backup_filename);
    
    try
        movefile(filepath, backup_filepath);
        fprintf('  백업 생성: %s\n', backup_filename);
    catch
        fprintf('  ⚠ 백업 실패\n');
    end
end

function saveFigure(fig, filename, config)
    % Figure 저장
    
    filepath = getOutputFilename(filename, config);
    
    try
        saveas(fig, filepath);
    catch
        fprintf('  ⚠ 그래프 저장 실패: %s\n', filename);
    end
end

%% === 최종 요약 함수 ===

function printFinalSummary(all_results, config)
    % 최종 결과 요약 출력
    
    fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
    fprintf('║                     최종 분석 요약                         ║\n');
    fprintf('╚════════════════════════════════════════════════════════════╝\n\n');
    
    % 최고 성능 방법
    methods = fieldnames(all_results.models);
    best_f1 = 0;
    best_method = '';
    
    for i = 1:length(methods)
        if all_results.models.(methods{i}).f1_score > best_f1
            best_f1 = all_results.models.(methods{i}).f1_score;
            best_method = methods{i};
        end
    end
    
    fprintf('【최적 예측 모델】\n');
    fprintf('  방법: %s\n', best_method);
    fprintf('  F1 Score: %.3f\n', best_f1);
    fprintf('  AUC: %.3f\n', all_results.models.(best_method).AUC);
    fprintf('  정확도: %.3f\n', all_results.models.(best_method).accuracy);
    
    % 핵심 역량 (상위 5개)
    fprintf('\n【핵심 역량 Top 5】\n');
    ensemble_weights = all_results.weights.ensemble;
    [sorted_weights, sort_idx] = sort(ensemble_weights, 'descend');
    
    for i = 1:min(5, length(sorted_weights))
        fprintf('  %d. %s: %.2f%%\n', i, ...
                all_results.weights.feature_names{sort_idx(i)}, ...
                sorted_weights(i));
    end
    
    % 통계적 유의성
    if ~isempty(all_results.permutation)
        fprintf('\n【통계적 유의성】\n');
        
        if isfield(all_results.permutation, best_method)
            method_result = all_results.permutation.(best_method);
            
            if isfield(method_result, 'f1') && isfield(method_result.f1, 'p_value')
                p_value_f1 = method_result.f1.p_value;
            else
                p_value_f1 = 1;
            end
            
            if isfield(method_result, 'auc') && isfield(method_result.auc, 'p_value')
                p_value_auc = method_result.auc.p_value;
            else
                p_value_auc = 1;
            end
            
            if p_value_f1 < 0.001
                sig_level = '***';
            elseif p_value_f1 < 0.01
                sig_level = '**';
            elseif p_value_f1 < 0.05
                sig_level = '*';
            else
                sig_level = 'ns';
            end
            
            fprintf('  F1 p-value: %.4f %s\n', p_value_f1, sig_level);
            fprintf('  AUC p-value: %.4f\n', p_value_auc);
        end
    end
    
    % 인구통계 효과
    if isfield(all_results.statistics, 'gender_analysis')
        fprintf('\n【성별 효과】\n');
        gender_data = all_results.statistics.gender_analysis;
        if isfield(gender_data, 'performance')
            fprintf('  성별 성과 차이: p = %.4f, Cohen''s d = %.3f\n', ...
                    gender_data.performance.p_value, gender_data.performance.cohens_d);
            fprintf('  남성 평균: %.3f, 여성 평균: %.3f\n', ...
                    gender_data.performance.male_mean, gender_data.performance.female_mean);
        end
    end
    
    if isfield(all_results.statistics, 'age_analysis')
        fprintf('\n【나이 효과】\n');
        age_data = all_results.statistics.age_analysis;
        if isfield(age_data, 'correlation')
            fprintf('  나이-성과 상관: r = %.3f\n', age_data.correlation);
        end
    end
    
    % T-test 결과 요약
    if isfield(all_results.statistics, 'ttest')
        fprintf('\n【T-test 결과】\n');
        ttest_data = all_results.statistics.ttest;
        sig_features = ttest_data.p_value < 0.05 & abs(ttest_data.Cohen_d) > 0.5;
        fprintf('  유의한 차별화 역량: %d개 (p<0.05 & |d|>0.5)\n', sum(sig_features));
        
        if sum(sig_features) > 0
            [~, top_idx] = sort(abs(ttest_data.Cohen_d(sig_features)), 'descend');
            top_features = ttest_data(sig_features, :);
            fprintf('  상위 차별화 역량:\n');
            for i = 1:min(3, length(top_idx))
                idx = top_idx(i);
                fprintf('    - %s: d=%.3f, p=%.4f\n', ...
                        top_features.Feature{idx}, ...
                        top_features.Cohen_d(idx), ...
                        top_features.p_value(idx));
            end
        end
    end
    
    % 마할라노비스 거리
    if isfield(all_results.statistics, 'mahalanobis')
        fprintf('\n【그룹 분리도】\n');
        mahal_data = all_results.statistics.mahalanobis;
        fprintf('  마할라노비스 거리: %.3f\n', mahal_data.distance);
        fprintf('  해석: %s\n', mahal_data.interpretation);
    end
    
    % Bootstrap 안정성
    if ~isempty(all_results.bootstrap)
        fprintf('\n【Bootstrap 안정성】\n');
        stable_features = all_results.bootstrap.CV < 0.3;
        fprintf('  안정적인 역량 (CV<0.3): %d개\n', sum(stable_features));
    end
    
    fprintf('\n════════════════════════════════════════════════════════════\n');
    fprintf('✅ 모든 분석이 완료되었습니다!\n');
    fprintf('📊 결과는 Excel 파일과 시각화 차트로 저장되었습니다.\n');
    fprintf('════════════════════════════════════════════════════════════\n');
end