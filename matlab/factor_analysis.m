%% 성과 데이터 통합 및 요인분석 스크립트 (신뢰도 분석 포함)
% 여러 평가 시트의 데이터를 하나의 성과 데이터셋으로 통합하고 체계적 분석 수행
% 작성자: 데이터 분석팀
% 날짜: 2025년 2월

clear; clc;close all

%% 1. 데이터 파일 읽기
filename='C:\Users\MIDASIT\Neurocompetency_project\Competency\(25.02.05) 과거 준거_보정완료 1.xlsx';

% 각 시트 읽기
fprintf('데이터 파일 읽는 중...\n');
try
    team_downward = readtable(filename, 'Sheet', '팀장하향평가_보정');
    diag24_down = readtable(filename, 'Sheet', '24상진단_하향_보정');
    diag24_horizontal = readtable(filename, 'Sheet', '24상진단_수평_보정');
    fprintf('데이터 읽기 완료!\n');
catch ME
    fprintf('오류: %s\n', ME.message);
    return;
end

%% 2. 데이터 구조 확인
fprintf('\n=== 데이터 구조 확인 ===\n');
fprintf('팀장하향평가: %d명, %d개 컬럼\n', height(team_downward), width(team_downward));
fprintf('24상진단_하향: %d명, %d개 컬럼\n', height(diag24_down), width(diag24_down));
fprintf('24상진단_수평: %d명, %d개 컬럼\n', height(diag24_horizontal), width(diag24_horizontal));

%% 3. 측정치 컬럼 식별
% 팀장하향평가: edit_q로 시작하는 컬럼들
team_cols = team_downward.Properties.VariableNames;
team_measure_cols = team_cols(startsWith(team_cols, 'edit_q'));

% 24상진단: edit_LINKED로 시작하는 컬럼들
diag_down_cols = diag24_down.Properties.VariableNames;
diag_down_measure_cols = diag_down_cols(startsWith(diag_down_cols, 'edit_LINKED'));

diag_horizontal_cols = diag24_horizontal.Properties.VariableNames;
diag_horizontal_measure_cols = diag_horizontal_cols(startsWith(diag_horizontal_cols, 'edit_LINKED'));

fprintf('\n=== 측정치 컬럼 수 ===\n');
fprintf('팀장하향평가: %d개 문항\n', length(team_measure_cols));
fprintf('24상진단_하향: %d개 문항\n', length(diag_down_measure_cols));
fprintf('24상진단_수평: %d개 문항\n', length(diag_horizontal_measure_cols));

%% 4. 모든 참가자 ID 수집
all_participants = unique([team_downward.MR_CODE; diag24_down.MR_CODE; diag24_horizontal.MR_CODE]);
fprintf('\n총 참가자 수: %d명\n', length(all_participants));

%% 5. 통합 성과 데이터셋 생성
fprintf('\n통합 데이터셋 생성 중...\n');

% 컬럼명 생성 (중복 구분)
team_col_names = strcat('TeamDown_', team_measure_cols);
diag_down_col_names = strcat('Diag24Down_', diag_down_measure_cols);
diag_horizontal_col_names = strcat('Diag24Horiz_', diag_horizontal_measure_cols);

% 전체 컬럼명
all_measure_cols = [team_col_names, diag_down_col_names, diag_horizontal_col_names];
total_measures = length(all_measure_cols);

fprintf('총 측정치 수: %d개\n', total_measures);

% 통합 데이터 매트릭스 초기화 (NaN으로)
integrated_data = nan(length(all_participants), total_measures);

% 참가자 ID를 인덱스로 매핑
participant_map = containers.Map(all_participants, 1:length(all_participants));

%% 6. 각 시트 데이터를 통합 매트릭스에 입력

% 팀장하향평가 데이터 입력
fprintf('팀장하향평가 데이터 통합 중...\n');
for i = 1:height(team_downward)
    participant_id = team_downward.MR_CODE{i};
    if isKey(participant_map, participant_id)
        row_idx = participant_map(participant_id);
        for j = 1:length(team_measure_cols)
            col_idx = j;
            value = team_downward{i, team_measure_cols{j}};
            if isnumeric(value) && ~isnan(value)
                integrated_data(row_idx, col_idx) = value;
            end
        end
    end
end

% 24상진단_하향 데이터 입력
fprintf('24상진단_하향 데이터 통합 중...\n');
start_col = length(team_measure_cols) + 1;
for i = 1:height(diag24_down)
    participant_id = diag24_down.MR_CODE{i};
    if isKey(participant_map, participant_id)
        row_idx = participant_map(participant_id);
        for j = 1:length(diag_down_measure_cols)
            col_idx = start_col + j - 1;
            value = diag24_down{i, diag_down_measure_cols{j}};
            if isnumeric(value) && ~isnan(value)
                integrated_data(row_idx, col_idx) = value;
            end
        end
    end
end

% 24상진단_수평 데이터 입력
fprintf('24상진단_수평 데이터 통합 중...\n');
start_col = length(team_measure_cols) + length(diag_down_measure_cols) + 1;
for i = 1:height(diag24_horizontal)
    participant_id = diag24_horizontal.MR_CODE{i};
    if isKey(participant_map, participant_id)
        row_idx = participant_map(participant_id);
        for j = 1:length(diag_horizontal_measure_cols)
            col_idx = start_col + j - 1;
            value = diag24_horizontal{i, diag_horizontal_measure_cols{j}};
            if isnumeric(value) && ~isnan(value)
                integrated_data(row_idx, col_idx) = value;
            end
        end
    end
end

%% 7. 통합 테이블 생성
fprintf('\n통합 테이블 생성 중...\n');
integrated_table = array2table(integrated_data, ...
    'VariableNames', all_measure_cols, ...
    'RowNames', all_participants);

fprintf('통합 완료!\n');
fprintf('최종 데이터 크기: %d명 × %d개 측정치\n', height(integrated_table), width(integrated_table));

%% ===== 신뢰도 분석 및 문항 검증 섹션 =====

if ~exist('integrated_table', 'var')
    fprintf('통합 데이터가 없습니다. 먼저 데이터 통합 스크립트를 실행하세요.\n');
    return;
end

fprintf('\n\n===============================================\n');
fprintf('=== 신뢰도 분석 및 문항 검증 시작 ===\n');
fprintf('===============================================\n');

%% 8. 데이터 전처리
data_matrix = table2array(integrated_table);
complete_cases = ~any(isnan(data_matrix), 2);
clean_data = data_matrix(complete_cases, :);
participant_ids = integrated_table.Properties.RowNames(complete_cases);

fprintf('\n=== 데이터 전처리 결과 ===\n');
fprintf('완전한 케이스: %d명 (%.1f%%)\n', size(clean_data, 1), ...
    100 * size(clean_data, 1) / height(integrated_table));
fprintf('분석용 데이터 크기: %d × %d\n', size(clean_data));

%% 9. 기술통계 및 이상치 검출
fprintf('\n=== 기술통계 및 이상치 검출 ===\n');

% 각 변수별 기술통계
data_means = mean(clean_data, 'omitnan');
data_stds = std(clean_data, 'omitnan');
data_medians = median(clean_data, 'omitnan');
data_mins = min(clean_data);
data_maxs = max(clean_data);

fprintf('전체 측정치 기술통계:\n');
fprintf('  평균 범위: %.2f ~ %.2f\n', min(data_means), max(data_means));
fprintf('  표준편차 범위: %.2f ~ %.2f\n', min(data_stds), max(data_stds));
fprintf('  최솟값 범위: %.2f ~ %.2f\n', min(data_mins), max(data_mins));
fprintf('  최댓값 범위: %.2f ~ %.2f\n', min(data_maxs), max(data_maxs));

% 이상치 검출 (Z-score > 3.29, p < 0.001 기준)
fprintf('\n이상치 검출 (|Z-score| > 3.29):\n');
outlier_count = 0;
for j = 1:size(clean_data, 2)
    z_scores = abs(zscore(clean_data(:, j)));
    outliers_in_var = sum(z_scores > 3.29);
    if outliers_in_var > 0
        fprintf('  %s: %d개 이상치\n', all_measure_cols{j}, outliers_in_var);
        outlier_count = outlier_count + outliers_in_var;
    end
end
if outlier_count == 0
    fprintf('  이상치 없음\n');
end

%% 10. 분포 정규성 검정 (Shapiro-Wilk)
fprintf('\n=== 정규성 검정 ===\n');
non_normal_count = 0;

% 표본 크기가 5000 미만일 때만 Shapiro-Wilk 검정 수행
if size(clean_data, 1) < 5000
    fprintf('Shapiro-Wilk 정규성 검정 수행 중...\n');
    for j = 1:min(10, size(clean_data, 2)) % 처음 10개 변수만 검정
        [h, p] = swtest(clean_data(:, j));
        if h == 1 % 정규성 가정 위반
            non_normal_count = non_normal_count + 1;
        end
        if j <= 5 % 처음 5개만 출력
            fprintf('  %s: p = %.4f', all_measure_cols{j}, p);
            if h == 1
                fprintf(' (비정규)\n');
            else
                fprintf(' (정규)\n');
            end
        end
    end
    if size(clean_data, 2) > 10
        fprintf('  ... (나머지 %d개 변수 검정 생략)\n', size(clean_data, 2) - 10);
    end
else
    fprintf('표본 크기가 큼 (n=%d) - 정규성 검정 생략\n', size(clean_data, 1));
end

%% 11. 척도별 신뢰도 분석 (Cronbach's Alpha)
fprintf('\n=== 신뢰도 분석 (Cronbach Alpha) ===\n');

% 각 척도별로 분리해서 신뢰도 계산
scales = struct();
scales.team_downward = 1:length(team_measure_cols);
scales.diag24_down = (length(team_measure_cols)+1):(length(team_measure_cols)+length(diag_down_measure_cols));
scales.diag24_horizontal = (length(team_measure_cols)+length(diag_down_measure_cols)+1):total_measures;

scale_names = {'팀장하향평가', '24상진단_하향', '24상진단_수평'};
scale_fields = {'team_downward', 'diag24_down', 'diag24_horizontal'};

reliability_results = struct();

for s = 1:length(scale_fields)
    scale_name = scale_names{s};
    scale_indices = scales.(scale_fields{s});
    scale_data = clean_data(:, scale_indices);
    
    fprintf('\n--- %s ---\n', scale_name);
    fprintf('문항 수: %d개\n', length(scale_indices));
    
    % 전체 척도 Cronbach Alpha
    alpha_total = cronbach_alpha(scale_data);
    fprintf('전체 Cronbach Alpha: %.3f', alpha_total);
    
    if alpha_total >= 0.9
        fprintf(' (훌륭함)\n');
    elseif alpha_total >= 0.8
        fprintf(' (좋음)\n');
    elseif alpha_total >= 0.7
        fprintf(' (수용가능)\n');
    elseif alpha_total >= 0.6
        fprintf(' (의심스러움)\n');
    else
        fprintf(' (나쁨)\n');
    end
    
    % 문항-전체 상관 분석
    fprintf('\n문항-전체 상관 분석:\n');
    item_total_corrs = zeros(size(scale_data, 2), 1);
    alpha_if_deleted = zeros(size(scale_data, 2), 1);
    
    for i = 1:size(scale_data, 2)
        % 문항-전체 상관 (해당 문항 제외한 나머지와의 상관)
        other_items = scale_data(:, setdiff(1:size(scale_data, 2), i));
        total_score = sum(other_items, 2);
        item_total_corrs(i) = corr(scale_data(:, i), total_score, 'Type', 'Pearson');
        
        % 해당 문항 삭제 시 Alpha
        alpha_if_deleted(i) = cronbach_alpha(other_items);
    end
    
    % 문제 문항 식별
    low_correlation_items = find(item_total_corrs < 0.3);
    high_alpha_if_deleted = find(alpha_if_deleted > alpha_total + 0.05);
    
    fprintf('  평균 문항-전체 상관: %.3f\n', mean(item_total_corrs));
    fprintf('  최저 문항-전체 상관: %.3f\n', min(item_total_corrs));
    
    if ~isempty(low_correlation_items)
        fprintf('  낮은 상관 문항 (r < 0.3): %d개\n', length(low_correlation_items));
        for idx = low_correlation_items'
            fprintf('    %s: r = %.3f\n', all_measure_cols{scale_indices(idx)}, item_total_corrs(idx));
        end
    end
    
    if ~isempty(high_alpha_if_deleted)
        fprintf('  삭제 시 신뢰도 향상 문항: %d개\n', length(high_alpha_if_deleted));
        for idx = high_alpha_if_deleted'
            fprintf('    %s: Alpha = %.3f → %.3f\n', all_measure_cols{scale_indices(idx)}, ...
                alpha_total, alpha_if_deleted(idx));
        end
    end
    
    % 결과 저장
    reliability_results.(scale_fields{s}) = struct();
    reliability_results.(scale_fields{s}).alpha = alpha_total;
    reliability_results.(scale_fields{s}).item_total_corr = item_total_corrs;
    reliability_results.(scale_fields{s}).alpha_if_deleted = alpha_if_deleted;
    reliability_results.(scale_fields{s}).problematic_items = union(low_correlation_items, high_alpha_if_deleted);
end

%% 12. 전체 데이터셋 신뢰도 분석
fprintf('\n=== 전체 데이터셋 신뢰도 ===\n');
overall_alpha = cronbach_alpha(clean_data);
fprintf('전체 Cronbach Alpha: %.3f', overall_alpha);

if overall_alpha >= 0.9
    fprintf(' (훌륭함)\n');
elseif overall_alpha >= 0.8
    fprintf(' (좋음)\n');
elseif overall_alpha >= 0.7
    fprintf(' (수용가능)\n');
elseif overall_alpha >= 0.6
    fprintf(' (의심스러움)\n');
else
    fprintf(' (나쁨)\n');
end

%% 13. 상관분석 및 다중공선성 검진
fprintf('\n=== 상관분석 및 다중공선성 검진 ===\n');

correlation_matrix = corr(clean_data, 'Type', 'Pearson');
upper_tri_corrs = correlation_matrix(triu(true(size(correlation_matrix)), 1));

fprintf('상관계수 기술통계:\n');
fprintf('  평균 상관계수: %.3f\n', mean(upper_tri_corrs));
fprintf('  최대 상관계수: %.3f\n', max(upper_tri_corrs));
fprintf('  최소 상관계수: %.3f\n', min(upper_tri_corrs));

% 높은 상관 (다중공선성 의심) 검출
high_corr_threshold = 0.8;
[high_corr_rows, high_corr_cols] = find(triu(correlation_matrix, 1) > high_corr_threshold);

if ~isempty(high_corr_rows)
    fprintf('\n높은 상관관계 (r > %.1f) 검출:\n', high_corr_threshold);
    for i = 1:length(high_corr_rows)
        corr_val = correlation_matrix(high_corr_rows(i), high_corr_cols(i));
        fprintf('  %s ↔ %s: r = %.3f\n', ...
            all_measure_cols{high_corr_rows(i)}, all_measure_cols{high_corr_cols(i)}, corr_val);
    end
else
    fprintf('다중공선성 문제 없음 (모든 상관계수 < %.1f)\n', high_corr_threshold);
end

%% 14. 요인분석 적합성 재검정
fprintf('\n=== 요인분석 적합성 검정 ===\n');

% 표본 충분성 확인
sample_to_variable_ratio = size(clean_data, 1) / size(clean_data, 2);
fprintf('표본/변수 비율: %.1f:1', sample_to_variable_ratio);
if sample_to_variable_ratio >= 10
    fprintf(' (매우 좋음)\n');
elseif sample_to_variable_ratio >= 5
    fprintf(' (좋음)\n');
elseif sample_to_variable_ratio >= 3
    fprintf(' (최소 기준)\n');
else
    fprintf(' (부족함)\n');
    warning('표본 크기가 부족할 수 있습니다. 요인분석 결과 해석에 주의하세요.');
end

% KMO(Kaiser-Meyer-Olkin) 측도 계산
kmo_value = calculate_kmo(correlation_matrix);
fprintf('KMO 측도: %.3f ', kmo_value);
if kmo_value >= 0.8
    fprintf('(매우 좋음)\n');
elseif kmo_value >= 0.7
    fprintf('(좋음)\n');
elseif kmo_value >= 0.6
    fprintf('(보통)\n');
elseif kmo_value >= 0.5
    fprintf('(나쁨)\n');
else
    fprintf('(매우 나쁨)\n');
end

% 개별 변수 MSA (Measure of Sampling Adequacy) 계산
individual_msa = calculate_individual_msa(correlation_matrix);
low_msa_vars = find(individual_msa < 0.5);

if ~isempty(low_msa_vars)
    fprintf('\n낮은 MSA 변수들 (< 0.5):\n');
    for idx = low_msa_vars'
        fprintf('  %s: MSA = %.3f\n', all_measure_cols{idx}, individual_msa(idx));
    end
else
    fprintf('모든 변수의 MSA ≥ 0.5\n');
end

% Bartlett의 구형성 검정
[bartlett_chi2, bartlett_p] = bartlett_test(correlation_matrix, size(clean_data, 1));
fprintf('\nBartlett 구형성 검정:\n');
fprintf('  χ² = %.2f, p = %.6f\n', bartlett_chi2, bartlett_p);
if bartlett_p < 0.001
    fprintf('  구형성 검정 통과 (p < 0.001) - 요인분석 적합\n');
else
    fprintf('  주의: 구형성 검정 미통과 (p ≥ 0.001)\n');
end

%% 15. 문항 간 상관 매트릭스 분석
fprintf('\n=== 문항 간 상관 패턴 분석 ===\n');

% 상관계수 분포 분석
corr_dist = upper_tri_corrs(~isnan(upper_tri_corrs));
fprintf('상관계수 분포:\n');
fprintf('  Q1 (25%%): %.3f\n', prctile(corr_dist, 25));
fprintf('  중앙값: %.3f\n', median(corr_dist));
fprintf('  Q3 (75%%): %.3f\n', prctile(corr_dist, 75));

% 약한 상관관계 변수 식별
weak_corr_threshold = 0.1;
weak_corr_vars = [];
for i = 1:size(correlation_matrix, 1)
    avg_corr_with_others = mean(abs(correlation_matrix(i, setdiff(1:size(correlation_matrix, 1), i))));
    if avg_corr_with_others < weak_corr_threshold
        weak_corr_vars = [weak_corr_vars, i];
    end
end

if ~isempty(weak_corr_vars)
    fprintf('\n약한 상관관계 변수들 (평균 |r| < %.1f):\n', weak_corr_threshold);
    for idx = weak_corr_vars
        avg_corr = mean(abs(correlation_matrix(idx, setdiff(1:size(correlation_matrix, 1), idx))));
        fprintf('  %s: 평균 |r| = %.3f\n', all_measure_cols{idx}, avg_corr);
    end
else
    fprintf('모든 변수가 적절한 상관관계를 보임\n');
end

%% 16. 최적 요인 수 결정 (개선된 버전)
fprintf('\n\n=== 최적 요인 수 결정 ===\n');
max_factors = min(10, size(clean_data, 2) - 1);

% 고유값 계산
eigenvalues = eig(correlation_matrix);
eigenvalues = sort(eigenvalues, 'descend');

% 1) Kaiser 기준 (고유값 > 1)
kaiser_factors = sum(eigenvalues > 1);
fprintf('Kaiser 기준 (고유값 > 1): %d개 요인\n', kaiser_factors);

% 2) 스크리 검사용 고유값 출력
fprintf('\n고유값 및 분산 설명률:\n');
total_variance = sum(eigenvalues);
for i = 1:min(10, length(eigenvalues))
    var_explained = eigenvalues(i) / total_variance * 100;
    cumvar_explained = sum(eigenvalues(1:i)) / total_variance * 100;
    fprintf('  요인 %2d: 고유값 = %6.3f, 설명률 = %5.1f%%, 누적 = %5.1f%%', ...
        i, eigenvalues(i), var_explained, cumvar_explained);
    if eigenvalues(i) > 1
        fprintf(' *');
    end
    fprintf('\n');
end

% 3) 분산 설명률 기준
cumvar_explained = cumsum(eigenvalues) / sum(eigenvalues) * 100;
factors_60 = find(cumvar_explained >= 60, 1);
factors_70 = find(cumvar_explained >= 70, 1);
factors_80 = find(cumvar_explained >= 80, 1);

fprintf('\n분산 설명률 기준:\n');
fprintf('  60%% 설명: %d개 요인 (%.1f%%)\n', factors_60, cumvar_explained(factors_60));
fprintf('  70%% 설명: %d개 요인 (%.1f%%)\n', factors_70, cumvar_explained(factors_70));
fprintf('  80%% 설명: %d개 요인 (%.1f%%)\n', factors_80, cumvar_explained(factors_80));

% 4) Parallel Analysis (개선된 버전)
fprintf('\n병렬분석 수행 중...\n');
parallel_factors = parallel_analysis_robust(clean_data, 1000);
fprintf('병렬분석 권장 요인 수: %d개\n', parallel_factors);

% 5) Velicer's MAP (Minimum Average Partial) 기준
fprintf('\nVelicer MAP 기준 계산 중...\n');
map_factors = velicer_map(correlation_matrix);
fprintf('MAP 기준 권장 요인 수: %d개\n', map_factors);

% 최종 권장 요인 수 결정 (여러 기준의 가중 평균)
criteria_factors = [kaiser_factors, parallel_factors, map_factors, factors_70];
weights = [0.2, 0.3, 0.3, 0.2]; % 병렬분석과 MAP에 더 큰 가중치
recommended_factors = round(sum(criteria_factors .* weights));

if recommended_factors < 2
    recommended_factors = 2;
elseif recommended_factors > 8
    recommended_factors = 8;
end

fprintf('\n권장 요인 수 종합:\n');
fprintf('  Kaiser: %d, Parallel: %d, MAP: %d, 70%% 분산: %d\n', ...
    kaiser_factors, parallel_factors, map_factors, factors_70);
fprintf('  최종 권장: %d개 요인\n', recommended_factors);

%% 17. 요인분석 수행 및 모델 비교
fprintf('\n=== 요인분석 수행 및 모델 비교 ===\n');




factor_range = max(2, recommended_factors-1):min(recommended_factors+2, max_factors);
fit_indices = zeros(length(factor_range), 6); % [요인수, χ²p값, RMSEA, 평균공통성, 해석가능성, 종합점수]

fprintf('다양한 요인 수 모델 비교:\n');
fprintf('%-6s %-8s %-8s %-8s %-8s %-8s\n', '요인수', 'χ²p값', 'RMSEA', '평균공통성', '해석성', '종합점수');
fprintf('%s\n', repmat('-', 1, 55));

for idx = 1:length(factor_range)
    num_factors = factor_range(idx);
    
    try
        % 주축 요인분석 수행
        [loadings, specific_var, T, stats] = factoran(clean_data, num_factors, ...
            'rotate', 'varimax', 'scores', 'regression');
        
        % 적합도 지수들 계산
        chi2_p = stats.p;
        
        % RMSEA 계산 (근사적)
        if isfield(stats, 'chi2')
            df = stats.dfe;
            rmsea = sqrt(max(0, (stats.chi2 - df) / (df * (size(clean_data, 1) - 1))));
        else
            rmsea = NaN;
        end
        
        % 평균 공통성
        communality = sum(loadings.^2, 2);
        avg_communality = mean(communality);
        
        % 해석가능성 점수 (각 요인에 대해 로딩 0.4 이상인 변수 수의 균형성)
        interpretability_score = calculate_interpretability(loadings);
        
        % 종합 적합도 점수
        composite_score = avg_communality * 0.4 + (1-rmsea) * 0.3 + interpretability_score * 0.3;
        if isnan(rmsea)
            composite_score = avg_communality * 0.7 + interpretability_score * 0.3;
        end
        
        fit_indices(idx, :) = [num_factors, chi2_p, rmsea, avg_communality, interpretability_score, composite_score];
        
        fprintf('%-6d %-8.4f %-8.4f %-8.3f %-8.3f %-8.3f\n', ...
            num_factors, chi2_p, rmsea, avg_communality, interpretability_score, composite_score);
        
    catch ME
        fprintf('%-6d 분석 실패: %s\n', num_factors, ME.message);
        fit_indices(idx, :) = [num_factors, NaN, NaN, NaN, NaN, -999];
    end
end

% 최적 요인 수 선택
valid_rows = fit_indices(:, 6) > -999;
if any(valid_rows)
    [~, best_idx] = max(fit_indices(valid_rows, 6));
    valid_indices = find(valid_rows);
    optimal_factors = fit_indices(valid_indices(best_idx), 1);
    fprintf('\n최적 요인 수: %d개 (종합 점수: %.3f)\n', optimal_factors, fit_indices(valid_indices(best_idx), 6));
else
    optimal_factors = recommended_factors;
    fprintf('\n모든 모델 분석 실패. 권장 요인 수 사용: %d개\n', optimal_factors);
end

%% 18. 최종 요인분석 및 상세 해석
fprintf('\n\n=== 최종 요인분석 결과 (Varimax 회전) ===\n');

try
    [final_loadings, specific_var, T, stats] = factoran(clean_data, optimal_factors, ...
        'rotate', 'varimax', 'scores', 'regression');
    
    % 공통성 및 특수성 계산
    communality = sum(final_loadings.^2, 2);
    specificity = specific_var;
    
    % 요인별 분산 설명률
    factor_variance = sum(final_loadings.^2, 1);
    percent_variance = factor_variance / size(clean_data, 2) * 100;
    
    fprintf('\n요인별 분산 설명률:\n');
    cumulative_variance = 0;
    for f = 1:optimal_factors
        cumulative_variance = cumulative_variance + percent_variance(f);
        fprintf('  요인 %d: %.1f%% (고유값: %.2f, 누적: %.1f%%)\n', ...
            f, percent_variance(f), factor_variance(f), cumulative_variance);
    end
    
    %% 19. 요인 로딩 매트릭스 출력 (개선된 버전)
    fprintf('\n=== 요인 로딩 매트릭스 ===\n');
    fprintf('(절댓값 0.3 이상 굵게 표시, 0.5 이상 *** 표시)\n\n');
    
    % 헤더 출력
    fprintf('%-30s', '측정치');
    for f = 1:optimal_factors
        fprintf('%10s', sprintf('F%d', f));
    end
    fprintf('%10s %10s\n', 'h²', '특수성');
    fprintf('%s\n', repmat('-', 1, 30 + 10*optimal_factors + 20));
    
    % 각 측정치별 로딩 출력 (높은 로딩 순으로 정렬)
    [~, sort_order] = sort(max(abs(final_loadings), [], 2), 'descend');
    
    for idx = 1:length(sort_order)
        i = sort_order(idx);
        % 측정치명 축약
        var_name = all_measure_cols{i};
        if length(var_name) > 29
            var_name = [var_name(1:26), '...'];
        end
        fprintf('%-30s', var_name);
        
        % 요인 로딩 출력
        max_loading = 0;
        for f = 1:optimal_factors
            loading = final_loadings(i, f);
            if abs(loading) >= 0.5
                fprintf('%9.3f*', loading);
            elseif abs(loading) >= 0.3
                fprintf('%10.3f', loading);
            else
                fprintf('%10s', '');
            end
            max_loading = max(max_loading, abs(loading));
        end
        
        % 공통성과 특수성 출력
        fprintf('%10.3f %10.3f', communality(i), specificity(i));
        
        % 복잡성 표시
        high_loadings = sum(abs(final_loadings(i, :)) >= 0.3);
        if high_loadings > 1
            fprintf(' (복잡)');
        end
        fprintf('\n');
    end
    
    %% 20. 요인별 상세 해석
    fprintf('\n=== 요인별 상세 해석 ===\n');
    
    factor_interpretation = cell(optimal_factors, 1);
    
    for f = 1:optimal_factors
        fprintf('\n--- 요인 %d (분산 설명률: %.1f%%) ---\n', f, percent_variance(f));
        
        % 해당 요인의 로딩을 절댓값 기준으로 정렬
        [sorted_loadings, sort_idx] = sort(abs(final_loadings(:, f)), 'descend');
        
        % 고로딩 변수들 (0.4 이상)
        high_loading_vars = [];
        moderate_loading_vars = [];
        
        for i = 1:length(sort_idx)
            var_idx = sort_idx(i);
            loading_val = final_loadings(var_idx, f);
            abs_loading = abs(loading_val);
            
            if abs_loading >= 0.5
                high_loading_vars = [high_loading_vars; var_idx];
                fprintf('  ★★ %s: %.3f (강한 로딩)\n', all_measure_cols{var_idx}, loading_val);
            elseif abs_loading >= 0.3
                moderate_loading_vars = [moderate_loading_vars; var_idx];
                fprintf('  ★  %s: %.3f (중간 로딩)\n', all_measure_cols{var_idx}, loading_val);
            end
        end
        
        % 요인 해석 제안
        if ~isempty(high_loading_vars) || ~isempty(moderate_loading_vars)
            fprintf('\n  해석 제안:\n');
            
            % 변수명에서 공통 패턴 추출
            significant_vars = [high_loading_vars; moderate_loading_vars];
            var_names = all_measure_cols(significant_vars);
            
            % 척도별 분포 확인
            team_count = sum(contains(var_names, 'TeamDown'));
            diag_down_count = sum(contains(var_names, 'Diag24Down'));
            diag_horiz_count = sum(contains(var_names, 'Diag24Horiz'));
            
            fprintf('    - 팀장하향평가: %d개 문항\n', team_count);
            fprintf('    - 24상진단_하향: %d개 문항\n', diag_down_count);
            fprintf('    - 24상진단_수평: %d개 문항\n', diag_horiz_count);
            
            % 주요 평가 영역 추론
            if team_count > diag_down_count + diag_horiz_count
                fprintf('    → 주로 "팀장 관점의 하향 평가" 요인\n');
            elseif diag_down_count > team_count + diag_horiz_count
                fprintf('    → 주로 "상급자 진단 하향 평가" 요인\n');
            elseif diag_horiz_count > team_count + diag_down_count
                fprintf('    → 주로 "동료 간 수평 평가" 요인\n');
            else
                fprintf('    → 혼합된 평가 관점 요인\n');
            end
        else
            fprintf('  해석 어려움: 유의한 로딩 변수 없음\n');
        end
        
        factor_interpretation{f} = var_names;
    end
    
    %% 21. 신뢰도 및 타당도 종합 평가
    fprintf('\n\n=== 신뢰도 및 타당도 종합 평가 ===\n');
    
    % 각 요인별 신뢰도 계산
    fprintf('요인별 신뢰도 (Cronbach Alpha):\n');
    factor_reliabilities = zeros(optimal_factors, 1);
    
    for f = 1:optimal_factors
        % 해당 요인에 높게 로딩되는 변수들로 신뢰도 계산
        high_loading_indices = find(abs(final_loadings(:, f)) >= 0.4);
        
        if length(high_loading_indices) >= 3 % 최소 3개 문항
            factor_data = clean_data(:, high_loading_indices);
            factor_alpha = cronbach_alpha(factor_data);
            factor_reliabilities(f) = factor_alpha;
            
            fprintf('  요인 %d: α = %.3f (%d개 문항)', f, factor_alpha, length(high_loading_indices));
            if factor_alpha >= 0.8
                fprintf(' (좋음)\n');
            elseif factor_alpha >= 0.7
                fprintf(' (수용가능)\n');
            else
                fprintf(' (개선필요)\n');
            end
        else
            fprintf('  요인 %d: 신뢰도 계산 불가 (문항 수 부족: %d개)\n', f, length(high_loading_indices));
            factor_reliabilities(f) = NaN;
        end
    end
    
    % 전체 모델 평가
    fprintf('\n전체 모델 평가:\n');
    avg_communality = mean(communality);
    low_communality_count = sum(communality < 0.5);
    
    fprintf('  평균 공통성: %.3f\n', avg_communality);
    fprintf('  낮은 공통성 변수 (< 0.5): %d개\n', low_communality_count);
    
    if avg_communality >= 0.6
        fprintf('  공통성 수준: 좋음\n');
    elseif avg_communality >= 0.4
        fprintf('  공통성 수준: 보통\n');
    else
        fprintf('  공통성 수준: 개선 필요\n');
    end
    
    %% 22. 문제 문항 및 개선 제안
    fprintf('\n=== 문제 문항 식별 및 개선 제안 ===\n');
    
    problematic_items = [];
    
    % 1) 낮은 공통성 문항
    low_communality_items = find(communality < 0.3);
    if ~isempty(low_communality_items)
        fprintf('\n1) 낮은 공통성 문항 (h² < 0.3):\n');
        for idx = low_communality_items'
            fprintf('  %s: h² = %.3f\n', all_measure_cols{idx}, communality(idx));
            problematic_items = [problematic_items; idx];
        end
    end
    
    % 2) 복잡한 로딩 구조 문항
    complex_items = [];
    for i = 1:size(final_loadings, 1)
        high_loadings_count = sum(abs(final_loadings(i, :)) >= 0.3);
        if high_loadings_count > 1
            complex_items = [complex_items; i];
        end
    end
    
    if ~isempty(complex_items)
        fprintf('\n2) 복잡한 로딩 구조 문항 (2개 이상 요인에 0.3+ 로딩):\n');
        for idx = complex_items'
            loadings_str = '';
            for f = 1:optimal_factors
                if abs(final_loadings(idx, f)) >= 0.3
                    loadings_str = [loadings_str, sprintf('F%d:%.2f ', f, final_loadings(idx, f))];
                end
            end
            fprintf('  %s: %s\n', all_measure_cols{idx}, loadings_str);
        end
    end
    
    % 3) 이전 신뢰도 분석에서 발견된 문제 문항들
    all_problematic_reliability = [];
    for s = 1:length(scale_fields)
        scale_indices = scales.(scale_fields{s});
        prob_items = reliability_results.(scale_fields{s}).problematic_items;
        if ~isempty(prob_items)
            all_problematic_reliability = [all_problematic_reliability; scale_indices(prob_items)'];
        end
    end
    
    if ~isempty(all_problematic_reliability)
        fprintf('\n3) 신뢰도 분석에서 발견된 문제 문항:\n');
        for idx = all_problematic_reliability'
            fprintf('  %s\n', all_measure_cols{idx});
        end
    end
    
    % 종합 문제 문항 리스트
    all_problematic = unique([problematic_items; complex_items; all_problematic_reliability]);
    
    if ~isempty(all_problematic)
        fprintf('\n=== 개선 제안 ===\n');
        fprintf('검토가 필요한 문항: %d개 (전체의 %.1f%%)\n', ...
            length(all_problematic), length(all_problematic)/total_measures*100);
        fprintf('\n권장 조치:\n');
        fprintf('  1) 낮은 공통성 문항: 문항 내용 재검토 또는 삭제 고려\n');
        fprintf('  2) 복잡한 로딩 문항: 문항 표현 명확화 또는 분리\n');
        fprintf('  3) 낮은 신뢰도 문항: 척도 내 일관성 검토\n');
        
        % 삭제 권장 문항 (여러 문제가 중복되는 경우)
        item_problem_count = zeros(total_measures, 1);
        for idx = all_problematic'
            if ismember(idx, low_communality_items)
                item_problem_count(idx) = item_problem_count(idx) + 1;
            end
            if ismember(idx, complex_items)
                item_problem_count(idx) = item_problem_count(idx) + 1;
            end
            if ismember(idx, all_problematic_reliability)
                item_problem_count(idx) = item_problem_count(idx) + 1;
            end
        end
        
        severe_problem_items = find(item_problem_count >= 2);
        if ~isempty(severe_problem_items)
            fprintf('\n  우선 삭제 검토 대상 (2개 이상 문제):\n');
            for idx = severe_problem_items'
                fprintf('    %s (문제 수: %d)\n', all_measure_cols{idx}, item_problem_count(idx));
            end
        end
    else
        fprintf('\n=== 분석 결과: 모든 문항이 적절함 ===\n');
    end
    
    %% 23. 요인 점수 계산 및 기술통계
    fprintf('\n=== 요인 점수 분석 ===\n');
    
    % 요인 점수 계산 (회귀 방법)
    factor_scores = clean_data * pinv(final_loadings');
    
    fprintf('요인 점수 기술통계:\n');
    for f = 1:optimal_factors
        scores_f = factor_scores(:, f);
        fprintf('  요인 %d: M = %6.3f, SD = %6.3f, 범위 = [%6.3f, %6.3f]\n', ...
            f, mean(scores_f), std(scores_f), min(scores_f), max(scores_f));
    end
    
    % 요인 간 상관
    factor_correlations = corr(factor_scores);
    fprintf('\n요인 간 상관:\n');
    for i = 1:optimal_factors
        for j = i+1:optimal_factors
            fprintf('  요인 %d ↔ 요인 %d: r = %.3f\n', i, j, factor_correlations(i, j));
        end
    end
    
    %% 24. 시각화 (개선된 버전)
    fprintf('\n시각화 생성 중...\n');
    
    % 1) 스크리 도표 (개선된)
    figure('Name', '스크리 도표', 'Position', [100, 600, 700, 450]);
    subplot(2, 1, 1);
    plot(1:min(15, length(eigenvalues)), eigenvalues(1:min(15, length(eigenvalues))), 'bo-', 'LineWidth', 2, 'MarkerSize', 8);
    hold on;
    yline(1, 'r--', 'Kaiser 기준', 'LineWidth', 1.5);
    if parallel_factors <= 15
        xline(parallel_factors + 0.5, 'g--', 'Parallel Analysis', 'LineWidth', 1.5);
    end
    xlabel('요인 번호');
    ylabel('고유값');
    title('스크리 도표');
    grid on;
    
    % 2) 누적 분산 설명률
    subplot(2, 1, 2);
    bar(1:min(10, length(eigenvalues)), cumvar_explained(1:min(10, length(eigenvalues))), 'FaceColor', [0.3, 0.7, 0.9]);
    hold on;
    yline(60, 'r--', '60%', 'LineWidth', 1);
    yline(70, 'g--', '70%', 'LineWidth', 1);
    yline(80, 'b--', '80%', 'LineWidth', 1);
    xlabel('요인 수');
    ylabel('누적 분산 설명률 (%)');
    title('누적 분산 설명률');
    grid on;
    
    % 3) 요인 로딩 히트맵 (개선된)
    if optimal_factors <= 8
        figure('Name', '요인 로딩 히트맵', 'Position', [800, 600, 900, 600]);
        
        % 로딩 매트릭스를 최대 로딩 기준으로 정렬
        [~, sort_order] = sort(max(abs(final_loadings), [], 2), 'descend');
        sorted_loadings = final_loadings(sort_order, :);
        sorted_var_names = all_measure_cols(sort_order);
        
        imagesc(sorted_loadings');
        colorbar;
        colormap(redblue(256));
        caxis([-1, 1]);
        
        ylabel('요인');
        yticks(1:optimal_factors);
        yticklabels(arrayfun(@(x) sprintf('요인 %d', x), 1:optimal_factors, 'UniformOutput', false));
        
        xlabel('측정치 (높은 로딩 순)');
        if length(sorted_var_names) <= 20
            xticks(1:length(sorted_var_names));
            xticklabels(sorted_var_names);
            xtickangle(90);
        else
            xticks(1:5:length(sorted_var_names));
            xticklabels(sorted_var_names(1:5:end));
            xtickangle(90);
        end
        
        title('요인 로딩 히트맵 (정렬됨)');
        
        % 유의한 로딩 강조
        hold on;
        [row, col] = find(abs(sorted_loadings') >= 0.4);
        scatter(col, row, 40, 'k', 'filled', 'MarkerFaceAlpha', 0.8);
        [row, col] = find(abs(sorted_loadings') >= 0.3 & abs(sorted_loadings') < 0.4);
        scatter(col, row, 20, 'k', 'filled', 'MarkerFaceAlpha', 0.5);
    end
    
    % 4) 신뢰도 시각화
    figure('Name', '신뢰도 분석 결과', 'Position', [100, 100, 800, 500]);
    
    subplot(2, 2, 1);
    scale_alphas = [reliability_results.team_downward.alpha, ...
                   reliability_results.diag24_down.alpha, ...
                   reliability_results.diag24_horizontal.alpha];
    bar(scale_alphas, 'FaceColor', [0.4, 0.6, 0.8]);
    hold on;
    yline(0.7, 'r--', '수용 기준', 'LineWidth', 1.5);
    yline(0.8, 'g--', '좋음 기준', 'LineWidth', 1.5);
    xticklabels({'팀장하향', '상진단하향', '상진단수평'});
    ylabel('Cronbach Alpha');
    title('척도별 신뢰도');
    ylim([0, 1]);
    
    subplot(2, 2, 2);
    histogram(communality, 15, 'FaceColor', [0.6, 0.4, 0.8]);
    hold on;
    xline(0.3, 'r--', '최소 기준', 'LineWidth', 1.5);
    xline(0.5, 'g--', '좋음 기준', 'LineWidth', 1.5);
    xlabel('공통성 (h²)');
    ylabel('문항 수');
    title('공통성 분포');
    
    subplot(2, 2, 3);
    if ~all(isnan(factor_reliabilities))
        bar(factor_reliabilities(~isnan(factor_reliabilities)), 'FaceColor', [0.8, 0.6, 0.4]);
        hold on;
        yline(0.7, 'r--', '수용 기준', 'LineWidth', 1.5);
        valid_factors = find(~isnan(factor_reliabilities));
        xticklabels(arrayfun(@(x) sprintf('F%d', x), valid_factors, 'UniformOutput', false));
        ylabel('Cronbach Alpha');
        title('요인별 신뢰도');
        ylim([0, 1]);
    else
        text(0.5, 0.5, '계산 불가', 'HorizontalAlignment', 'center');
        title('요인별 신뢰도');
    end
    
    subplot(2, 2, 4);
    bar(1:optimal_factors, percent_variance, 'FaceColor', [0.2, 0.8, 0.6]);
    xlabel('요인');
    ylabel('분산 설명률 (%)');
    title('요인별 기여도');
    
    %% 25. 최종 결과 저장
    fprintf('\n=== 결과 저장 ===\n');
    
    % 종합 결과 구조체 생성
    comprehensive_results = struct();
    
    % 기본 정보
    comprehensive_results.metadata.total_participants = length(all_participants);
    comprehensive_results.metadata.complete_cases = size(clean_data, 1);
    comprehensive_results.metadata.total_measures = total_measures;
    comprehensive_results.metadata.analysis_date = datestr(now);
    
    % 신뢰도 결과
    comprehensive_results.reliability = reliability_results;
    comprehensive_results.reliability.overall_alpha = overall_alpha;
    
    % 요인분석 결과
    comprehensive_results.factor_analysis.num_factors = optimal_factors;
    comprehensive_results.factor_analysis.loadings = final_loadings;
    comprehensive_results.factor_analysis.factor_scores = factor_scores;
    comprehensive_results.factor_analysis.communality = communality;
    comprehensive_results.factor_analysis.specific_variance = specific_var;
    comprehensive_results.factor_analysis.percent_variance = percent_variance;
    comprehensive_results.factor_analysis.factor_reliabilities = factor_reliabilities;
    
    % 적합도 지수
    comprehensive_results.fit_indices.kmo = kmo_value;
    comprehensive_results.fit_indices.bartlett_p = bartlett_p;
    comprehensive_results.fit_indices.chi2_p = stats.p;
    if exist('rmsea', 'var')
        comprehensive_results.fit_indices.rmsea = rmsea;
    end
    
    % 문제 문항
    comprehensive_results.problematic_items.indices = all_problematic;
    comprehensive_results.problematic_items.names = all_measure_cols(all_problematic);
    
    % 변수 정보
    comprehensive_results.variables.names = all_measure_cols;
    comprehensive_results.variables.participant_ids = participant_ids;
    
    % 워크스페이스에 저장
    save('comprehensive_analysis_results.mat', 'comprehensive_results', 'integrated_table', 'clean_data');
    fprintf('종합 결과가 comprehensive_analysis_results.mat에 저장되었습니다.\n');
    
    % 요인 점수 및 신뢰도 정보를 Excel로 저장
    factor_score_table = array2table([factor_scores, communality], ...
        'VariableNames', [arrayfun(@(x) sprintf('Factor%d', x), 1:optimal_factors, 'UniformOutput', false), {'Communality'}], ...
        'RowNames', participant_ids);
    
    writetable(factor_score_table, 'factor_scores_with_reliability.xlsx', 'WriteRowNames', true);
    
    % 문항별 분석 결과 테이블
    item_analysis_table = table();
    item_analysis_table.Item = all_measure_cols';
    item_analysis_table.Communality = communality;
    item_analysis_table.Specificity = specific_var;
    
    % 각 요인별 로딩 추가
    for f = 1:optimal_factors
        item_analysis_table.(sprintf('Factor%d_Loading', f)) = final_loadings(:, f);
    end
    
    % 최대 로딩 요인 표시
    [max_loadings, primary_factor] = max(abs(final_loadings), [], 2);
    item_analysis_table.Primary_Factor = primary_factor;
    item_analysis_table.Max_Loading = max_loadings;
    
    % 문제 여부 표시
    item_problems = zeros(total_measures, 1);
    item_problems(all_problematic) = 1;
    item_analysis_table.Problematic = logical(item_problems);
    
    writetable(item_analysis_table, 'item_analysis_results.xlsx');
    fprintf('문항 분석 결과가 item_analysis_results.xlsx에 저장되었습니다.\n');
    
catch ME
    fprintf('분석 오류: %s\n', ME.message);
    fprintf('스택 트레이스:\n');
    for i = 1:length(ME.stack)
        fprintf('  %s (라인 %d)\n', ME.stack(i).name, ME.stack(i).line);
    end
    return;
end

fprintf('\n===============================================\n');
fprintf('=== 종합 분석 완료 ===\n');
fprintf('===============================================\n');

fprintf('\n📊 최종 요약:\n');
fprintf('  • 분석 대상: %d명, %d개 문항\n', size(clean_data, 1), total_measures);
fprintf('  • 전체 신뢰도: α = %.3f\n', overall_alpha);
fprintf('  • 최적 요인 수: %d개\n', optimal_factors);
fprintf('  • 총 분산 설명률: %.1f%%\n', sum(percent_variance));
fprintf('  • KMO 측도: %.3f\n', kmo_value);
if ~isempty(all_problematic)
    fprintf('  • 문제 문항: %d개 (%.1f%%)\n', length(all_problematic), length(all_problematic)/total_measures*100);
else
    fprintf('  • 문제 문항: 없음\n');
end

%% ===== 사용자 정의 함수들 =====

function alpha = cronbach_alpha(data)
    % Cronbach's Alpha 계산
    % data: n×k 매트릭스 (n=사례수, k=문항수)
    
    if size(data, 2) < 2
        alpha = NaN;
        return;
    end
    
    k = size(data, 2); % 문항 수
    
    % 각 문항의 분산
    item_variances = var(data, 'omitnan');
    sum_item_var = sum(item_variances);
    
    % 전체 점수의 분산
    total_scores = sum(data, 2, 'omitnan');
    total_var = var(total_scores, 'omitnan');
    
    % Cronbach's Alpha 공식
    alpha = (k / (k - 1)) * (1 - sum_item_var / total_var);
end

function interpretability = calculate_interpretability(loadings)
    % 요인 해석가능성 점수 계산
    % 각 요인별로 적절한 수의 변수가 고르게 로딩되는지 평가
    
    num_factors = size(loadings, 2);
    factor_complexity = zeros(num_factors, 1);
    
    for f = 1:num_factors
        high_loadings = sum(abs(loadings(:, f)) >= 0.4);
        moderate_loadings = sum(abs(loadings(:, f)) >= 0.3 & abs(loadings(:, f)) < 0.4);
        
        % 이상적인 로딩 수는 3-8개
        if high_loadings >= 3 && high_loadings <= 8
            factor_complexity(f) = 1.0;
        elseif high_loadings >= 2 && high_loadings <= 10
            factor_complexity(f) = 0.8;
        elseif high_loadings >= 1
            factor_complexity(f) = 0.5;
        else
            factor_complexity(f) = 0.1;
        end
        
        % 중간 로딩도 고려
        if moderate_loadings >= 1 && moderate_loadings <= 3
            factor_complexity(f) = factor_complexity(f) + 0.1;
        end
    end
    
    interpretability = mean(factor_complexity);
end

function kmo = calculate_kmo(R)
    % KMO(Kaiser-Meyer-Olkin) 측도 계산
    n = size(R, 1);
    kmo = 0;
    
    try
        % 정규화를 위해 대각선을 1로 설정
        R_diag = R;
        R_diag(logical(eye(n))) = 1;
        
        % 편상관행렬 계산
        inv_R = inv(R_diag + eye(n) * 1e-10); % 수치적 안정성을 위한 regularization
        partial_corr = zeros(n);
        
        for i = 1:n
            for j = 1:n
                if i ~= j
                    partial_corr(i, j) = -inv_R(i, j) / sqrt(inv_R(i, i) * inv_R(j, j));
                end
            end
        end
        
        % KMO 계산
        R_sq = R.^2;
        partial_sq = partial_corr.^2;
        
        % 대각선 제외한 원소들만 사용
        off_diag_mask = ~eye(n);
        sum_R = sum(R_sq(off_diag_mask));
        sum_partial = sum(partial_sq(off_diag_mask));
        
        kmo = sum_R / (sum_R + sum_partial);
        
    catch ME
        warning('KMO 계산 실패: %s. 기본값 0.5 사용', ME.message);
        kmo = 0.5;
    end
end

function msa_values = calculate_individual_msa(R)
    % 개별 변수의 MSA (Measure of Sampling Adequacy) 계산
    n = size(R, 1);
    msa_values = zeros(n, 1);
    
    try
        % 정규화를 위해 대각선을 1로 설정
        R_diag = R;
        R_diag(logical(eye(n))) = 1;
        
        inv_R = inv(R_diag + eye(n) * 1e-10);
        
        for i = 1:n
            % i번째 변수에 대한 MSA 계산
            R_i_sq = R(i, :).^2;
            partial_i_sq = zeros(1, n);
            
            for j = 1:n
                if i ~= j
                    partial_i_sq(j) = (-inv_R(i, j) / sqrt(inv_R(i, i) * inv_R(j, j)))^2;
                end
            end
            
            sum_R_i = sum(R_i_sq) - R_i_sq(i); % 자기 자신 제외
            sum_partial_i = sum(partial_i_sq) - partial_i_sq(i);
            
            msa_values(i) = sum_R_i / (sum_R_i + sum_partial_i);
        end
        
    catch ME
        warning('개별 MSA 계산 실패: %s. 기본값 0.5 사용', ME.message);
        msa_values = ones(n, 1) * 0.5;
    end
end

function [chi2, p_value] = bartlett_test(R, n)
    % Bartlett의 구형성 검정
    p = size(R, 1);
    
    try
        % 행렬식 계산 (수치적 안정성 확보)
        det_R = det(R);
        if det_R <= 0 || det_R < 1e-10
            det_R = 1e-10; % 특이행렬 처리
        end
        
        % 검정통계량 계산
        chi2 = -(n - 1 - (2*p + 5)/6) * log(det_R);
        df = p * (p - 1) / 2;
        
        % p-value 계산
        if chi2 < 0
            chi2 = 0;
        end
        p_value = 1 - chi2cdf(chi2, df);
        
    catch ME
        warning('Bartlett 검정 실패: %s', ME.message);
        chi2 = NaN;
        p_value = NaN;
    end
end

function optimal_factors = parallel_analysis_robust(data, n_iterations)
    % 강건한 병렬분석을 통한 최적 요인 수 결정
    [n_obs, n_vars] = size(data);
    
    try
        % 실제 데이터의 고유값
        actual_corr = corr(data, 'Type', 'Pearson');
        actual_eigenvals = eig(actual_corr);
        actual_eigenvals = sort(actual_eigenvals, 'descend');
        
        % 무작위 데이터의 고유값 분포 생성
        random_eigenvals = zeros(n_iterations, n_vars);
        
        fprintf('  병렬분석 진행률: ');
        for iter = 1:n_iterations
            if mod(iter, 100) == 0
                fprintf('%.0f%% ', iter/n_iterations*100);
            end
            
            % 무작위 데이터 생성 (정규분포, 실제 데이터와 유사한 분산)
            random_data = randn(n_obs, n_vars);
            
            % 실제 데이터의 분산 구조를 반영
            for j = 1:n_vars
                random_data(:, j) = random_data(:, j) * std(data(:, j));
            end
            
            try
                random_corr = corr(random_data, 'Type', 'Pearson');
                eigenvals = eig(random_corr);
                random_eigenvals(iter, :) = sort(eigenvals, 'descend');
            catch
                % 상관행렬 계산 실패 시 이전 결과 사용
                if iter > 1
                    random_eigenvals(iter, :) = random_eigenvals(iter-1, :);
                else
                    random_eigenvals(iter, :) = ones(1, n_vars) / n_vars;
                end
            end
        end
        fprintf('완료\n');
        
        % 95백분위수 계산
        random_95th = prctile(random_eigenvals, 95, 1);
        
        % 실제 고유값이 무작위보다 큰 요인 수 결정
        optimal_factors = sum(actual_eigenvals > random_95th');
        
        % 합리적 범위로 제한
        optimal_factors = max(1, min(optimal_factors, floor(n_vars/3)));
        
    catch ME
        warning('병렬분석 실패: %s. Kaiser 기준 사용', ME.message);
        optimal_factors = sum(actual_eigenvals > 1);
    end
end
function [optimal_factors, map_values] = velicer_map(R, maxFactors)
% Velicer의 MAP (Minimum Average Partial) 기준을 PCA 기반으로 계산
% R : 상관행렬(또는 공분산행렬)
% maxFactors : 최대 검토할 성분 수 (기본: min(8, n-1))
% 반환:
%   optimal_factors : MAP 값을 최소화하는 성분 수 m
%   map_values      : m=0..maxFactors에 대한 MAP 값 벡터

    %----- 입력 및 전처리 -----
    if nargin < 2
        maxFactors = min(8, size(R,1)-1);
    end
    R = (R + R')/2;                     % 대칭화
    n = size(R,1);
    if size(R,2) ~= n
        error('R은 정사각 행렬이어야 합니다.');
    end
    % 공분산 -> 상관행렬 변환 (상관행렬이면 변화 없음)
    s = sqrt(diag(R));
    if any(s == 0)
        error('분산 0인 변수가 있습니다.');
    end
    Dinv = diag(1./s);
    R = Dinv * R * Dinv;                % 이제 R은 상관행렬
    
    %----- 고유분해 (PCA on correlation) -----
    R = (R + R')/2;                     % 수치 안정화
    [V,D] = eig(R);
    [lambda, idx] = sort(diag(D), 'descend');
    V = V(:, idx);
    % 수치 오차 보정 (음수 고유값 0으로 클리핑)
    lambda = max(lambda, 0);
    
    %----- MAP 값 계산: m = 0..maxFactors -----
    maxFactors = min(maxFactors, n-1);
    map_values = zeros(maxFactors+1, 1);
    denom = n*(n-1); % 오프대각 원소 개수
    
    for m = 0:maxFactors
        if m == 0
            Rhat = zeros(n);
        else
            Lm = V(:,1:m) * diag(sqrt(lambda(1:m)));  % PCA 로딩 (V*sqrt(λ))
            Rhat = Lm * Lm.';                         % = V(:,1:m)*diag(λ)*V(:,1:m)'
        end
        Resid = R - Rhat;
        % 대각 0으로
        Resid(1:n+1:end) = 0;
        % 평균 제곱(오프대각) - Velicer MAP
        map_values(m+1) = sum(Resid(:).^2) / denom;
    end
    
    %----- 최소 MAP을 주는 m 선택 -----
    [~, idxMin] = min(map_values);
    optimal_factors = idxMin - 1;  % m = 0..maxFactors
    
end


function [h, p] = swtest(x)
    % 간단한 Shapiro-Wilk 정규성 검정 (근사적)
    % 표본 크기가 작을 때 (n < 5000)만 사용
    
    x = x(~isnan(x)); % NaN 제거
    n = length(x);
    
    if n < 3 || n > 5000
        h = NaN;
        p = NaN;
        return;
    end
    
    try
        % 데이터 정렬
        x_sorted = sort(x);
        
        % Shapiro-Wilk 통계량 계산 (간소화된 버전)
        % 실제로는 더 복잡한 계산이 필요하지만, 근사적 방법 사용
        
        % 표준화
        x_std = (x_sorted - mean(x_sorted)) / std(x_sorted);
        
        % Q-Q plot 기울기 기반 근사 통계량
        theoretical_quantiles = norminv((1:n) / (n + 1));
        
        % 상관계수 기반 근사 W 통계량
        W = corr(x_std, theoretical_quantiles')^2;
        
        % p-value 근사 (경험적 공식)
        if n <= 50
            % 작은 표본용 근사
            p = 1 - normcdf((W - 0.5) * sqrt(n));
        else
            % 큰 표본용 근사
            p = 1 - normcdf((W - 0.8) * sqrt(n) * 2);
        end
        
        % p-value 범위 제한
        p = max(0, min(1, p));
        
        % 귀무가설 기각 여부 (α = 0.05)
        h = p < 0.05;
        
    catch ME
        warning('Shapiro-Wilk 검정 실패: %s', ME.message);
        h = NaN;
        p = NaN;
    end
end

function cmap = redblue(m)
    % 빨강-하양-파랑 컬러맵 생성 (요인 로딩 시각화용)
    if nargin < 1
        m = 256;
    end
    
    if mod(m, 2) == 0
        % 짝수인 경우
        m1 = m/2;
        r = [linspace(0, 1, m1), ones(1, m1)];
        g = [linspace(0, 1, m1), linspace(1, 0, m1)];
        b = [ones(1, m1), linspace(1, 0, m1)];
    else
        % 홀수인 경우
        m1 = floor(m/2);
        r = [linspace(0, 1, m1), 1, ones(1, m1)];
        g = [linspace(0, 1, m1), 1, linspace(1, 0, m1)];
        b = [ones(1, m1), 1, linspace(1, 0, m1)];
    end
    
    cmap = [r', g', b'];
end

% ===== 추가 분석 함수들 =====

function reliability_stats = advanced_reliability_analysis(data, item_names)
    % 고급 신뢰도 분석 (Split-half, Guttman 등)
    
    n_items = size(data, 2);
    reliability_stats = struct();
    
    % 1) Cronbach's Alpha
    reliability_stats.cronbach_alpha = cronbach_alpha(data);
    
    % 2) Split-half 신뢰도 (홀수/짝수 분할)
    odd_items = data(:, 1:2:end);
    even_items = data(:, 2:2:end);
    
    if size(odd_items, 2) >= 2 && size(even_items, 2) >= 2
        odd_scores = sum(odd_items, 2);
        even_scores = sum(even_items, 2);
        split_half_corr = corr(odd_scores, even_scores);
        
        % Spearman-Brown 공식으로 보정
        reliability_stats.split_half = (2 * split_half_corr) / (1 + split_half_corr);
    else
        reliability_stats.split_half = NaN;
    end
    
    % 3) Guttman's Lambda-6 (Alpha의 하한값)
    item_vars = var(data);
    total_var = var(sum(data, 2));
    smc_sum = 0; % Squared Multiple Correlations 합
    
    for i = 1:n_items
        other_items = data(:, setdiff(1:n_items, i));
        if size(other_items, 2) > 0
            try
                mdl = fitlm(other_items, data(:, i));
                smc_sum = smc_sum + mdl.Rsquared.Ordinary * item_vars(i);
            catch
                smc_sum = smc_sum + 0; % 회귀 실패 시 0으로 처리
            end
        end
    end
    
    reliability_stats.guttman_lambda6 = (n_items / (n_items - 1)) * (1 - (sum(item_vars) - smc_sum) / total_var);
    
    % 4) McDonald's Omega (요인분석 기반)
    try
        if n_items >= 3
            [loadings, ~] = factoran(data, 1); % 1요인 모델
            sum_loadings = sum(loadings)^2;
            sum_uniqueness = sum(1 - loadings.^2);
            reliability_stats.mcdonalds_omega = sum_loadings / (sum_loadings + sum_uniqueness);
        else
            reliability_stats.mcdonalds_omega = NaN;
        end
    catch
        reliability_stats.mcdonalds_omega = NaN;
    end
end

function [outliers, outlier_indices] = detect_multivariate_outliers(data, method)
    % 다변량 이상치 검출
    % method: 'mahalanobis' 또는 'robust'
    
    if nargin < 2
        method = 'mahalanobis';
    end
    
    [n, p] = size(data);
    outliers = false(n, 1);
    outlier_indices = [];
    
    try
        switch lower(method)
            case 'mahalanobis'
                % 마할라노비스 거리 기반
                mu = mean(data);
                sigma = cov(data);
                
                % 특이행렬 처리
                if rank(sigma) < p
                    sigma = sigma + eye(p) * 1e-6;
                end
                
                mahal_dist = zeros(n, 1);
                for i = 1:n
                    diff = data(i, :) - mu;
                    mahal_dist(i) = sqrt(diff * (sigma \ diff'));
                end
                
                % 카이제곱 분포 기준 (p < 0.001)
                cutoff = sqrt(chi2inv(0.999, p));
                outliers = mahal_dist > cutoff;
                outlier_indices = find(outliers);
                
            case 'robust'
                % 로버스트 방법 (MCD 추정량 근사)
                % 간단한 근사: 중앙값과 MAD 기반
                medians = median(data);
                mads = mad(data, 1); % MAD (Median Absolute Deviation)
                
                robust_dist = zeros(n, 1);
                for i = 1:n
                    std_diff = abs(data(i, :) - medians) ./ mads;
                    robust_dist(i) = sqrt(sum(std_diff.^2));
                end
                
                cutoff = sqrt(chi2inv(0.999, p));
                outliers = robust_dist > cutoff;
                outlier_indices = find(outliers);
        end
        
    catch ME
        warning('이상치 검출 실패: %s', ME.message);
    end
end

function quality_score = assess_factor_quality(loadings, communality)
    % 요인 품질 종합 평가
    
    [n_vars, n_factors] = size(loadings);
    
    % 1) 단순 구조 점수 (각 변수가 하나의 요인에만 높게 로딩)
    simple_structure = 0;
    for i = 1:n_vars
        high_loadings = sum(abs(loadings(i, :)) >= 0.4);
        if high_loadings == 1
            simple_structure = simple_structure + 1;
        elseif high_loadings == 0
            simple_structure = simple_structure + 0.5; % 어느 요인에도 로딩되지 않음
        end
    end
    simple_structure_score = simple_structure / n_vars;
    
    % 2) 요인별 변수 분포 균형성
    factor_sizes = zeros(n_factors, 1);
    for f = 1:n_factors
        factor_sizes(f) = sum(abs(loadings(:, f)) >= 0.4);
    end
    
    % 이상적인 요인 크기는 3-7개 변수
    balance_score = 0;
    for f = 1:n_factors
        if factor_sizes(f) >= 3 && factor_sizes(f) <= 7
            balance_score = balance_score + 1;
        elseif factor_sizes(f) >= 2 && factor_sizes(f) <= 10
            balance_score = balance_score + 0.7;
        elseif factor_sizes(f) >= 1
            balance_score = balance_score + 0.3;
        end
    end
    balance_score = balance_score / n_factors;
    
    % 3) 전체 공통성 수준
    communality_score = mean(communality);
    
    % 종합 품질 점수
    quality_score = simple_structure_score * 0.4 + balance_score * 0.3 + communality_score * 0.3;
end

function summary_report = generate_analysis_summary(comprehensive_results)
    % 분석 결과 요약 리포트 생성
    
    summary_report = struct();
    
    % 기본 정보
    summary_report.sample_size = comprehensive_results.metadata.complete_cases;
    summary_report.total_items = comprehensive_results.metadata.total_measures;
    
    % 신뢰도 요약
    summary_report.overall_reliability = comprehensive_results.reliability.overall_alpha;
    
    scale_alphas = [comprehensive_results.reliability.team_downward.alpha, ...
                   comprehensive_results.reliability.diag24_down.alpha, ...
                   comprehensive_results.reliability.diag24_horizontal.alpha];
    summary_report.scale_reliabilities = scale_alphas;
    summary_report.min_scale_reliability = min(scale_alphas);
    
    % 요인분석 요약
    summary_report.num_factors = comprehensive_results.factor_analysis.num_factors;
    summary_report.total_variance_explained = sum(comprehensive_results.factor_analysis.percent_variance);
    summary_report.avg_communality = mean(comprehensive_results.factor_analysis.communality);
    
    % 적합도 요약
    summary_report.kmo = comprehensive_results.fit_indices.kmo;
    summary_report.bartlett_significant = comprehensive_results.fit_indices.bartlett_p < 0.001;
    
    % 문제 항목 요약
    summary_report.num_problematic_items = length(comprehensive_results.problematic_items.indices);
    summary_report.percent_problematic = summary_report.num_problematic_items / summary_report.total_items * 100;
    
    % 전체 품질 등급
    if summary_report.overall_reliability >= 0.9 && summary_report.kmo >= 0.8 && summary_report.percent_problematic < 10
        summary_report.overall_quality = 'Excellent';
    elseif summary_report.overall_reliability >= 0.8 && summary_report.kmo >= 0.7 && summary_report.percent_problematic < 20
        summary_report.overall_quality = 'Good';
    elseif summary_report.overall_reliability >= 0.7 && summary_report.kmo >= 0.6 && summary_report.percent_problematic < 30
        summary_report.overall_quality = 'Acceptable';
    else
        summary_report.overall_quality = 'Needs Improvement';
    end
end

fprintf('\n\n🎯 분석 완료! 결과 파일들을 확인하세요:\n');
fprintf('  • comprehensive_analysis_results.mat (전체 결과)\n');
fprintf('  • factor_scores_with_reliability.xlsx (요인 점수)\n');
fprintf('  • item_analysis_results.xlsx (문항별 분석)\n');
fprintf('\n💡 다음 단계 권장사항:\n');
fprintf('  1) 문제 문항 검토 및 수정\n');
fprintf('  2) 요인 명명 및 이론적 해석\n');
fprintf('  3) 확인적 요인분석 수행 (별도 표본)\n');
fprintf('  4) 준거 타당도 검증\n');