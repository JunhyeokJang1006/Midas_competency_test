%% 개선된 성과 데이터 통합 및 요인분석 스크립트
% 강화된 데이터 전처리 (Z-score, Winsorization) 및 RMSEA 문제 해결
% 작성자: 데이터 분석팀
% 날짜: 2025년 2월

clear; clc; close all

%% 1. 데이터 파일 읽기
filename='C:\Users\MIDASIT\Neurocompetency_project\Competency\(25.02.05) 과거 준거_보정완료 1.xlsx';

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

%% 2. 데이터 구조 확인 및 측정치 컬럼 식별
fprintf('\n=== 데이터 구조 확인 ===\n');
fprintf('팀장하향평가: %d명, %d개 컬럼\n', height(team_downward), width(team_downward));
fprintf('24상진단_하향: %d명, %d개 컬럼\n', height(diag24_down), width(diag24_down));
fprintf('24상진단_수평: %d명, %d개 컬럼\n', height(diag24_horizontal), width(diag24_horizontal));

% 측정치 컬럼 식별
team_cols = team_downward.Properties.VariableNames;
team_measure_cols = team_cols(startsWith(team_cols, 'edit_q'));

diag_down_cols = diag24_down.Properties.VariableNames;
diag_down_measure_cols = diag_down_cols(startsWith(diag_down_cols, 'edit_LINKED'));

diag_horizontal_cols = diag24_horizontal.Properties.VariableNames;
diag_horizontal_measure_cols = diag_horizontal_cols(startsWith(diag_horizontal_cols, 'edit_LINKED'));

fprintf('\n=== 측정치 컬럼 수 ===\n');
fprintf('팀장하향평가: %d개 문항\n', length(team_measure_cols));
fprintf('24상진단_하향: %d개 문항\n', length(diag_down_measure_cols));
fprintf('24상진단_수평: %d개 문항\n', length(diag_horizontal_measure_cols));

%% 3. 통합 데이터셋 생성 (개선된 버전)
fprintf('\n=== 통합 데이터셋 생성 ===\n');

all_participants = unique([team_downward.MR_CODE; diag24_down.MR_CODE; diag24_horizontal.MR_CODE]);
fprintf('총 참가자 수: %d명\n', length(all_participants));

% 컬럼명 생성
team_col_names = strcat('TeamDown_', team_measure_cols);
diag_down_col_names = strcat('Diag24Down_', diag_down_measure_cols);
diag_horizontal_col_names = strcat('Diag24Horiz_', diag_horizontal_measure_cols);
all_measure_cols = [team_col_names, diag_down_col_names, diag_horizontal_col_names];
total_measures = length(all_measure_cols);

% 통합 데이터 매트릭스 초기화
integrated_data = nan(length(all_participants), total_measures);
participant_map = containers.Map(all_participants, 1:length(all_participants));

% 데이터 통합 과정
fprintf('데이터 통합 중...\n');
for i = 1:height(team_downward)
    participant_id = team_downward.MR_CODE{i};
    if isKey(participant_map, participant_id)
        row_idx = participant_map(participant_id);
        for j = 1:length(team_measure_cols)
            value = team_downward{i, team_measure_cols{j}};
            if isnumeric(value) && ~isnan(value)
                integrated_data(row_idx, j) = value;
            end
        end
    end
end

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

% 통합 테이블 생성
integrated_table = array2table(integrated_data, ...
    'VariableNames', all_measure_cols, ...
    'RowNames', all_participants);

fprintf('통합 완료! 최종 크기: %d명 × %d개 측정치\n', height(integrated_table), width(integrated_table));

%% 4. 강화된 데이터 전처리
fprintf('\n\n=== 강화된 데이터 전처리 ===\n');

data_matrix = table2array(integrated_table);
[n_obs, n_vars] = size(data_matrix);

% Step 1: NaN이 있는 변수/참가자 제거 전략
fprintf('1단계: 결측치 분석\n');
missing_by_var = sum(isnan(data_matrix), 1) / n_obs * 100;
missing_by_case = sum(isnan(data_matrix), 2) / n_vars * 100;

fprintf('  변수별 결측률: %.1f%% ~ %.1f%%\n', min(missing_by_var), max(missing_by_var));
fprintf('  케이스별 결측률: %.1f%% ~ %.1f%%\n', min(missing_by_case), max(missing_by_case));

% 결측률 50% 이상인 변수 제거
high_missing_vars = find(missing_by_var > 50);
if ~isempty(high_missing_vars)
    fprintf('  결측률 50%% 초과 변수 %d개 제거\n', length(high_missing_vars));
    data_matrix(:, high_missing_vars) = [];
    all_measure_cols(high_missing_vars) = [];
    missing_by_var(high_missing_vars) = [];
end

% 결측률 50% 이상인 케이스 제거
high_missing_cases = find(missing_by_case > 50);
if ~isempty(high_missing_cases)
    fprintf('  결측률 50%% 초과 케이스 %d개 제거\n', length(high_missing_cases));
    data_matrix(high_missing_cases, :) = [];
    participant_ids_clean = all_participants;
    participant_ids_clean(high_missing_cases) = [];
else
    participant_ids_clean = all_participants;
end

% Step 2: 완전한 케이스만 선택 (listwise deletion)
fprintf('\n2단계: 완전한 케이스 선택\n');
complete_cases = ~any(isnan(data_matrix), 2);
clean_data = data_matrix(complete_cases, :);
participant_ids = participant_ids_clean(complete_cases);

fprintf('  완전한 케이스: %d명 (%.1f%%)\n', size(clean_data, 1), ...
    100 * size(clean_data, 1) / length(participant_ids_clean));
fprintf('  최종 분석 데이터: %d × %d\n', size(clean_data));

if size(clean_data, 1) < 100
    warning('표본 크기가 너무 작습니다 (n=%d). 분석 결과의 안정성이 낮을 수 있습니다.', size(clean_data, 1));
end

% Step 3: 이상치 검출 및 처리 (Winsorization)
fprintf('\n3단계: 이상치 검출 및 Winsorization\n');
winsorized_data = clean_data;
outlier_summary = zeros(size(clean_data, 2), 3); % [변수, 하한이상치수, 상한이상치수]

for j = 1:size(clean_data, 2)
    var_data = clean_data(:, j);
    
    % Z-score 기반 이상치 (|Z| > 3.29, p < 0.001)
    z_scores = abs(zscore(var_data));
    z_outliers = z_scores > 3.29;
    
    % IQR 기반 이상치 (더 엄격한 기준)
    q1 = prctile(var_data, 25);
    q3 = prctile(var_data, 75);
    iqr = q3 - q1;
    lower_bound = q1 - 2.5 * iqr;  % 더 엄격한 2.5*IQR
    upper_bound = q3 + 2.5 * iqr;
    
    iqr_outliers = var_data < lower_bound | var_data > upper_bound;
    
    % 두 방법 중 하나라도 이상치로 판별되면 처리
    all_outliers = z_outliers | iqr_outliers;
    
    if any(all_outliers)
        % Winsorization (5th와 95th percentile로 절단)
        p5 = prctile(var_data, 5);
        p95 = prctile(var_data, 95);
        
        lower_outliers = var_data < p5;
        upper_outliers = var_data > p95;
        
        winsorized_data(lower_outliers, j) = p5;
        winsorized_data(upper_outliers, j) = p95;
        
        outlier_summary(j, :) = [j, sum(lower_outliers), sum(upper_outliers)];
    end
end

% 이상치 처리 요약
total_outliers_processed = sum(outlier_summary(:, 2:3), 'all');
vars_with_outliers = sum(any(outlier_summary(:, 2:3), 2));

fprintf('  이상치 처리 완료:\n');
fprintf('    - 처리된 변수: %d개\n', vars_with_outliers);
fprintf('    - 총 처리된 이상치: %d개\n', total_outliers_processed);
fprintf('    - Winsorization 범위: 5th-95th percentile\n');

% Step 4: 데이터 표준화 (선택적)
fprintf('\n4단계: 데이터 품질 최종 점검\n');

% 분산이 0인 변수 확인
zero_var_cols = var(winsorized_data) < 1e-10;
if any(zero_var_cols)
    fprintf('  분산이 0인 변수 %d개 제거\n', sum(zero_var_cols));
    winsorized_data(:, zero_var_cols) = [];
    all_measure_cols(zero_var_cols) = [];
end

% 최종 데이터
final_data = winsorized_data;
final_var_names = all_measure_cols;

fprintf('  최종 분석용 데이터: %d × %d\n', size(final_data));
fprintf('  평균 범위: %.2f ~ %.2f\n', min(mean(final_data)), max(mean(final_data)));
fprintf('  표준편차 범위: %.2f ~ %.2f\n', min(std(final_data)), max(std(final_data)));

%% 5. 개선된 신뢰도 분석
fprintf('\n\n=== 신뢰도 분석 ===\n');

% 척도별 인덱스 재계산
n_team = sum(contains(final_var_names, 'TeamDown'));
n_diag_down = sum(contains(final_var_names, 'Diag24Down'));
n_diag_horiz = sum(contains(final_var_names, 'Diag24Horiz'));

if n_team > 0
    team_indices = find(contains(final_var_names, 'TeamDown'));
    team_alpha = cronbach_alpha(final_data(:, team_indices));
    fprintf('팀장하향평가 (n=%d): α = %.3f\n', n_team, team_alpha);
end

if n_diag_down > 0
    diag_down_indices = find(contains(final_var_names, 'Diag24Down'));
    diag_down_alpha = cronbach_alpha(final_data(:, diag_down_indices));
    fprintf('24상진단_하향 (n=%d): α = %.3f\n', n_diag_down, diag_down_alpha);
end

if n_diag_horiz > 0
    diag_horiz_indices = find(contains(final_var_names, 'Diag24Horiz'));
    diag_horiz_alpha = cronbach_alpha(final_data(:, diag_horiz_indices));
    fprintf('24상진단_수평 (n=%d): α = %.3f\n', n_diag_horiz, diag_horiz_alpha);
end

overall_alpha = cronbach_alpha(final_data);
fprintf('전체 척도: α = %.3f\n', overall_alpha);

%% 6. 요인분석 적합성 재검정 (개선된 버전)
fprintf('\n=== 요인분석 적합성 검정 ===\n');

% 표본 충분성
sample_to_var_ratio = size(final_data, 1) / size(final_data, 2);
fprintf('표본/변수 비율: %.1f:1 ', sample_to_var_ratio);
if sample_to_var_ratio >= 10
    fprintf('(매우 좋음)\n');
elseif sample_to_var_ratio >= 5
    fprintf('(좋음)\n');
else
    fprintf('(부족 - 주의 필요)\n');
end

% 상관행렬 계산 및 조건수 확인
correlation_matrix = corr(final_data, 'Type', 'Pearson');
condition_number = cond(correlation_matrix);
fprintf('상관행렬 조건수: %.2e ', condition_number);
if condition_number < 1000
    fprintf('(좋음)\n');
elseif condition_number < 10000
    fprintf('(보통)\n');
else
    fprintf('(높음 - 다중공선성 의심)\n');
end

% KMO 측도
kmo_value = calculate_kmo_robust(correlation_matrix);
fprintf('KMO 측도: %.3f ', kmo_value);
if kmo_value >= 0.8
    fprintf('(매우 좋음)\n');
elseif kmo_value >= 0.7
    fprintf('(좋음)\n');
elseif kmo_value >= 0.6
    fprintf('(보통)\n');
else
    fprintf('(부적절)\n');
end

% Bartlett 구형성 검정
[bartlett_chi2, bartlett_p] = bartlett_test_robust(correlation_matrix, size(final_data, 1));
fprintf('Bartlett 검정: χ² = %.2f, p = %.6f ', bartlett_chi2, bartlett_p);
if bartlett_p < 0.001
    fprintf('(적합)\n');
else
    fprintf('(부적합)\n');
end

%% 7. 최적 요인 수 결정 (개선된 방법)
fprintf('\n=== 최적 요인 수 결정 ===\n');
max_factors = min(8, size(final_data, 2) - 1);

% 고유값 분해
eigenvalues = eig(correlation_matrix);
eigenvalues = sort(eigenvalues, 'descend');

% 1) Kaiser 기준
kaiser_factors = sum(eigenvalues > 1);
fprintf('Kaiser 기준 (고유값 > 1): %d개\n', kaiser_factors);

% 2) 분산 설명률 기준
cumvar = cumsum(eigenvalues) / sum(eigenvalues) * 100;
factors_70 = find(cumvar >= 70, 1);
fprintf('70%% 분산 설명: %d개 (%.1f%%)\n', factors_70, cumvar(factors_70));

% 3) 개선된 평행분석
fprintf('평행분석 수행 중...\n');
parallel_factors = parallel_analysis_improved(final_data, 500);
fprintf('평행분석 권장: %d개\n', parallel_factors);

% 4) MAP 기준
map_factors = velicer_map_robust(correlation_matrix);
fprintf('MAP 기준: %d개\n', map_factors);

% 통합 권장사항
criteria = [kaiser_factors, parallel_factors, map_factors, factors_70];
weights = [0.2, 0.3, 0.3, 0.2];
recommended_factors = round(sum(criteria .* weights));
recommended_factors = max(2, min(recommended_factors, 6)); % 2-6개 범위로 제한

fprintf('통합 권장: %d개 요인\n', recommended_factors);

%% 8. 요인분석 수행 (개선된 RMSEA 계산)
fprintf('\n=== 요인분석 수행 ===\n');

try
    % Maximum likelihood 방법으로 요인분석 수행
    [loadings, specific_var, T, stats] = factoran(final_data, recommended_factors, ...
        'rotate', 'varimax', 'scores', 'regression', 'maxit', 1000);
    
    fprintf('요인분석 성공! %d개 요인 추출\n', recommended_factors);
    
    % 개선된 적합도 지수 계산
    fit_indices = calculate_fit_indices_robust(final_data, loadings, specific_var, stats);
    
    % 결과 출력
    fprintf('\n=== 적합도 지수 ===\n');
    if isfield(fit_indices, 'chi2') && ~isnan(fit_indices.chi2)
        fprintf('χ² = %.2f (df = %d, p = %.4f)\n', ...
            fit_indices.chi2, fit_indices.df, fit_indices.p_value);
    end
    
    if isfield(fit_indices, 'rmsea') && ~isnan(fit_indices.rmsea)
        fprintf('RMSEA = %.3f ', fit_indices.rmsea);
        if fit_indices.rmsea <= 0.05
            fprintf('(매우 좋음)\n');
        elseif fit_indices.rmsea <= 0.08
            fprintf('(수용가능)\n');
        else
            fprintf('(개선 필요)\n');
        end
        
        if isfield(fit_indices, 'rmsea_ci')
            fprintf('RMSEA 90%% CI: [%.3f, %.3f]\n', fit_indices.rmsea_ci(1), fit_indices.rmsea_ci(2));
        end
    else
        fprintf('RMSEA 계산 불가 (완벽 적합 또는 수치 문제)\n');
    end
    
    % 추가 적합도 지수들
    if isfield(fit_indices, 'cfi')
        fprintf('CFI = %.3f\n', fit_indices.cfi);
    end
    if isfield(fit_indices, 'tli')
        fprintf('TLI = %.3f\n', fit_indices.tli);
    end
    if isfield(fit_indices, 'srmr')
        fprintf('SRMR = %.3f\n', fit_indices.srmr);
    end
    
    % 공통성 분석
    communality = sum(loadings.^2, 2);
    fprintf('\n공통성 통계:\n');
    fprintf('  평균: %.3f, 범위: [%.3f, %.3f]\n', mean(communality), min(communality), max(communality));
    
    low_communality = sum(communality < 0.3);
    if low_communality > 0
        fprintf('  낮은 공통성 변수 (< 0.3): %d개\n', low_communality);
    end
    
catch ME
    fprintf('요인분석 실패: %s\n', ME.message);
    
    % 대안: 주성분분석
    fprintf('대안으로 주성분분석 수행\n');
    [coeff, score, latent] = pca(final_data);
    
    loadings = coeff(:, 1:recommended_factors) * diag(sqrt(latent(1:recommended_factors)));
    communality = sum(loadings.^2, 2);
    specific_var = 1 - communality;
    
    fprintf('주성분분석 완료 (%d개 성분)\n', recommended_factors);
    fprintf('총 분산 설명률: %.1f%%\n', sum(latent(1:recommended_factors))/sum(latent)*100);
end

%% 9. 요인 로딩 매트릭스 출력
fprintf('\n=== 요인 로딩 매트릭스 ===\n');
fprintf('%-30s', '측정치');
for f = 1:recommended_factors
    fprintf('%10s', sprintf('F%d', f));
end
fprintf('%10s\n', 'h²');
fprintf('%s\n', repmat('-', 1, 30 + 10*recommended_factors + 10));

% 최대 로딩 순으로 정렬
[~, sort_order] = sort(max(abs(loadings), [], 2), 'descend');

for idx = 1:min(20, length(sort_order)) % 상위 20개만 표시
    i = sort_order(idx);
    var_name = final_var_names{i};
    if length(var_name) > 29
        var_name = [var_name(1:26), '...'];
    end
    fprintf('%-30s', var_name);
    
    for f = 1:recommended_factors
        loading = loadings(i, f);
        if abs(loading) >= 0.4
            fprintf('%9.3f*', loading);
        elseif abs(loading) >= 0.3
            fprintf('%10.3f', loading);
        else
            fprintf('%10s', '');
        end
    end
    
    fprintf('%10.3f\n', communality(i));
end

if length(sort_order) > 20
    fprintf('... (나머지 %d개 변수 생략)\n', length(sort_order) - 20);
end

%% 10. 결과 저장
fprintf('\n=== 결과 저장 ===\n');

% 종합 결과 구조체
analysis_results = struct();
analysis_results.metadata.sample_size = size(final_data, 1);
analysis_results.metadata.n_variables = size(final_data, 2);
analysis_results.metadata.outliers_processed = total_outliers_processed;

analysis_results.reliability.overall_alpha = overall_alpha;
if exist('team_alpha', 'var')
    analysis_results.reliability.team_alpha = team_alpha;
end
if exist('diag_down_alpha', 'var')
    analysis_results.reliability.diag_down_alpha = diag_down_alpha;
end
if exist('diag_horiz_alpha', 'var')
    analysis_results.reliability.diag_horiz_alpha = diag_horiz_alpha;
end

analysis_results.factor_analysis.n_factors = recommended_factors;
analysis_results.factor_analysis.loadings = loadings;
analysis_results.factor_analysis.communality = communality;
analysis_results.factor_analysis.specific_variance = specific_var;

analysis_results.fit_indices = fit_indices;
analysis_results.variables = final_var_names;
analysis_results.participant_ids = participant_ids;

% 파일 저장
save('enhanced_analysis_results.mat', 'analysis_results', 'final_data', 'final_var_names');

% 요인 점수 계산 및 Excel 저장
if exist('loadings', 'var')
    try
        factor_scores = final_data * pinv(loadings');
        
        factor_table = array2table([factor_scores, communality], ...
            'VariableNames', [arrayfun(@(x) sprintf('Factor%d', x), 1:recommended_factors, 'UniformOutput', false), {'Communality'}], ...
            'RowNames', participant_ids);
        
        writetable(factor_table, 'enhanced_factor_results.xlsx', 'WriteRowNames', true);
        fprintf('결과가 enhanced_factor_results.xlsx에 저장되었습니다.\n');
    catch
        fprintf('Excel 저장 실패 - MAT 파일만 저장됨\n');
    end
end

fprintf('분석 완료!\n');
%% 11. 요인별 세부 분석 및 커스텀 요인점수 산출
fprintf('\n\n=== 요인별 세부 분석 및 커스텀 요인점수 산출 ===\n');

%% 11.1 각 요인의 주요 문항 식별
factor_interpretation = cell(recommended_factors, 1);
factor_items = cell(recommended_factors, 1);
factor_loadings_detail = cell(recommended_factors, 1);

fprintf('\n--- 요인별 주요 문항 분석 ---\n');

for f = 1:recommended_factors
    fprintf('\n=== 요인 %d ===\n', f);
    
    % 해당 요인에 높게 로딩되는 문항들 찾기 (절댓값 0.4 이상)
    high_loadings = abs(loadings(:, f)) >= 0.4;
    factor_items{f} = find(high_loadings);
    factor_loadings_detail{f} = loadings(high_loadings, f);
    
    if isempty(factor_items{f})
        % 0.4 이상이 없으면 기준을 0.3으로 낮춤
        high_loadings = abs(loadings(:, f)) >= 0.3;
        factor_items{f} = find(high_loadings);
        factor_loadings_detail{f} = loadings(high_loadings, f);
        fprintf('주요 문항 (로딩 ≥ 0.3): %d개\n', length(factor_items{f}));
    else
        fprintf('주요 문항 (로딩 ≥ 0.4): %d개\n', length(factor_items{f}));
    end
    
    % 주요 문항들과 로딩값 출력
    [sorted_loadings, sort_idx] = sort(abs(factor_loadings_detail{f}), 'descend');
    for i = 1:min(5, length(factor_items{f})) % 상위 5개만 표시
        item_idx = factor_items{f}(sort_idx(i));
        item_name = final_var_names{item_idx};
        loading_value = factor_loadings_detail{f}(sort_idx(i));
        
        % 문항명 축약 (너무 길면)
        if length(item_name) > 50
            item_name = [item_name(1:47), '...'];
        end
        
        fprintf('  %2d. %-50s (%.3f)\n', i, item_name, loading_value);
    end
    
    if length(factor_items{f}) > 5
        fprintf('  ... 기타 %d개 문항\n', length(factor_items{f}) - 5);
    end
    
    % 요인 해석을 위한 프롬프트
    fprintf('  >> 이 요인의 의미를 입력하세요 (예: 리더십역량, 의사소통능력 등):\n');
    factor_name = input('     요인명: ', 's');
    if isempty(factor_name)
        factor_name = sprintf('Factor_%d', f);
    end
    factor_interpretation{f} = factor_name;
end

%% 11.2 사용자 정의 요인점수 계산 방법 선택
fprintf('\n\n--- 요인점수 계산 방법 선택 ---\n');
fprintf('1. 단순합계 (Simple Sum): 선택된 문항들의 단순 합계\n');
fprintf('2. 평균점수 (Mean Score): 선택된 문항들의 평균\n');
fprintf('3. 가중합계 (Weighted Sum): 요인로딩을 가중치로 사용한 가중합계\n');
fprintf('4. 표준화 점수 (Standardized): Z-score로 표준화된 점수\n');
fprintf('5. 회귀점수 (Regression): 회귀분석 기반 요인점수\n');

scoring_method = input('계산 방법을 선택하세요 (1-5): ');
if isempty(scoring_method) || scoring_method < 1 || scoring_method > 5
    scoring_method = 2; % 기본값: 평균점수
    fprintf('기본값으로 평균점수 방법을 사용합니다.\n');
end

%% 11.3 커스텀 요인점수 계산
factor_scores_custom = zeros(size(final_data, 1), recommended_factors);
factor_score_details = struct();

fprintf('\n--- 커스텀 요인점수 계산 ---\n');

for f = 1:recommended_factors
    factor_name = factor_interpretation{f};
    selected_items = factor_items{f};
    selected_data = final_data(:, selected_items);
    
    fprintf('\n%s (요인 %d):\n', factor_name, f);
    fprintf('  사용 문항 수: %d개\n', length(selected_items));
    
    switch scoring_method
        case 1 % 단순합계
            factor_scores_custom(:, f) = sum(selected_data, 2);
            method_name = 'Simple Sum';
            
        case 2 % 평균점수
            factor_scores_custom(:, f) = mean(selected_data, 2);
            method_name = 'Mean Score';
            
        case 3 % 가중합계
            weights = abs(factor_loadings_detail{f});
            weights = weights / sum(weights); % 정규화
            factor_scores_custom(:, f) = selected_data * weights;
            method_name = 'Weighted Sum';
            
        case 4 % 표준화 점수
            mean_score = mean(selected_data, 2);
            factor_scores_custom(:, f) = zscore(mean_score);
            method_name = 'Standardized Score';
            
        case 5 % 회귀점수 (기존 factoran 결과 사용)
            if exist('stats', 'var') && isfield(stats, 'scores')
                factor_scores_custom(:, f) = stats.scores(:, f);
                method_name = 'Regression Score';
            else
                % 회귀점수 계산 불가시 평균점수로 대체
                factor_scores_custom(:, f) = mean(selected_data, 2);
                method_name = 'Mean Score (fallback)';
            end
    end
    
    % 기술통계량 계산
    factor_mean = mean(factor_scores_custom(:, f));
    factor_std = std(factor_scores_custom(:, f));
    factor_min = min(factor_scores_custom(:, f));
    factor_max = max(factor_scores_custom(:, f));
    
    fprintf('  계산 방법: %s\n', method_name);
    fprintf('  평균: %.3f, 표준편차: %.3f\n', factor_mean, factor_std);
    fprintf('  범위: [%.3f, %.3f]\n', factor_min, factor_max);
    
    % 세부 정보 저장
    factor_score_details.(sprintf('factor_%d', f)) = struct();
    factor_score_details.(sprintf('factor_%d', f)).name = factor_name;
    factor_score_details.(sprintf('factor_%d', f)).items = selected_items;
    factor_score_details.(sprintf('factor_%d', f)).item_names = final_var_names(selected_items);
    factor_score_details.(sprintf('factor_%d', f)).loadings = factor_loadings_detail{f};
    factor_score_details.(sprintf('factor_%d', f)).method = method_name;
    factor_score_details.(sprintf('factor_%d', f)).statistics = struct('mean', factor_mean, 'std', factor_std, 'min', factor_min, 'max', factor_max);
end

%% 11.4 특정 요인 선택 및 심화 분석
fprintf('\n\n--- 특정 요인 심화 분석 ---\n');
fprintf('분석할 요인을 선택하세요:\n');
for f = 1:recommended_factors
    fprintf('%d. %s\n', f, factor_interpretation{f});
end

selected_factor = input('선택 (번호): ');
if isempty(selected_factor) || selected_factor < 1 || selected_factor > recommended_factors
    selected_factor = 1;
    fprintf('기본값으로 첫 번째 요인을 선택했습니다.\n');
end

% 선택된 요인의 심화 분석
selected_factor_name = factor_interpretation{selected_factor};
selected_factor_scores = factor_scores_custom(:, selected_factor);

fprintf('\n=== %s 심화 분석 ===\n', selected_factor_name);

% 분포 분석
fprintf('분포 분석:\n');
fprintf('  25%%ile: %.3f\n', prctile(selected_factor_scores, 25));
fprintf('  50%%ile (중위수): %.3f\n', prctile(selected_factor_scores, 50));
fprintf('  75%%ile: %.3f\n', prctile(selected_factor_scores, 75));
fprintf('  왜도: %.3f\n', skewness(selected_factor_scores));
fprintf('  첨도: %.3f\n', kurtosis(selected_factor_scores));

% 상위/하위 그룹 식별
top_10_pct_threshold = prctile(selected_factor_scores, 90);
bottom_10_pct_threshold = prctile(selected_factor_scores, 10);

top_performers = find(selected_factor_scores >= top_10_pct_threshold);
bottom_performers = find(selected_factor_scores <= bottom_10_pct_threshold);

fprintf('\n상하위 그룹 분석:\n');
fprintf('  상위 10%% 그룹: %d명 (점수 ≥ %.3f)\n', length(top_performers), top_10_pct_threshold);
fprintf('  하위 10%% 그룹: %d명 (점수 ≤ %.3f)\n', length(bottom_performers), bottom_10_pct_threshold);

%% 11.5 개별 참가자 분석 기능
fprintf('\n--- 개별 참가자 분석 ---\n');
fprintf('특정 참가자의 요인점수를 확인하시겠습니까? (y/n): ');
check_individual = input('', 's');

if strcmpi(check_individual, 'y')
    fprintf('\n참가자 ID를 입력하세요 (예: %s):\n', participant_ids{1});
    target_participant = input('', 's');
    
    participant_idx = find(strcmp(participant_ids, target_participant));
    
    if ~isempty(participant_idx)
        fprintf('\n=== %s의 요인점수 프로필 ===\n', target_participant);
        
        participant_scores = factor_scores_custom(participant_idx, :);
        
        for f = 1:recommended_factors
            factor_name = factor_interpretation{f};
            score = participant_scores(f);
            
            % 백분위수 계산
            percentile = mean(factor_scores_custom(:, f) <= score) * 100;
            
            % 등급 분류
            if percentile >= 90
                grade = 'A+ (상위 10%)';
            elseif percentile >= 75
                grade = 'A (상위 25%)';
            elseif percentile >= 50
                grade = 'B (평균 이상)';
            elseif percentile >= 25
                grade = 'C (평균 이하)';
            else
                grade = 'D (하위 25%)';
            end
            
            fprintf('  %s: %.3f (%.1f백분위, %s)\n', factor_name, score, percentile, grade);
        end
        
    else
        fprintf('해당 참가자를 찾을 수 없습니다.\n');
    end
end

%% 11.6 요인점수 상관관계 분석
fprintf('\n--- 요인간 상관관계 분석 ---\n');
factor_correlations = corr(factor_scores_custom);

fprintf('요인간 상관계수:\n');
fprintf('%15s', '');
for f = 1:recommended_factors
    fprintf('%12s', sprintf('F%d', f));
end
fprintf('\n');

for i = 1:recommended_factors
    factor_name_short = factor_interpretation{i};
    if length(factor_name_short) > 14
        factor_name_short = [factor_name_short(1:11), '...'];
    end
    fprintf('%-15s', factor_name_short);
    
    for j = 1:recommended_factors
        if i == j
            fprintf('%12s', '1.000');
        elseif i > j
            fprintf('%12.3f', factor_correlations(i, j));
            if abs(factor_correlations(i, j)) > 0.7
                fprintf('*'); % 높은 상관관계 표시
            else
                fprintf(' ');
            end
        else
            fprintf('%12s', '');
        end
    end
    fprintf('\n');
end

%% 11.7 결과 저장 (확장된 버전)
fprintf('\n--- 확장된 결과 저장 ---\n');

% 커스텀 요인점수 테이블 생성
factor_names_for_table = cell(1, recommended_factors);
for f = 1:recommended_factors
    factor_names_for_table{f} = sprintf('%s_Score', factor_interpretation{f});
    % 특수문자 제거 (Excel 호환성)
    factor_names_for_table{f} = regexprep(factor_names_for_table{f}, '[^a-zA-Z0-9_]', '_');
end

% 개별 문항점수와 요인점수를 포함한 종합 테이블
comprehensive_results = array2table([final_data, factor_scores_custom], ...
    'VariableNames', [final_var_names, factor_names_for_table], ...
    'RowNames', participant_ids);

% 요인별 세부 분석 결과 테이블
factor_summary_table = table();
for f = 1:recommended_factors
    factor_info = struct();
    factor_info.Factor_Name = {factor_interpretation{f}};
    factor_info.N_Items = length(factor_items{f});
    factor_info.Mean_Score = mean(factor_scores_custom(:, f));
    factor_info.Std_Score = std(factor_scores_custom(:, f));
    factor_info.Min_Score = min(factor_scores_custom(:, f));
    factor_info.Max_Score = max(factor_scores_custom(:, f));
    factor_info.Top_Loading_Item = {final_var_names{factor_items{f}(1)}};
    factor_info.Max_Loading = max(abs(factor_loadings_detail{f}));
    
    if f == 1
        factor_summary_table = struct2table(factor_info);
    else
        factor_summary_table = [factor_summary_table; struct2table(factor_info)];
    end
end

% Excel 파일로 저장
try
    % 종합 결과 시트
    writetable(comprehensive_results, 'detailed_factor_analysis_results.xlsx', ...
        'Sheet', 'Comprehensive_Data', 'WriteRowNames', true);
    
    % 요인점수만 별도 시트
    factor_scores_table = array2table(factor_scores_custom, ...
        'VariableNames', factor_names_for_table, ...
        'RowNames', participant_ids);
    writetable(factor_scores_table, 'detailed_factor_analysis_results.xlsx', ...
        'Sheet', 'Factor_Scores', 'WriteRowNames', true);
    
    % 요인 요약 정보 시트
    writetable(factor_summary_table, 'detailed_factor_analysis_results.xlsx', ...
        'Sheet', 'Factor_Summary', 'WriteRowNames', false);
    
    % 요인-문항 매핑 시트
    mapping_data = {};
    for f = 1:recommended_factors
        for i = 1:length(factor_items{f})
            item_idx = factor_items{f}(i);
            mapping_data{end+1, 1} = factor_interpretation{f};
            mapping_data{end, 2} = f;
            mapping_data{end, 3} = final_var_names{item_idx};
            mapping_data{end, 4} = factor_loadings_detail{f}(i);
            mapping_data{end, 5} = communality(item_idx);
        end
    end
    
    mapping_table = cell2table(mapping_data, ...
        'VariableNames', {'Factor_Name', 'Factor_Number', 'Item_Name', 'Loading', 'Communality'});
    writetable(mapping_table, 'detailed_factor_analysis_results.xlsx', ...
        'Sheet', 'Factor_Item_Mapping', 'WriteRowNames', false);
    
    fprintf('상세 결과가 detailed_factor_analysis_results.xlsx에 저장되었습니다.\n');
    fprintf('포함된 시트:\n');
    fprintf('  - Comprehensive_Data: 전체 데이터 (문항 + 요인점수)\n');
    fprintf('  - Factor_Scores: 요인점수만\n');
    fprintf('  - Factor_Summary: 요인별 기술통계\n');
    fprintf('  - Factor_Item_Mapping: 요인-문항 매핑\n');
    
catch ME
    fprintf('Excel 저장 실패: %s\n', ME.message);
    fprintf('MAT 파일로 저장합니다.\n');
end

% MAT 파일에 확장된 결과 저장
analysis_results.custom_factor_scores = factor_scores_custom;
analysis_results.factor_interpretations = factor_interpretation;
analysis_results.factor_score_details = factor_score_details;
analysis_results.factor_correlations = factor_correlations;
analysis_results.scoring_method = method_name;

save('detailed_factor_analysis_results.mat', 'analysis_results', 'factor_scores_custom', ...
     'factor_interpretation', 'factor_score_details', 'comprehensive_results');

fprintf('\n=== 커스텀 요인점수 분석 완료! ===\n');
fprintf('선택된 방법: %s\n', method_name);
fprintf('생성된 요인점수: %d개\n', recommended_factors);
fprintf('분석 대상: %d명\n', length(participant_ids));

%% 사용자 정의 함수들

function alpha = cronbach_alpha(data)
    if size(data, 2) < 2
        alpha = NaN;
        return;
    end
    
    k = size(data, 2);
    item_variances = var(data, 'omitnan');
    sum_item_var = sum(item_variances);
    total_scores = sum(data, 2, 'omitnan');
    total_var = var(total_scores, 'omitnan');
    alpha = (k / (k - 1)) * (1 - sum_item_var / total_var);
end

function kmo = calculate_kmo_robust(R)
    n = size(R, 1);
    try
        % 수치적 안정성을 위한 정규화
        R_reg = R + eye(n) * 1e-8; % Ridge regularization
        inv_R = inv(R_reg);
        
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
        off_diag = ~eye(n);
        
        sum_R = sum(R_sq(off_diag));
        sum_partial = sum(partial_sq(off_diag));
        
        kmo = sum_R / (sum_R + sum_partial);
        
    catch ME
        warning('KMO 계산 실패: %s', ME.message);
        kmo = 0.5; % 기본값
    end
end

function [chi2, p_value] = bartlett_test_robust(R, n)
    p = size(R, 1);
    try
        % 행렬식 계산 (수치적 안정성 확보)
        det_R = det(R);
        if det_R <= 0
            det_R = 1e-10; % 특이행렬 처리
        end
        
        % Bartlett 검정통계량
        chi2 = -(n - 1 - (2*p + 5)/6) * log(det_R);
        df = p * (p - 1) / 2;
        
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

function optimal_factors = parallel_analysis_improved(data, n_iterations)
    [n_obs, n_vars] = size(data);
    
    try
        % 실제 상관행렬의 고유값
        actual_corr = corr(data, 'Type', 'Pearson');
        actual_eigenvals = eig(actual_corr);
        actual_eigenvals = sort(actual_eigenvals, 'descend');
        
        % 무작위 데이터의 고유값 분포
        random_eigenvals = zeros(n_iterations, n_vars);
        
        for iter = 1:n_iterations
            % 정규분포 무작위 데이터 생성
            random_data = randn(n_obs, n_vars);
            
            % 실제 데이터의 분산 구조 반영
            for j = 1:n_vars
                random_data(:, j) = random_data(:, j) * std(data(:, j));
            end
            
            try
                random_corr = corr(random_data, 'Type', 'Pearson');
                eigenvals = eig(random_corr);
                random_eigenvals(iter, :) = sort(eigenvals, 'descend');
            catch
                if iter > 1
                    random_eigenvals(iter, :) = random_eigenvals(iter-1, :);
                else
                    random_eigenvals(iter, :) = ones(1, n_vars) / n_vars;
                end
            end
        end
        
        % 95백분위수와 비교
        random_95th = prctile(random_eigenvals, 95, 1);
        optimal_factors = sum(actual_eigenvals > random_95th');
        optimal_factors = max(1, min(optimal_factors, floor(n_vars/3)));
        
    catch ME
        warning('병렬분석 실패: %s', ME.message);
        optimal_factors = sum(actual_eigenvals > 1); % Kaiser 기준으로 대체
    end
end

function optimal_factors = velicer_map_robust(R)
    n = size(R, 1);
    max_factors = min(8, n-1);
    
    try
        % 고유분해
        [V, D] = eig(R);
        [lambda, idx] = sort(diag(D), 'descend');
        V = V(:, idx);
        lambda = max(lambda, 0); % 음수 고유값 처리
        
        % MAP 값 계산
        map_values = zeros(max_factors+1, 1);
        
        for m = 0:max_factors
            if m == 0
                reconstructed = zeros(n);
            else
                loadings = V(:, 1:m) * diag(sqrt(lambda(1:m)));
                reconstructed = loadings * loadings';
            end
            
            residual = R - reconstructed;
            residual(1:n+1:end) = 0; % 대각선 0으로
            
            map_values(m+1) = mean(residual(:).^2);
        end
        
        [~, min_idx] = min(map_values);
        optimal_factors = min_idx - 1;
        
    catch ME
        warning('MAP 계산 실패: %s', ME.message);
        eigenvals = eig(R);
        optimal_factors = sum(eigenvals > 1);
    end
end

function fit_indices = calculate_fit_indices_robust(data, loadings, specific_var, stats)
    fit_indices = struct();
    [n, p] = size(data);
    k = size(loadings, 2); % 요인 수
    
    try
        % 기본 통계량
        if isfield(stats, 'chi2') && ~isnan(stats.chi2)
            fit_indices.chi2 = stats.chi2;
            fit_indices.df = stats.dfe;
            fit_indices.p_value = stats.p;
            
            % RMSEA 계산 (개선된 버전)
            if stats.dfe > 0 && stats.chi2 > stats.dfe
                fit_indices.rmsea = sqrt((stats.chi2 - stats.dfe) / (stats.dfe * (n - 1)));
                
                % RMSEA 90% 신뢰구간 (근사적)
                ncp_lower = max(0, stats.chi2 - stats.dfe - 1.96*sqrt(2*stats.dfe));
                ncp_upper = stats.chi2 - stats.dfe + 1.96*sqrt(2*stats.dfe);
                
                fit_indices.rmsea_ci = [sqrt(ncp_lower/(stats.dfe*(n-1))), ...
                                       sqrt(ncp_upper/(stats.dfe*(n-1)))];
                
                % CFI 계산 (근사적)
                null_chi2 = (p * (p-1) / 2) * (n-1); % 독립모델 가정
                fit_indices.cfi = 1 - max(0, stats.chi2 - stats.dfe) / max(stats.chi2 - stats.dfe, null_chi2 - (p*(p-1)/2));
                
                % TLI (Tucker-Lewis Index)
                null_df = p * (p-1) / 2;
                fit_indices.tli = (null_chi2/null_df - stats.chi2/stats.dfe) / (null_chi2/null_df - 1);
                
            elseif stats.chi2 <= stats.dfe
                % 완벽 적합 또는 과적합
                fit_indices.rmsea = 0;
                fit_indices.rmsea_ci = [0, 0];
                fit_indices.cfi = 1;
                fit_indices.tli = 1;
            else
                fit_indices.rmsea = NaN;
                fit_indices.cfi = NaN;
                fit_indices.tli = NaN;
            end
        else
            fit_indices.chi2 = NaN;
            fit_indices.df = NaN;
            fit_indices.p_value = NaN;
            fit_indices.rmsea = NaN;
            fit_indices.cfi = NaN;
            fit_indices.tli = NaN;
        end
        
        % SRMR 계산 (Standardized Root Mean Square Residual)
        observed_corr = corr(data);
        implied_corr = loadings * loadings' + diag(specific_var);
        
        % 대각선 원소 제외한 잔차
        residual_corr = observed_corr - implied_corr;
        residual_corr(logical(eye(p))) = 0; % 대각선 0으로
        
        fit_indices.srmr = sqrt(mean(residual_corr(:).^2));
        
        % 추가 정보
        fit_indices.sample_size = n;
        fit_indices.n_variables = p;
        fit_indices.n_factors = k;
        
    catch ME
        warning('적합도 지수 계산 실패: %s', ME.message);
        fit_indices.chi2 = NaN;
        fit_indices.rmsea = NaN;
        fit_indices.cfi = NaN;
        fit_indices.tli = NaN;
        fit_indices.srmr = NaN;
    end
end