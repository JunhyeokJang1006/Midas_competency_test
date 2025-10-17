% ========================================================================
% MODEL_DIFF 버전: 새로운 그룹 정의로 분석
% ========================================================================
%
% 분석 그룹 재정의:
%   그룹 A (고역량): 자연성, 유익한 불연성
%   그룹 B (저역량): 유능한 불연성, 무능한 불연성
%   제외 그룹: 소화성, 성실한 가연성, 게으른 가연성, 위장형 소화성
%
% 주요 특징:
% 1. 개별 레이더 차트 생성 (통일된 스케일)
% 2. 그룹 A vs 그룹 B 역량 프로파일 비교
% 3. 하이퍼파라미터 튜닝 및 교차검증
% 4. 고도화된 상관 기반 가중치 시스템
% 5. Logistic regression과 상관분석 통합
% 6. 학습된 모델 저장 및 재사용 기능
% 7. 비활성/과활성 포함 분석, 신뢰불가 데이터 제외 추가
% 8. F1 스코어 통계적 유의성 검증
%    - 퍼뮤테이션 테스트 (5000회)
%    - 캐싱 시스템으로 재계산 방지
%    - Cohen's d 효과 크기 측정
% 9. 이상치 탐지 및 제거 기능
%    - IQR, Z-score, 백분위수 기반 이상치 탐지
%    - 역량별 이상치 개수 및 제거 결과 보고
%    - 설정 가능한 임계값 및 제거 옵션   
 
clear; clc; close all;

rng(42)  
%% ========================================================================
%                          PART 1: 초기 설정 및 데이터 로딩
% =========================================================================

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);
set(0, 'DefaultLineLineWidth', 2);

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent_model_diff';  % MODEL_DIFF 전용 디렉토리
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
config.model_file = fullfile(config.output_dir, 'trained_logitboost_model_diff.mat');
config.use_saved_model = false;  % 새 그룹 정의이므로 모델 재학습 필요
config.baseline_type = 'weighted';  % 'simple' 또는 'weighted' 선택

% 그룹 비교 방식: 커스텀 그룹 정의
% 그룹 A (고역량): 자연성, 유익한 불연성
% 그룹 B (저역량): 유능한 불연성, 무능한 불연성

% 퍼뮤테이션 테스트 설정
config.force_recalc_permutation = false;  % 퍼뮤테이션 강제 재계산 여부

% 파일 관리 시스템 설정
config.create_backup = true;  % 백업 폴더 생성 여부
config.backup_folder = 'backup';  % 백업 폴더 이름
config.use_timestamp = false;  % 파일명에 타임스탬프 사용 여부

% 이상치 제거 설정
config.outlier_removal = struct();
config.outlier_removal.enabled = true;  % 이상치 제거 활성화 여부
config.outlier_removal.method = 'zscore';  % 'iqr', 'zscore', 'percentile', 'none' 중 선택
config.outlier_removal.iqr_multiplier = 1.5;  % IQR 방법의 배수 (기본값: 1.5)
config.outlier_removal.zscore_threshold = 3;  % Z-score 방법의 임계값 (기본값: 3)
config.outlier_removal.percentile_bounds = [5, 95];  % 백분위수 방법의 범위 (기본값: 5%, 95%)
config.outlier_removal.apply_to_competencies = true;  % 역량 점수에 적용 여부
config.outlier_removal.report_outliers = true;  % 이상치 제거 결과 보고 여부

% 성과 순위 정의 (MODEL_DIFF 버전 - 분석 대상 그룹만)
config.performance_ranking = containers.Map(...
    {'자연성', '유익한 불연성', '유능한 불연성', '무능한 불연성'}, ...
    [4, 3, 2, 1]);

% 순서형 로지스틱 회귀용 레벨 정의 (1: 최하위, 4: 최상위)
config.ordinal_levels = containers.Map(...
    {'무능한 불연성', '유능한 불연성', '유익한 불연성', '자연성'}, ...
    [1, 2, 3, 4]);

config.level_names = {'무능한 불연성', '유능한 불연성', '유익한 불연성', '자연성'};

% 그룹 정의 (이진 분류용)
config.high_performers = {'자연성', '유익한 불연성'};  % 그룹 A (고역량)
config.low_performers = {'유능한 불연성', '무능한 불연성'};  % 그룹 B (저역량)
config.excluded_from_analysis = {'소화성', '성실한 가연성', '게으른 가연성'};  % 분석에서 제외
config.excluded_types = {'위장형 소화성'};  % 위장형 소화성도 제외

fprintf('\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║       MODEL_DIFF 버전: 새로운 그룹 정의 분석               ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n');
fprintf('▶ 그룹 A (고역량): %s\n', strjoin(config.high_performers, ', '));
fprintf('▶ 그룹 B (저역량): %s\n', strjoin(config.low_performers, ', '));
fprintf('▶ 제외 그룹: %s\n', strjoin([config.excluded_from_analysis, config.excluded_types], ', '));
fprintf('────────────────────────────────────────────\n\n');

%% 1.1 데이터 로딩
fprintf('【STEP 1】 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

% 파일 존재 여부 확인
if ~exist(config.hr_file, 'file')
    error('HR 데이터 파일을 찾을 수 없습니다: %s', config.hr_file);
end

if ~exist(config.comp_file, 'file')
    error('역량검사 데이터 파일을 찾을 수 없습니다: %s', config.comp_file);
end

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

%% 파일 관리 시스템 함수 정의
function backup_and_prepare_file(filepath, config)
    % 백업 처리 및 파일 준비 함수
    if ~config.create_backup
        return;
    end

    % 백업 폴더 생성
    backup_dir = fullfile(fileparts(filepath), config.backup_folder);
    if ~exist(backup_dir, 'dir')
        mkdir(backup_dir);
        fprintf('   백업 폴더 생성: %s\n', backup_dir);
    end

    % 기존 파일이 있으면 백업으로 이동
    if exist(filepath, 'file')
        [~, filename, ext] = fileparts(filepath);
        backup_timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
        backup_filename = sprintf('%s_%s%s', filename, backup_timestamp, ext);
        backup_filepath = fullfile(backup_dir, backup_filename);

        try
            movefile(filepath, backup_filepath);
            fprintf('   기존 파일을 백업 폴더로 이동: %s → %s/%s\n', ...
                [filename ext], config.backup_folder, backup_filename);
        catch
            fprintf('  ⚠ 백업 실패: %s\n', filepath);
        end
    end
end

function final_filepath = get_managed_filepath(base_dir, filename, config)
    % 파일 관리 시스템에 따른 최종 파일 경로 생성
    if config.use_timestamp && ~config.create_backup
        % 타임스탬프 사용 모드 (기존 방식)
        [~, name, ext] = fileparts(filename);
        timestamped_filename = sprintf('%s_%s%s', name, config.timestamp, ext);
        final_filepath = fullfile(base_dir, timestamped_filename);
    else
        % 고정 파일명 사용 (백업 시스템과 함께)
        final_filepath = fullfile(base_dir, filename);
    end
end

%% 1.1-1 신뢰가능성 필터링 추가
fprintf('\n【STEP 1-1】 신뢰가능성 필터링\n');
fprintf('────────────────────────────────────────────\n');

% 신뢰가능성 컬럼 찾기
reliability_col_idx = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col_idx)
    fprintf('▶ 신뢰가능성 컬럼 발견: %s\n', comp_upper.Properties.VariableNames{reliability_col_idx});
    
    % 신뢰불가 데이터 제외
    reliability_data = comp_upper{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(comp_upper), 1);
    end
    
    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));
    
    % 신뢰가능한 데이터만 유지
    comp_upper = comp_upper(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능한 데이터: %d명\n', height(comp_upper));
else
    fprintf('  ⚠ 신뢰가능성 컬럼이 없습니다. 모든 데이터를 사용합니다.\n');
end

%% 1.1-2 개발자 데이터 필터링 옵션
fprintf('\n【STEP 1-2】 개발자 데이터 필터링 옵션\n');
fprintf('────────────────────────────────────────────\n');

% 개발자여부 컬럼 찾기
developer_col_idx = find(contains(comp_upper.Properties.VariableNames, {'개발자여부', '개발자', 'developer'}), 1);
if ~isempty(developer_col_idx)
    fprintf('▶ 개발자여부 컬럼 발견: %s\n', comp_upper.Properties.VariableNames{developer_col_idx});

    % 개발자 데이터 분석
    developer_data = comp_upper{:, developer_col_idx};

    % 개발자 판별 로직 (다양한 형태 지원)
    if iscell(developer_data)
        developer_idx = strcmp(developer_data, '개발자') | strcmp(developer_data, 'Y') | ...
                       strcmp(developer_data, 'Yes') | strcmp(developer_data, '예');
    else
        developer_idx = (developer_data == 1) | (developer_data == true);
    end

    dev_count = sum(developer_idx);
    non_dev_count = sum(~developer_idx);
    total_count = height(comp_upper);

    fprintf('  현재 데이터 구성:\n');
    fprintf('    - 개발자: %d명 (%.1f%%)\n', dev_count, dev_count/total_count*100);
    fprintf('    - 비개발자: %d명 (%.1f%%)\n', non_dev_count, non_dev_count/total_count*100);

    % 사용자 선택 옵션 제공
    fprintf('\n  필터링 옵션을 선택하세요:\n');
    fprintf('    1. 모든 데이터 사용 (기본값)\n');
    fprintf('    2. 개발자 제외\n');
    fprintf('    3. 개발자만 분석\n');

    % 자동으로 기본값 선택 (모든 데이터 포함 분석)
    user_choice = '1';
    fprintf('  ✓ 자동 선택: 모든 데이터 포함 분석\n');

    % 선택에 따른 필터링 적용
    switch user_choice
        case '1'
            fprintf('  ✓ 모든 데이터 사용: %d명\n', height(comp_upper));
            config.developer_filter = 'all';
            config.applied_filter = '모든 데이터 사용';

        case '2'
            comp_upper = comp_upper(~developer_idx, :);
            fprintf('  ✓ 개발자 제외 적용: %d명 → %d명\n', total_count, height(comp_upper));
            config.developer_filter = 'exclude_dev';
            config.applied_filter = '개발자 제외';

        case '3'
            comp_upper = comp_upper(developer_idx, :);
            fprintf('  ✓ 개발자만 분석: %d명 → %d명\n', total_count, height(comp_upper));
            config.developer_filter = 'dev_only';
            config.applied_filter = '개발자만 분석';

        otherwise
            fprintf('  ⚠ 잘못된 선택. 기본값(모든 데이터) 사용\n');
            config.developer_filter = 'all';
            config.applied_filter = '모든 데이터 사용';
    end

else
    fprintf('  ⚠ 개발자여부 컬럼이 없습니다. 모든 데이터를 사용합니다.\n');
    config.developer_filter = 'all';
    config.applied_filter = '모든 데이터 사용 (개발자여부 컬럼 없음)';
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

% 위장형 소화성만 제외
excluded_mask = strcmp(hr_clean{:, talent_col_idx}, '위장형 소화성');
hr_clean = hr_clean(~excluded_mask, :);

% 인재유형 분포 분석
talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts = accumarray(type_indices, 1);

fprintf('\n초기 인재유형 분포 (HR 데이터, 역량검사 매칭 전):\n');
for i = 1:length(unique_types)
    fprintf('  • %-20s: %3d명 (%5.1f%%)\n', ...
        unique_types{i}, type_counts(i), type_counts(i)/sum(type_counts)*100);
end
fprintf('  ⚠️ 주의: 이 숫자는 중간 단계입니다.\n');
fprintf('     최종 분석 인원은 결측치 필터링 후 확정됩니다 (STEP 9 참조).\n');

%% 1.3 역량 데이터 처리 - 비활성/과활성 포함 분석
fprintf('\n【STEP 3】 역량 데이터 처리 (비활성/과활성 포함)\n');
fprintf('────────────────────────────────────────────\n');

% ID 컬럼 찾기
comp_id_col = find(contains(lower(comp_upper.Properties.VariableNames), {'id', '사번'}), 1);
if isempty(comp_id_col)
    error('역량 데이터에서 ID 컬럼을 찾을 수 없습니다.');
end

% 비활성/과활성 포함하여 분석 (모든 역량 데이터 사용)
fprintf('▶ 비활성/과활성을 포함하여 모든 역량 데이터 분석\n');

% 유효한 역량 컬럼 추출 (비활성/과활성 포함)
valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(comp_upper)
    col_name = comp_upper.Properties.VariableNames{i};

    % 모든 역량 컬럼 포함 (비활성/과활성도 포함)
    
    % 숫자 데이터인지 확인
    col_data = comp_upper{:, i};
    if isnumeric(col_data) && ~all(isnan(col_data))
        valid_data = col_data(~isnan(col_data));
        if length(valid_data) >= 5
            % 분산이 0인 경우도 처리
            data_var = var(valid_data);
            if (data_var > 0 || length(unique(valid_data)) > 1) && ...
               all(valid_data >= 0) && all(valid_data <= 100)
                valid_comp_cols{end+1} = col_name;
                valid_comp_indices(end+1) = i;
            end
        end
    end
end

fprintf('\n포함된 모든 역량 컬럼 (%d개) - 비활성/과활성 포함:\n', length(valid_comp_cols));
for i = 1:min(length(valid_comp_cols), 10)  % 처음 10개만 출력
    fprintf('  - %s\n', valid_comp_cols{i});
end
if length(valid_comp_cols) > 10
    fprintf('  ... 외 %d개 더\n', length(valid_comp_cols) - 10);
end

if isempty(valid_comp_cols)
    error('유효한 역량 컬럼을 찾을 수 없습니다. 데이터를 확인해주세요.');
end

fprintf('\n▶ 사용할 역량 항목: %d개 (비활성/과활성 포함)\n', length(valid_comp_cols));
fprintf('  유효 역량 목록:\n');
for i = 1:min(10, length(valid_comp_cols))
    fprintf('    %d. %s\n', i, valid_comp_cols{i});
end
if length(valid_comp_cols) > 10
    fprintf('    ... 외 %d개\n', length(valid_comp_cols) - 10);
end

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

% 종합점수 매칭
comp_total_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_total{:, 1}, 'UniformOutput', false);
[~, ~, total_idx] = intersect(matched_ids, comp_total_ids_str);

if ~isempty(total_idx)
    total_scores_raw = comp_total{total_idx, end};
    % cell array인 경우 숫자 배열로 변환
    if iscell(total_scores_raw)
        total_scores = cell2mat(total_scores_raw);
    else
        total_scores = total_scores_raw;
    end
    % NaN이 아닌 숫자 데이터만 유지
    if isnumeric(total_scores)
        total_scores = total_scores(~isnan(total_scores));
    else
        total_scores = [];
    end
    fprintf('▶ 종합점수 통합: %d명\n', length(total_scores));
else
    total_scores = [];
    fprintf('⚠ 종합점수 데이터 없음\n');
end
%% ========================================================================
%            PART 2: 개선된 레이더 차트 (개별 Figure, 통일 스케일)
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║     PART 2: 개선된 레이더 차트 (통일 스케일)             ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% STEP 4.5: 역량별 Range 및 동질성 분석
fprintf('\n【STEP 4.5】 역량별 Range 및 동질성 분석\n');
fprintf('────────────────────────────────────────────\n\n');

% 역량별 기술통계 계산
range_stats = table();
range_stats.Competency = valid_comp_cols';
range_stats.Min = zeros(length(valid_comp_cols), 1);
range_stats.Q1 = zeros(length(valid_comp_cols), 1);
range_stats.Median = zeros(length(valid_comp_cols), 1);
range_stats.Q3 = zeros(length(valid_comp_cols), 1);
range_stats.Max = zeros(length(valid_comp_cols), 1);
range_stats.Range = zeros(length(valid_comp_cols), 1);
range_stats.IQR = zeros(length(valid_comp_cols), 1);
range_stats.CV = zeros(length(valid_comp_cols), 1);
range_stats.Homogeneity = categorical(zeros(length(valid_comp_cols), 1));

% 통계 계산 루프
for i = 1:length(valid_comp_cols)
    comp_data = matched_comp{:, i};
    valid_data = comp_data(~isnan(comp_data));

    range_stats.Min(i) = min(valid_data);
    range_stats.Q1(i) = prctile(valid_data, 25);
    range_stats.Median(i) = median(valid_data);
    range_stats.Q3(i) = prctile(valid_data, 75);
    range_stats.Max(i) = max(valid_data);
    range_stats.Range(i) = range_stats.Max(i) - range_stats.Min(i);
    range_stats.IQR(i) = range_stats.Q3(i) - range_stats.Q1(i);

    if mean(valid_data) > 0
        range_stats.CV(i) = std(valid_data) / mean(valid_data);
    else
        range_stats.CV(i) = NaN;
    end

    if range_stats.CV(i) < 0.15
        range_stats.Homogeneity(i) = 'High';
    elseif range_stats.CV(i) < 0.3
        range_stats.Homogeneity(i) = 'Medium';
    else
        range_stats.Homogeneity(i) = 'Low';
    end
end

% 결과 출력
high_homogeneity = range_stats.Homogeneity == 'High';
fprintf('동질성 높은 역량 (CV < 0.15): %d개\n', sum(high_homogeneity));
narrow_range = range_stats.Range < 30;
fprintf('Range가 좁은 역량 (30점 미만): %d개\n', sum(narrow_range));

% 동질성 높은 역량 저장 (나중에 해석에서 사용)
high_homo_comps = range_stats(high_homogeneity, :);

fprintf('\n✅ 역량별 Range 및 동질성 분석 완료\n');

%% STEP 4.5.2: 역량 검사 분포 시각화
fprintf('\n【STEP 4.5.2】 역량 검사 분포 시각화\n');
fprintf('────────────────────────────────────────────\n');

% 분포 분석을 위한 데이터 준비 (개선된 Cut-off 분석)
n_comps = length(valid_comp_cols);
distribution_stats = table();
distribution_stats.Competency = valid_comp_cols';
distribution_stats.Mean = zeros(n_comps, 1);
distribution_stats.Std = zeros(n_comps, 1);
distribution_stats.Skewness = zeros(n_comps, 1);
distribution_stats.Kurtosis = zeros(n_comps, 1);

% 새로운 Cut-off 분석 필드들
distribution_stats.Has_Cutoff = false(n_comps, 1);
distribution_stats.Cutoff_Types = cell(n_comps, 1);
distribution_stats.Cutoff_Severity = zeros(n_comps, 1);
distribution_stats.Range_Utilization = zeros(n_comps, 1);
distribution_stats.Normality_Test = cell(n_comps, 1);
distribution_stats.Normality_P_Value = zeros(n_comps, 1);
distribution_stats.Floor_Effect = false(n_comps, 1);
distribution_stats.Ceiling_Effect = false(n_comps, 1);
distribution_stats.Lower_Truncation = false(n_comps, 1);
distribution_stats.Upper_Truncation = false(n_comps, 1);
distribution_stats.Distribution_Type = categorical(repmat({'Normal'}, n_comps, 1));

fprintf('▶ 역량별 분포 특성 분석 중...\n');

% 각 역량별 분포 분석
for i = 1:n_comps
    comp_data = matched_comp{:, i};
    valid_data = comp_data(~isnan(comp_data));

    if length(valid_data) < 10
        continue;
    end

    % 기본 통계
    distribution_stats.Mean(i) = mean(valid_data);
    distribution_stats.Std(i) = std(valid_data);
    distribution_stats.Skewness(i) = skewness(valid_data);
    distribution_stats.Kurtosis(i) = kurtosis(valid_data);

    % 개선된 Cut-off 효과 탐지
    theoretical_range = [0, 100]; % 역량 점수 이론적 범위

    try
        [has_cutoff, cutoff_info] = detect_cutoff_effect(valid_data, theoretical_range);

        % 결과를 테이블에 저장
        distribution_stats.Has_Cutoff(i) = has_cutoff;
        distribution_stats.Cutoff_Types{i} = cutoff_info.cutoff_type_string;
        distribution_stats.Cutoff_Severity(i) = cutoff_info.severity_score;
        distribution_stats.Range_Utilization(i) = cutoff_info.range_utilization;
        distribution_stats.Normality_Test{i} = cutoff_info.normality_test;
        distribution_stats.Normality_P_Value(i) = cutoff_info.normality_p;
        distribution_stats.Floor_Effect(i) = cutoff_info.floor_effect;
        distribution_stats.Ceiling_Effect(i) = cutoff_info.ceiling_effect;
        distribution_stats.Lower_Truncation(i) = cutoff_info.lower_truncation;
        distribution_stats.Upper_Truncation(i) = cutoff_info.upper_truncation;

    catch ME
        fprintf('  ⚠ %s 역량 Cut-off 분석 오류: %s\n', valid_comp_cols{i}, ME.message);
        distribution_stats.Has_Cutoff(i) = false;
        distribution_stats.Cutoff_Types{i} = 'Analysis_Failed';
        distribution_stats.Cutoff_Severity(i) = 0;
        distribution_stats.Range_Utilization(i) = NaN;
        distribution_stats.Normality_Test{i} = 'Failed';
        distribution_stats.Normality_P_Value(i) = NaN;
    end

    % 개선된 분포 타입 분류
    if distribution_stats.Has_Cutoff(i)
        % Cut-off 효과가 있는 경우 더 세부적인 분류
        if contains(distribution_stats.Cutoff_Types{i}, 'Floor_Effect')
            distribution_stats.Distribution_Type(i) = 'Floor_Effect';
        elseif contains(distribution_stats.Cutoff_Types{i}, 'Ceiling_Effect')
            distribution_stats.Distribution_Type(i) = 'Ceiling_Effect';
        elseif contains(distribution_stats.Cutoff_Types{i}, 'Lower_Truncation')
            distribution_stats.Distribution_Type(i) = 'Lower_Truncated';
        elseif contains(distribution_stats.Cutoff_Types{i}, 'Upper_Truncation')
            distribution_stats.Distribution_Type(i) = 'Upper_Truncated';
        else
            distribution_stats.Distribution_Type(i) = 'Cut_off_Effect';
        end
    else
        % 기존 분류 방식 유지
        if abs(distribution_stats.Skewness(i)) > 1
            if distribution_stats.Skewness(i) > 0
                distribution_stats.Distribution_Type(i) = 'Right_Skewed';
            else
                distribution_stats.Distribution_Type(i) = 'Left_Skewed';
            end
        elseif distribution_stats.Kurtosis(i) > 4
            distribution_stats.Distribution_Type(i) = 'Heavy_Tailed';
        elseif distribution_stats.Kurtosis(i) < 2
            distribution_stats.Distribution_Type(i) = 'Light_Tailed';
        end
    end
end

% Cut-off 효과 분석 결과 요약
fprintf('\n▶ Cut-off 효과 분석 결과 요약...\n');
cutoff_summary = struct();
cutoff_summary.total_competencies = n_comps;
cutoff_summary.with_cutoff = sum(distribution_stats.Has_Cutoff);
cutoff_summary.floor_effects = sum(distribution_stats.Floor_Effect);
cutoff_summary.ceiling_effects = sum(distribution_stats.Ceiling_Effect);
cutoff_summary.lower_truncations = sum(distribution_stats.Lower_Truncation);
cutoff_summary.upper_truncations = sum(distribution_stats.Upper_Truncation);

fprintf('  전체 역량 수: %d개\n', cutoff_summary.total_competencies);
fprintf('  Cut-off 효과 발견: %d개 (%.1f%%)\n', cutoff_summary.with_cutoff, ...
        cutoff_summary.with_cutoff / cutoff_summary.total_competencies * 100);
fprintf('  - Floor 효과: %d개\n', cutoff_summary.floor_effects);
fprintf('  - Ceiling 효과: %d개\n', cutoff_summary.ceiling_effects);
fprintf('  - 하위 절단: %d개\n', cutoff_summary.lower_truncations);
fprintf('  - 상위 절단: %d개\n', cutoff_summary.upper_truncations);

% 가장 심각한 Cut-off 효과를 가진 역량 식별
if cutoff_summary.with_cutoff > 0
    [max_severity, max_idx] = max(distribution_stats.Cutoff_Severity);
    fprintf('  가장 심각한 Cut-off: %s (심각도: %.3f)\n', ...
            distribution_stats.Competency{max_idx}, max_severity);
end

% Cut-off 효과별 분포 타입 통계
dist_types = categories(distribution_stats.Distribution_Type);
fprintf('\n분포 타입별 통계:\n');
for dt = 1:length(dist_types)
    count = sum(distribution_stats.Distribution_Type == dist_types{dt});
    if count > 0
        fprintf('  - %s: %d개\n', dist_types{dt}, count);
    end
end

% 분포 시각화 생성
fprintf('\n▶ 역량 분포 subplot 그리드 생성 중...\n');

% subplot 그리드 크기 계산 (최대 20개까지만 표시)
n_display = min(n_comps, 20);
cols = ceil(sqrt(n_display));
rows = ceil(n_display / cols);

% Figure 생성
fig_dist = figure('Position', [50, 50, 300*cols, 250*rows], ...
                  'Name', '역량 분포 분석', 'Color', 'white');

for i = 1:n_display
    subplot(rows, cols, i);

    comp_data = matched_comp{:, i};
    valid_data = comp_data(~isnan(comp_data));

    if length(valid_data) < 10
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'middle', 'FontSize', 12, 'Interpreter', 'none');
        title(valid_comp_cols{i}, 'FontSize', 10, 'Interpreter', 'none');
        continue;
    end

    % 개선된 히스토그램과 분포 비교 시각화

    % Cut-off 효과에 따른 색상 설정
    if distribution_stats.Has_Cutoff(i)
        if distribution_stats.Floor_Effect(i)
            hist_color = [0.8 0.3 0.3]; % 빨간색 (Floor 효과)
        elseif distribution_stats.Ceiling_Effect(i)
            hist_color = [0.3 0.3 0.8]; % 파란색 (Ceiling 효과)
        elseif distribution_stats.Lower_Truncation(i)
            hist_color = [0.8 0.6 0.2]; % 주황색 (하위 절단)
        elseif distribution_stats.Upper_Truncation(i)
            hist_color = [0.6 0.2 0.8]; % 보라색 (상위 절단)
        else
            hist_color = [0.8 0.8 0.3]; % 노란색 (기타 Cut-off)
        end
    else
        hist_color = [0.3 0.6 0.9]; % 기본 파란색 (정상)
    end

    % 히스토그램
    histogram(valid_data, 'Normalization', 'pdf', 'FaceAlpha', 0.7, ...
              'EdgeColor', 'none', 'FaceColor', hist_color);
    hold on;

    % 이론적 정규분포 곡선 (전체 범위)
    theoretical_x = linspace(0, 100, 200);
    theoretical_norm = normpdf(theoretical_x, mean(valid_data), std(valid_data));
    plot(theoretical_x, theoretical_norm, '--', 'Color', [0.5 0.5 0.5], 'LineWidth', 1.5, 'DisplayName', '이론적 정규분포');

    % 실제 데이터 기반 정규분포 곡선
    data_x = linspace(min(valid_data), max(valid_data), 100);
    data_norm = normpdf(data_x, mean(valid_data), std(valid_data));
    plot(data_x, data_norm, 'r-', 'LineWidth', 2, 'DisplayName', '실제 데이터 기반');

    % Cut-off 영향 구간 강조
    if distribution_stats.Has_Cutoff(i)
        ylims = ylim;

        % Floor 효과 강조
        if distribution_stats.Floor_Effect(i)
            floor_thresh = distribution_stats.Normality_P_Value(i); % 임시로 다른 값 사용
            if i <= size(distribution_stats, 1) % 인덱스 범위 확인
                try
                    [~, cutoff_info] = detect_cutoff_effect(valid_data, [0, 100]);
                    floor_thresh = cutoff_info.floor_threshold;
                    fill([min(valid_data), floor_thresh, floor_thresh, min(valid_data)], ...
                         [0, 0, ylims(2), ylims(2)], 'r', 'FaceAlpha', 0.2, 'EdgeColor', 'none');
                catch, end
            end
        end

        % Ceiling 효과 강조
        if distribution_stats.Ceiling_Effect(i)
            try
                [~, cutoff_info] = detect_cutoff_effect(valid_data, [0, 100]);
                ceiling_thresh = cutoff_info.ceiling_threshold;
                fill([ceiling_thresh, max(valid_data), max(valid_data), ceiling_thresh], ...
                     [0, 0, ylims(2), ylims(2)], 'b', 'FaceAlpha', 0.2, 'EdgeColor', 'none');
            catch, end
        end
    end

    % 개선된 통계 정보 텍스트
    stats_text = sprintf('M=%.1f, SD=%.1f\n왜도=%.2f, 첨도=%.2f\n정규성 p=%.3f', ...
                        distribution_stats.Mean(i), distribution_stats.Std(i), ...
                        distribution_stats.Skewness(i), distribution_stats.Kurtosis(i), ...
                        distribution_stats.Normality_P_Value(i));

    % Cut-off 효과 정보 추가
    if distribution_stats.Has_Cutoff(i)
        cutoff_text = sprintf('\n Cut-off (%.2f)', distribution_stats.Cutoff_Severity(i));
        cutoff_types = strsplit(distribution_stats.Cutoff_Types{i}, ', ');
        if length(cutoff_types) <= 2
            cutoff_text = [cutoff_text sprintf('\n%s', distribution_stats.Cutoff_Types{i})];
        else
            cutoff_text = [cutoff_text sprintf('\n%s+', cutoff_types{1})];
        end
        stats_text = [stats_text cutoff_text];
    end

    text(0.02, 0.98, stats_text, 'Units', 'normalized', 'VerticalAlignment', 'top', ...
         'FontSize', 7, 'BackgroundColor', 'white', 'EdgeColor', 'black', 'Interpreter', 'none');

    title(valid_comp_cols{i}, 'FontSize', 10, 'Interpreter', 'none');
    xlabel('점수', 'FontSize', 8);
    ylabel('밀도', 'FontSize', 8);
    grid on;
    hold off;
end

sgtitle('역량별 분포 분석 (히스토그램 + 정규분포 곡선)', 'FontSize', 14, 'FontWeight', 'bold');

% Figure 저장
dist_file = get_managed_filepath(config.output_dir, 'competency_distributions.png', config);
saveas(fig_dist, dist_file);
fprintf('  ✓ 분포 그래프 저장: %s\n', dist_file);

% Q-Q Plot 생성 (정규성 검정 시각화)
fprintf('▶ Q-Q Plot 생성 중...\n');

fig_qq = figure('Position', [150, 150, 300*cols, 250*rows], ...
                'Name', '역량별 Q-Q Plot (정규성 검정)', 'Color', 'white');

for i = 1:n_display
    subplot(rows, cols, i);

    comp_data = matched_comp{:, i};
    valid_data = comp_data(~isnan(comp_data));

    if length(valid_data) < 10
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'middle', 'FontSize', 12, 'Interpreter', 'none');
        title(valid_comp_cols{i}, 'FontSize', 10, 'Interpreter', 'none');
        continue;
    end

    % Q-Q plot 생성
    try
        % 표준화된 데이터
        standardized_data = zscore(valid_data);

        % 이론적 분위수와 실제 분위수 계산
        n = length(standardized_data);
        theoretical_quantiles = norminv((1:n)' / (n+1));
        sample_quantiles = sort(standardized_data);

        % Q-Q plot
        plot(theoretical_quantiles, sample_quantiles, 'o', 'MarkerSize', 4, ...
             'MarkerFaceColor', hist_color, 'MarkerEdgeColor', hist_color);
        hold on;

        % 이론적 직선 (y = x)
        plot([-3, 3], [-3, 3], 'r--', 'LineWidth', 2);

        % Cut-off 효과가 있는 경우 색상 강조
        if distribution_stats.Has_Cutoff(i)
            % 벗어나는 점들 강조
            deviation = abs(sample_quantiles - theoretical_quantiles);
            outlier_threshold = 1.5 * std(deviation);
            outliers = deviation > outlier_threshold;

            if any(outliers)
                plot(theoretical_quantiles(outliers), sample_quantiles(outliers), ...
                     's', 'MarkerSize', 6, 'MarkerFaceColor', 'red', 'MarkerEdgeColor', 'darkred');
            end
        end

        % 정규성 검정 결과 표시
        normality_text = sprintf('정규성 검정\np = %.4f', distribution_stats.Normality_P_Value(i));
        if distribution_stats.Normality_P_Value(i) < 0.05
            normality_text = [normality_text '\n❌ 비정규'];
        else
            normality_text = [normality_text '\n✓ 정규'];
        end

        % Cut-off 정보 추가
        if distribution_stats.Has_Cutoff(i)
            normality_text = [normality_text sprintf('\n Cut-off\n심각도: %.2f', ...
                                                   distribution_stats.Cutoff_Severity(i))];
        end

        text(0.05, 0.95, normality_text, 'Units', 'normalized', 'VerticalAlignment', 'top', ...
             'FontSize', 7, 'BackgroundColor', 'white', 'EdgeColor', 'black', 'Interpreter', 'none');

        xlabel('이론적 분위수', 'FontSize', 8);
        ylabel('표본 분위수', 'FontSize', 8);
        title(valid_comp_cols{i}, 'FontSize', 10, 'Interpreter', 'none');
        grid on;
        axis equal;
        xlim([-3, 3]);
        ylim([-3, 3]);
        hold off;

    catch ME
        text(0.5, 0.5, sprintf('Q-Q Plot 오류:\n%s', ME.message), ...
             'HorizontalAlignment', 'center', 'VerticalAlignment', 'middle', ...
             'FontSize', 8, 'Interpreter', 'none');
        title(valid_comp_cols{i}, 'FontSize', 10, 'Interpreter', 'none');
    end
end

sgtitle('역량별 Q-Q Plot (정규분포 적합도 검정)', 'FontSize', 14, 'FontWeight', 'bold');

% Q-Q Plot 저장
qq_file = get_managed_filepath(config.output_dir, 'competency_qq_plots.png', config);
saveas(fig_qq, qq_file);
fprintf('  ✓ Q-Q Plot 저장: %s\n', qq_file);

% 박스플롯 생성 (이상치 탐지)
fprintf('▶ 역량별 박스플롯 생성 중...\n');

fig_box = figure('Position', [100, 100, 1400, 800], ...
                 'Name', '역량별 박스플롯 (이상치 탐지)', 'Color', 'white');

% 데이터 행렬 준비 (동일한 길이로 맞추기 위해 NaN 패딩)
max_samples = size(matched_comp, 1);
box_data = nan(max_samples, n_display);

for i = 1:n_display
    comp_data = matched_comp{:, i};
    valid_data = comp_data(~isnan(comp_data));
    box_data(1:length(valid_data), i) = valid_data;
end

boxplot(box_data, 'Labels', valid_comp_cols(1:n_display), 'Symbol', 'ro');
title('역량별 분포 및 이상치 탐지', 'FontSize', 14, 'FontWeight', 'bold');
xlabel('역량', 'FontSize', 12);
ylabel('점수', 'FontSize', 12);
xtickangle(45);
grid on;

% 박스플롯 저장
box_file = get_managed_filepath(config.output_dir, 'competency_boxplots.png', config);
saveas(fig_box, box_file);
fprintf('  ✓ 박스플롯 저장: %s\n', box_file);

% 분포 특성 요약 테이블 출력
fprintf('\n【분포 특성 요약】\n');
fprintf('════════════════════════════════════════════════════════════\n');

% Cut-off 효과 요약
cutoff_min_count = sum(distribution_stats.Floor_Effect);
cutoff_max_count = sum(distribution_stats.Ceiling_Effect);
fprintf('선발 효과(Truncated Distribution) 탐지:\n');
fprintf('  - 최소값 Cut-off 효과: %d개 역량\n', cutoff_min_count);
fprintf('  - 최대값 Cut-off 효과: %d개 역량\n', cutoff_max_count);

% 분포 타입별 개수
dist_types = categories(distribution_stats.Distribution_Type);
fprintf('\n분포 형태별 분류:\n');
for i = 1:length(dist_types)
    count = sum(distribution_stats.Distribution_Type == dist_types{i});
    fprintf('  - %s: %d개\n', dist_types{i}, count);
end

% 극단 분포 경고
extreme_skew = abs(distribution_stats.Skewness) > 2;
extreme_kurt = distribution_stats.Kurtosis > 6 | distribution_stats.Kurtosis < 1;

if any(extreme_skew)
    fprintf('\n 극단 왜도 역량 (|skewness| > 2):\n');
    extreme_comps = distribution_stats.Competency(extreme_skew);
    for i = 1:length(extreme_comps)
        fprintf('  - %s (왜도: %.2f)\n', extreme_comps{i}, ...
                distribution_stats.Skewness(find(strcmp(distribution_stats.Competency, extreme_comps{i}))));
    end
end

if any(extreme_kurt)
    fprintf('\n 극단 첨도 역량 (kurtosis > 6 or < 1):\n');
    extreme_comps = distribution_stats.Competency(extreme_kurt);
    for i = 1:length(extreme_comps)
        fprintf('  - %s (첨도: %.2f)\n', extreme_comps{i}, ...
                distribution_stats.Kurtosis(find(strcmp(distribution_stats.Competency, extreme_comps{i}))));
    end
end

fprintf('\n✅ 역량 검사 분포 시각화 완료\n');

%% 이상치 제거
if config.outlier_removal.enabled
    fprintf('\n【STEP 4.8】 이상치 탐지 및 제거\n');
    fprintf('────────────────────────────────────────────\n');

    fprintf('▶ 이상치 제거 방법: %s\n', config.outlier_removal.method);

    % 역량 데이터에 이상치 제거 적용
    if config.outlier_removal.apply_to_competencies
        original_size = height(matched_comp);
        comp_data = matched_comp{:, :};
        outlier_mask = false(size(comp_data));
        outlier_report = struct();

        % 이상치 탐지 방법에 따른 처리
        switch config.outlier_removal.method
            case 'iqr'
                Q1 = prctile(comp_data, 25, 1);
                Q3 = prctile(comp_data, 75, 1);
                IQR = Q3 - Q1;
                lower_bound = Q1 - config.outlier_removal.iqr_multiplier * IQR;
                upper_bound = Q3 + config.outlier_removal.iqr_multiplier * IQR;

                for col = 1:size(comp_data, 2)
                    col_data = comp_data(:, col);
                    outlier_mask(:, col) = (col_data < lower_bound(col)) | ...
                                          (col_data > upper_bound(col));
                end

                outlier_report.method = 'IQR';
                outlier_report.multiplier = config.outlier_removal.iqr_multiplier;
                outlier_report.bounds = [lower_bound; upper_bound];

            case 'zscore'
                % NaN 값 처리하여 zscore 계산
                data_zscore = abs(zscore(comp_data, 0, 1));
                outlier_mask = data_zscore > config.outlier_removal.zscore_threshold;

                outlier_report.method = 'Z-score';
                outlier_report.threshold = config.outlier_removal.zscore_threshold;

            case 'percentile'
                lower_pct = config.outlier_removal.percentile_bounds(1);
                upper_pct = config.outlier_removal.percentile_bounds(2);

                lower_bound = prctile(comp_data, lower_pct, 1);
                upper_bound = prctile(comp_data, upper_pct, 1);

                for col = 1:size(comp_data, 2)
                    col_data = comp_data(:, col);
                    outlier_mask(:, col) = (col_data < lower_bound(col)) | ...
                                          (col_data > upper_bound(col));
                end

                outlier_report.method = 'Percentile';
                outlier_report.bounds_pct = [lower_pct, upper_pct];
                outlier_report.bounds = [lower_bound; upper_bound];

            case 'none'
                % 이상치 제거 안함
                fprintf('▶ 이상치 제거 비활성화됨 (method = none)\n');
            otherwise
                fprintf('▶ 알 수 없는 이상치 제거 방법: %s\n', config.outlier_removal.method);
        end

        if ~strcmp(config.outlier_removal.method, 'none')
            % 행별 이상치 판별 (한 역량에서라도 이상치인 경우 해당 행 제거)
            row_outliers = any(outlier_mask, 2);
            clean_indices = ~row_outliers;

            outlier_report.total_outliers = sum(row_outliers);
            outlier_report.outlier_percentage = sum(row_outliers) / size(comp_data, 1) * 100;
            outlier_report.outlier_by_competency = sum(outlier_mask, 1);

            % 모든 관련 데이터 동일하게 필터링
            matched_comp = matched_comp(clean_indices, :);
            matched_talent_types = matched_talent_types(clean_indices);
            matched_hr = matched_hr(clean_indices, :);
            if ~isempty(total_scores)
                total_scores = total_scores(clean_indices);
            end

            % 결과 보고
            if config.outlier_removal.report_outliers
                fprintf('\n▶ 이상치 제거 결과:\n');
                fprintf('  원본 데이터: %d명\n', original_size);
                fprintf('  이상치 제거 후: %d명\n', height(matched_comp));
                fprintf('  제거된 이상치: %d명 (%.1f%%)\n', ...
                    outlier_report.total_outliers, outlier_report.outlier_percentage);

                fprintf('\n▶ 역량별 이상치 발견 횟수:\n');
                for i = 1:length(valid_comp_cols)
                    if outlier_report.outlier_by_competency(i) > 0
                        fprintf('  • %-30s: %d개\n', ...
                            valid_comp_cols{i}, outlier_report.outlier_by_competency(i));
                    end
                end

                if strcmp(outlier_report.method, 'IQR')
                    fprintf('\n▶ IQR 기반 이상치 제거 기준 (배수: %.1f):\n', ...
                        outlier_report.multiplier);
                elseif strcmp(outlier_report.method, 'Z-score')
                    fprintf('\n▶ Z-score 기준: |z| > %.1f\n', outlier_report.threshold);
                elseif strcmp(outlier_report.method, 'Percentile')
                    fprintf('\n▶ 백분위 기준: %.1f%% - %.1f%%\n', ...
                        outlier_report.bounds_pct(1), outlier_report.bounds_pct(2));
                end
            end

            fprintf('\n✅ 이상치 제거 완료\n');
        end
    else
        fprintf('▶ 이상치 제거 비활성화됨 (apply_to_competencies = false)\n');
    end
else
    fprintf('\n▶ 이상치 제거 단계 생략됨\n');
end

%% 2.1 유형별 프로파일 계산 및 스케일 범위 설정
fprintf('\n【STEP 5】 유형별 프로파일 계산 및 스케일 설정\n');
fprintf('────────────────────────────────────────────\n');

unique_matched_types = unique(matched_talent_types);
n_types = length(unique_matched_types);

% 프로파일 계산
type_profiles = zeros(n_types, length(valid_comp_cols));
profile_stats = table();
profile_stats.TalentType = unique_matched_types;
profile_stats.Count = zeros(n_types, 1);
profile_stats.CompetencyMean = zeros(n_types, 1);
profile_stats.CompetencyStd = zeros(n_types, 1);
profile_stats.TotalScoreMean = zeros(n_types, 1);
profile_stats.PerformanceRank = zeros(n_types, 1);

for i = 1:n_types
    type_name = unique_matched_types{i};
    type_mask = strcmp(matched_talent_types, type_name);
    type_comp_data = matched_comp{type_mask, :};
    type_profiles(i, :) = nanmean(type_comp_data, 1);

    % 통계 정보 수집
    profile_stats.Count(i) = sum(type_mask);
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
end






% 상위 12개 주요 역량 선정 (분산 기준)
comp_variance = var(table2array(matched_comp), 0, 1, 'omitnan');
[~, var_idx] = sort(comp_variance, 'descend');
top_comp_idx = var_idx(1:min(12, length(var_idx)));
top_comp_names = valid_comp_cols(top_comp_idx);

% 전체 평균 프로파일 (단순평균)
overall_mean_profile = nanmean(table2array(matched_comp), 1);

%% 가중평균 계산 및 검증
fprintf('\n▶ 가중평균 baseline 계산 및 검증\n');

% 샘플 수 기반 가중평균 계산
sample_counts = profile_stats.Count;
total_samples = sum(sample_counts);
weighted_overall_mean = zeros(1, length(valid_comp_cols));

% 가중치 계산 및 검증
weights = sample_counts / total_samples;
weight_sum = sum(weights);

fprintf('  - 총 샘플 수: %d명\n', total_samples);
if abs(weight_sum - 1.0) < 1e-10
    fprintf('  - 가중치 합계: %.6f (정확성: ✓ 정확)\n', weight_sum);
else
    fprintf('  - 가중치 합계: %.6f (정확성: ✗ 오류)\n', weight_sum);
end

% 가중평균 계산
for i = 1:n_types
    weighted_overall_mean = weighted_overall_mean + weights(i) * type_profiles(i, :);
end

% 단순평균 vs 가중평균 차이 계산 (RMS)
profile_diff = overall_mean_profile - weighted_overall_mean;
rms_diff = sqrt(mean(profile_diff.^2));

fprintf('  - 단순평균 vs 가중평균 RMS 차이: %.4f\n', rms_diff);

% 샘플 불균형 경고
max_samples = max(sample_counts);
min_samples = min(sample_counts);
imbalance_ratio = max_samples / min_samples;

if imbalance_ratio >= 5
    fprintf('  ⚠ 샘플 불균형 경고: 최대/최소 비율 = %.1f\n', imbalance_ratio);
    fprintf('    최대: %d명, 최소: %d명\n', max_samples, min_samples);
    fprintf('     가중평균(weighted) baseline 사용을 권장합니다.\n');
else
    fprintf('  ✓ 샘플 균형: 최대/최소 비율 = %.1f\n', imbalance_ratio);
end

fprintf('  - 선택된 baseline 유형: %s\n', config.baseline_type);

% 통일된 스케일 범위 (데이터 기반, 10 단위로 확장)
all_profile_data = type_profiles(:, top_comp_idx);
data_min = min(all_profile_data(:));
data_max = max(all_profile_data(:));

% 10의 배수로 확장 (여유값 포함)
global_min = floor((data_min - 5) / 10) * 10;  % 10 단위로 내림
global_max = ceil((data_max + 5) / 10) * 10;   % 10 단위로 올림

fprintf('▶ 통일 스케일 범위: %.1f ~ %.1f (10 단위 확장)\n', global_min, global_max);
fprintf('  (실제 데이터: %.1f ~ %.1f)\n', data_min, data_max);
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

    % baseline 선택 (스위치 기반)
    if strcmp(config.baseline_type, 'weighted')
        baseline = weighted_overall_mean(top_comp_idx);
        baseline_text = '가중평균';
    else
        baseline = overall_mean_profile(top_comp_idx);
        baseline_text = '단순평균';
    end

    % 개선된 레이더 차트 그리기 (인라인 코드)
    data = type_profile;
    baseline_data = baseline;
    labels = top_comp_names;
    sample_count = profile_stats.Count(i);
    title_text = sprintf(unique_matched_types{i});
    color = colors(i,:);
    min_val = global_min;
    max_val = global_max;

    % 레이더 차트 생성
    n_vars = length(data);
    angles = linspace(0, 2*pi, n_vars+1);

    % 스케일 정규화 (50~90 범위로 클리핑)
    data_clipped = max(min(data, max_val), min_val);  % 범위 클리핑
    baseline_clipped = max(min(baseline_data, max_val), min_val);

    data_norm = (data_clipped - min_val) / (max_val - min_val);
    baseline_norm = (baseline_clipped - min_val) / (max_val - min_val);

    % 순환을 위해 첫 번째 값을 마지막에 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];

    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);

    hold on;

    % 그리드 그리기 (고정 눈금: min_val부터 max_val까지 10 단위)
    grid_values = min_val:10:max_val;

    for j = 1:length(grid_values)
        grid_value = grid_values(j);
        r = (grid_value - min_val) / (max_val - min_val);

        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.8, 'Color', [0.6 0.6 0.6], 'HandleVisibility', 'off');

        % 그리드 레이블
        text(0, r, sprintf('%.0f', grid_value), ...
             'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'bottom', ...
             'FontSize', 9, 'Color', [0.4 0.4 0.4]);
    end

    % 방사선 그리기
    for j = 1:n_vars
        plot([0, cos(angles(j))], [0, sin(angles(j))], 'k:', ...
             'LineWidth', 0.8, 'Color', [0.6 0.6 0.6], 'HandleVisibility', 'off');
    end

    % 기준선 (범례 포함)
    h_baseline = plot(x_base, y_base, '--', 'Color', [0.7 0.7 0.7], 'LineWidth', 2, ...
        'DisplayName', '평균선');

    % 데이터 플롯 (범례 포함)
    h_data = patch(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2.5, ...
        'DisplayName', '해당 유형');

    % 데이터 포인트 (범례에서 제외)
    scatter(x_data(1:end-1), y_data(1:end-1), 60, color, 'filled', ...
            'MarkerEdgeColor', 'white', 'LineWidth', 1, 'HandleVisibility', 'off');

    % 레이블 및 값
    label_radius = 1.25;
    for j = 1:n_vars
        [lx, ly] = pol2cart(angles(j), label_radius);

        % 차이값 계산
        diff_val = data(j) - baseline_data(j);
        diff_str = sprintf('%+.1f', diff_val);
        if diff_val > 0
            diff_color = [0, 0.5, 0];
        else
            diff_color = [0.8, 0, 0];
        end

        text(lx, ly, sprintf('%s\n%.1f\n(%s)', labels{j}, data(j), diff_str), ...
             'HorizontalAlignment', 'center', ...
             'FontSize', 10, 'FontWeight', 'bold');
    end

    % 제목
    title(title_text, 'FontSize', 16, 'FontWeight', 'bold');

    % 개선된 범례 (특정 핸들만 사용)
    legend([h_baseline, h_data], 'Location', 'northeast', 'FontSize', 10, ...
           'Box', 'on', 'TextColor', 'black');

    axis equal;
    axis([-1.5 1.5 -1.5 1.5]);
    axis off;
    hold off;

    % 추가 정보 표시
    if config.performance_ranking.isKey(unique_matched_types{i})
        perf_rank = config.performance_ranking(unique_matched_types{i});
        text(0.5, -0.05, sprintf('CODE: %d', perf_rank), ...
             'Units', 'normalized', ...
             'HorizontalAlignment', 'center', ...
             'FontWeight', 'bold', 'FontSize', 14);
    end

    % 샘플 크기 표시 (항상 표시)
    if sample_count < 10
        % 샘플 수가 적을 때는 경고 색상으로 표시
        text(0.5, -0.12, sprintf('샘플 수: n=%d (⚠ 주의: 샘플 수 부족)', sample_count), ...
             'Units', 'normalized', ...
             'HorizontalAlignment', 'center', ...
             'FontWeight', 'bold', 'FontSize', 12, ...
             'Color', [0.8, 0, 0]);
    else
        % 샘플 수가 충분할 때는 일반 색상으로 표시
        text(0.5, -0.12, sprintf('샘플 수: n=%d', sample_count), ...
             'Units', 'normalized', ...
             'HorizontalAlignment', 'center', ...
             'FontWeight', 'bold', 'FontSize', 12, ...
             'Color', [0.2, 0.2, 0.2]);
    end

    % Figure 저장
    % saveas(fig, sprintf('radar_chart_%s_%s.png', ...
    %        strrep(unique_matched_types{i}, ' ', '_'), config.timestamp));

    fprintf('  ✓ %s 차트 생성 완료\n', unique_matched_types{i});
end

%% ========================================================================
%                    PART 3: 고도화된 상관 기반 가중치 분석
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║              PART 3: 고도화된 상관 기반 가중치 분석      ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% 3.1 성과점수 기반 상관분석
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

% 상관계수 계산
n_competencies = length(valid_comp_cols);
correlation_results = table();
correlation_results.Competency = valid_comp_cols';
correlation_results.Correlation = zeros(n_competencies, 1);
correlation_results.PValue = zeros(n_competencies, 1);
correlation_results.Significance = cell(n_competencies, 1);
correlation_results.HighPerf_Mean = zeros(n_competencies, 1);
correlation_results.LowPerf_Mean = zeros(n_competencies, 1);
correlation_results.Difference = zeros(n_competencies, 1);
correlation_results.EffectSize = zeros(n_competencies, 1);

% 성과 상위/하위 그룹 분류 (상위 25%, 하위 25%)
perf_q75 = quantile(valid_performance, 0.75);
perf_q25 = quantile(valid_performance, 0.25);
high_perf_idx = valid_performance >= perf_q75;
low_perf_idx = valid_performance <= perf_q25;

for i = 1:n_competencies
    comp_scores = valid_competencies(:, i);
    valid_idx = ~isnan(comp_scores);

    if sum(valid_idx) >= 10
        % 상관계수 계산 (Spearman)
        [r, p] = corr(comp_scores(valid_idx), valid_performance(valid_idx), 'Type', 'Spearman');
        correlation_results.Correlation(i) = r;
        correlation_results.PValue(i) = p;

        % 유의성 표시
        if p < 0.001
            correlation_results.Significance{i} = '***';
        elseif p < 0.01
            correlation_results.Significance{i} = '**';
        elseif p < 0.05
            correlation_results.Significance{i} = '*';
        else
            correlation_results.Significance{i} = '';
        end

        % 그룹별 평균
        correlation_results.HighPerf_Mean(i) = nanmean(comp_scores(high_perf_idx));
        correlation_results.LowPerf_Mean(i) = nanmean(comp_scores(low_perf_idx));
        correlation_results.Difference(i) = correlation_results.HighPerf_Mean(i) - ...
                                           correlation_results.LowPerf_Mean(i);

        % Effect Size (Cohen's d)
        high_scores = comp_scores(high_perf_idx & valid_idx);
        low_scores = comp_scores(low_perf_idx & valid_idx);
        if length(high_scores) > 1 && length(low_scores) > 1
            pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                              (length(low_scores)-1)*var(low_scores)) / ...
                              (length(high_scores) + length(low_scores) - 2));
            correlation_results.EffectSize(i) = (mean(high_scores) - mean(low_scores)) / pooled_std;
        end
    end
end

% 가중치 계산
positive_corr = max(0, correlation_results.Correlation);
weights_corr = positive_corr / (sum(positive_corr) + eps);
correlation_results.Weight = weights_corr * 100;

correlation_results = sortrows(correlation_results, 'Correlation', 'descend');

fprintf('\n상위 10개 성과 예측 역량:\n');
fprintf('%-25s | 상관계수 | p-값 | 효과크기 | 가중치(%%)\n', '역량');
fprintf('%s\n', repmat('-', 75, 1));

for i = 1:min(10, height(correlation_results))
    fprintf('%-25s | %8.4f%s | %6.4f | %8.2f | %7.2f\n', ...
        correlation_results.Competency{i}, ...
        correlation_results.Correlation(i), correlation_results.Significance{i}, ...
        correlation_results.PValue(i), ...
        correlation_results.EffectSize(i), ...
        correlation_results.Weight(i));
end

%% 3.3 상관분석 시각화
% Figure 2: 상관분석 결과
colors_vis = struct('primary', [0.2, 0.4, 0.8], 'secondary', [0.8, 0.3, 0.2], ...
               'tertiary', [0.3, 0.7, 0.4], 'gray', [0.5, 0.5, 0.5]);

fig2 = figure('Position', [100, 100, 1600, 900], 'Color', 'white');

% 상위 15개 역량의 상관계수와 가중치
subplot(2, 2, [1, 2]);
top_15 = correlation_results(1:min(15, height(correlation_results)), :);
x = 1:height(top_15);

yyaxis left
bar(x, top_15.Correlation, 'FaceColor', colors_vis.primary, 'EdgeColor', 'none');
ylabel('상관계수', 'FontSize', 12, 'FontWeight', 'bold');
ylim([-0.2, max(top_15.Correlation)*1.2]);

yyaxis right
plot(x, top_15.Weight, '-o', 'Color', colors_vis.secondary, 'LineWidth', 2, ...
     'MarkerFaceColor', colors_vis.secondary, 'MarkerSize', 8);
ylabel('가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');

set(gca, 'XTick', x, 'XTickLabel', top_15.Competency, 'XTickLabelRotation', 45);
xlabel('역량 항목', 'FontSize', 12, 'FontWeight', 'bold');
title('역량-성과 상관분석 및 가중치', 'FontSize', 14, 'FontWeight', 'bold');
grid on;
box off;

% 누적 가중치
subplot(2, 2, [3 ,4]);
cumulative_weight = cumsum(correlation_results.Weight);
plot(cumulative_weight, 'LineWidth', 2.5, 'Color', colors_vis.tertiary);
hold on;
plot([1, length(cumulative_weight)], [80, 80], '--', 'Color', colors_vis.gray, 'LineWidth', 2);
xlabel('역량 개수', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('누적 가중치 (%)', 'FontSize', 12, 'FontWeight', 'bold');
title('누적 설명력 분석', 'FontSize', 14, 'FontWeight', 'bold');
legend('누적 가중치', '80% 기준선', 'Location', 'southeast', 'FontSize', 10);
grid on;
box off;

sgtitle('역량-성과 상관분석 종합 결과', 'FontSize', 16, 'FontWeight', 'bold');

%% 【STEP 8-1】 7개 인재유형 간 마할라노비스 거리 분석
fprintf('\n【STEP 8-1】 7개 인재유형 간 마할라노비스 거리 분석\n');
fprintf('────────────────────────────────────────────\n\n');

% 유형별 데이터 준비
unique_types = unique(matched_talent_types);
n_types = length(unique_types);

% 각 유형의 샘플 수 확인
type_counts = zeros(n_types, 1);
for i = 1:n_types
    type_counts(i) = sum(strcmp(matched_talent_types, unique_types{i}));
end

fprintf('인재유형별 샘플 수:\n');
for i = 1:n_types
    fprintf('  %s: %d명\n', unique_types{i}, type_counts(i));
end
fprintf('\n');

% 각 유형의 평균 프로파일 계산
type_means = zeros(n_types, length(valid_comp_cols));
type_covs = cell(n_types, 1);

for i = 1:n_types
    type_mask = strcmp(matched_talent_types, unique_types{i});
    type_data = table2array(matched_comp(type_mask, :));

    % 결측값 처리 (평균으로 대체)
    for j = 1:size(type_data, 2)
        missing = isnan(type_data(:, j));
        if any(missing)
            type_data(missing, j) = nanmean(type_data(:, j));
        end
    end

    type_means(i, :) = mean(type_data, 1);

    % 샘플이 3개 이상인 경우만 공분산 계산
    if type_counts(i) >= 3
        type_covs{i} = cov(type_data);
    else
        type_covs{i} = eye(size(type_data, 2));  % 단위행렬 사용
    end
end

% 그룹별 pooled 공분산 행렬 계산 (STEP 9-1과 동일한 방식)
fprintf('그룹별 공분산 행렬 계산 중...\n');

% 유효한 유형들만 사용 (샘플 수 3개 이상)
valid_types_idx = type_counts >= 3;
valid_types = unique_types(valid_types_idx);
valid_type_means = type_means(valid_types_idx, :);
valid_type_covs = type_covs(valid_types_idx);
valid_type_counts = type_counts(valid_types_idx);

if sum(valid_types_idx) < 2
    warning('유효한 유형이 2개 미만입니다. 전체 데이터 공분산을 사용합니다.');
    X_all = table2array(matched_comp);
    for j = 1:size(X_all, 2)
        missing = isnan(X_all(:, j));
        if any(missing)
            X_all(missing, j) = nanmean(X_all(:, j));
        end
    end
    pooled_cov = cov(X_all);
    pooled_cov_reg = pooled_cov + eye(size(pooled_cov)) * 1e-6;

    % 모든 유형 간 거리 계산
    distance_matrix = zeros(n_types, n_types);
    for i = 1:n_types
        for j = i+1:n_types
            diff = type_means(i, :) - type_means(j, :);
            % 백슬래시 연산자 사용 (더 안전함)
            distance_matrix(i, j) = sqrt(diff * (pooled_cov_reg \ diff'));
            distance_matrix(j, i) = distance_matrix(i, j);
        end
    end
else
    % 유효한 유형들의 pooled 공분산 계산
    total_samples = sum(valid_type_counts);
    pooled_cov = zeros(length(valid_comp_cols), length(valid_comp_cols));

    for i = 1:length(valid_types)
        weight = (valid_type_counts(i) - 1) / (total_samples - length(valid_types));
        pooled_cov = pooled_cov + weight * valid_type_covs{i};
    end

    % 정규화 추가 (특이점 방지)
    pooled_cov_reg = pooled_cov + eye(size(pooled_cov)) * 1e-6;

    % 모든 유형 간 마할라노비스 거리 계산
    distance_matrix = zeros(n_types, n_types);

    fprintf('유형 간 마할라노비스 거리 계산 중...\n');
    for i = 1:n_types
        for j = i+1:n_types
            diff = type_means(i, :) - type_means(j, :);

            try
                % Cholesky 분해 사용 (더 안전하고 빠름)
                L = chol(pooled_cov_reg, 'lower');
                v = L \ diff';
                distance_matrix(i, j) = sqrt(v' * v);
                distance_matrix(j, i) = distance_matrix(i, j);
            catch chol_error
                warning('Cholesky 분해 실패 (%s vs %s): %s. pinv 사용.', unique_types{i}, unique_types{j}, chol_error.message);
                % pinv로 대체
                try
                    distance_matrix(i, j) = sqrt(diff * pinv(pooled_cov_reg) * diff');
                    distance_matrix(j, i) = distance_matrix(i, j);
                catch
                    % 유클리드 거리로 대체
                % 유클리드 거리로 대체
                distance_matrix(i, j) = sqrt(sum(diff.^2));
                distance_matrix(j, i) = distance_matrix(i, j);
            end
        end
    end
end

fprintf('마할라노비스 거리 계산 완료\n');

% 거리 행렬 출력
fprintf('\n【마할라노비스 거리 행렬】\n');
fprintf('%-20s', '');
for i = 1:n_types
    fprintf('%8s', unique_types{i}(1:min(7,end)));
end
fprintf('\n');

for i = 1:n_types
    fprintf('%-20s', unique_types{i});
    for j = 1:n_types
        if i == j
            fprintf('%8s', '-');
        else
            fprintf('%8.2f', distance_matrix(i, j));
        end
    end
    fprintf('\n');
end
end
% 성과 순위와 거리의 관계 분석
fprintf('\n【성과 순위와 거리 관계】\n');
fprintf('────────────────────────────────────────────\n');

% 성과 순위 맵핑
performance_ranks = zeros(n_types, 1);
for i = 1:n_types
    if config.performance_ranking.isKey(unique_types{i})
        performance_ranks(i) = config.performance_ranking(unique_types{i});
    else
        performance_ranks(i) = 0;  % 순위 없음
    end
end

% 순위 차이와 거리의 상관관계
rank_diffs = [];
distances = [];
for i = 1:n_types
    for j = i+1:n_types
        if performance_ranks(i) > 0 && performance_ranks(j) > 0
            rank_diffs(end+1) = abs(performance_ranks(i) - performance_ranks(j));
            distances(end+1) = distance_matrix(i, j);
        end
    end
end

if ~isempty(rank_diffs)
    correlation = corr(rank_diffs', distances', 'Type', 'Spearman');
    fprintf('성과 순위 차이와 마할라노비스 거리의 상관계수: %.3f\n', correlation);

    if correlation > 0.5
        fprintf('→ 성과가 다를수록 역량 프로파일도 다름 (타당한 분류)\n');
    elseif correlation > 0
        fprintf('→ 약한 양의 관계 (부분적 타당성)\n');
    else
        fprintf('→ 성과와 역량 프로파일이 일치하지 않음 (재검토 필요)\n');
    end
end

% 클러스터링 가능성 분석
fprintf('\n【유형 그룹핑 분석】\n');
fprintf('────────────────────────────────────────────\n');

% 거리 기준 그룹핑 (임계값: 1.0)
threshold = 1.0;
fprintf('거리 %.1f 이하로 묶이는 그룹:\n', threshold);

visited = false(n_types, 1);
group_num = 0;

for i = 1:n_types
    if ~visited(i)
        group_num = group_num + 1;
        group_members = unique_types(i);
        visited(i) = true;

        for j = i+1:n_types
            if ~visited(j) && distance_matrix(i, j) < threshold
                group_members = [group_members; unique_types(j)];
                visited(j) = true;
            end
        end

        if length(group_members) > 1
            fprintf('\n그룹 %d:\n', group_num);
            for k = 1:length(group_members)
                fprintf('  - %s (CODE: %d)\n', group_members{k}, ...
                    config.performance_ranking(group_members{k}));
            end

            % 그룹 내 평균 거리
            if length(group_members) == 2
                idx1 = find(strcmp(unique_types, group_members{1}));
                idx2 = find(strcmp(unique_types, group_members{2}));
                avg_dist = distance_matrix(idx1, idx2);
            else
                group_dists = [];
                for m = 1:length(group_members)-1
                    for n = m+1:length(group_members)
                        idx1 = find(strcmp(unique_types, group_members{m}));
                        idx2 = find(strcmp(unique_types, group_members{n}));
                        group_dists(end+1) = distance_matrix(idx1, idx2);
                    end
                end
                avg_dist = mean(group_dists);
            end
            fprintf('  평균 거리: %.2f\n', avg_dist);
        end
    end
end

% 가장 가까운/먼 유형 쌍 찾기 (수정된 방법)
U = triu(true(n_types), 1);
distance_vector = distance_matrix(U);

% 0이 아닌 거리만 고려
valid_distances = distance_vector(distance_vector > 0);
if ~isempty(valid_distances)
    [min_dist, ~] = min(valid_distances);
    [max_dist, ~] = max(valid_distances);

    % 원래 인덱스 찾기
    [i_all, j_all] = find(U);
    min_idx = find(distance_vector == min_dist, 1);
    max_idx = find(distance_vector == max_dist, 1);

    min_i = i_all(min_idx);
    min_j = j_all(min_idx);
    max_i = i_all(max_idx);
    max_j = j_all(max_idx);
else
    % 예외 처리
    warning('유효한 거리를 찾을 수 없습니다.');
    min_dist = NaN; max_dist = NaN;
    min_i = 1; min_j = 2; max_i = 1; max_j = 2;
end

fprintf('\n【극단 케이스】\n');
fprintf('가장 유사한 유형 쌍:\n');
fprintf('  %s ↔ %s (거리: %.2f)\n', unique_types{min_i}, unique_types{min_j}, min_dist);

fprintf('\n가장 다른 유형 쌍:\n');
fprintf('  %s ↔ %s (거리: %.2f)\n', unique_types{max_i}, unique_types{max_j}, max_dist);

% 시각화: 히트맵
try
    fprintf('\n히트맵 생성 중...\n');
    fig_heatmap = figure('Position', [100, 100, 800, 700], 'Color', 'white');

    % 거리 행렬 시각화
    imagesc(distance_matrix);
    colormap('hot');
    c = colorbar;
    c.Label.String = '마할라노비스 거리';
    c.Label.FontSize = 12;

    title('인재유형 간 마할라노비스 거리 히트맵', 'FontSize', 14, 'FontWeight', 'bold');
    xlabel('인재유형', 'FontWeight', 'bold', 'FontSize', 12);
    ylabel('인재유형', 'FontWeight', 'bold', 'FontSize', 12);

    % 축 레이블 설정 (짧은 이름 사용)
    short_names = cell(n_types, 1);
    for i = 1:n_types
        if length(unique_types{i}) > 8
            short_names{i} = unique_types{i}(1:8);
        else
            short_names{i} = unique_types{i};
        end
    end

    set(gca, 'XTick', 1:n_types, 'XTickLabel', short_names, 'XTickLabelRotation', 45);
    set(gca, 'YTick', 1:n_types, 'YTickLabel', short_names);
    set(gca, 'FontSize', 10);

    % 값 표시 (대각선 제외)
    for i = 1:n_types
        for j = 1:n_types
            if i ~= j
                distance_val = distance_matrix(i, j);
                if distance_val > 0  % 유효한 거리만 표시
                    text(j, i, sprintf('%.1f', distance_val), ...
                        'HorizontalAlignment', 'center', 'Color', 'black', ...
                        'FontWeight', 'bold', 'FontSize', 9);
                end
            end
        end
    end

    % % 저장
    % heatmap_filename = sprintf('type_distance_heatmap_%s.png', config.timestamp);
    % try
    %     saveas(gcf, heatmap_filename);
    %     fprintf('히트맵 저장 완료: %s\n', heatmap_filename);
    % catch save_error
    %     warning('히트맵 저장 실패: %s', save_error.message);
    % end

catch plot_error
    warning('히트맵 생성 실패: %s', plot_error.message);
    fprintf('→ 텍스트 결과만 제공됩니다.\n');
end

fprintf('\n✅ 7개 유형 간 마할라노비스 거리 분석 완료\n');


%% ========================================================================
%        PART 4: Cost-Sensitive Learning 기반 고성과자 예측 시스템
% =========================================================================

fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║    PART 4: Cost-Sensitive Learning 기반 고성과자 예측     ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');


%% 4.1 데이터 준비 및 클래스 불균형 해결
fprintf('【STEP 9】 데이터 준비 및 클래스 불균형 분석\n');
fprintf('────────────────────────────────────────────\n');

% 이진분류를 위한 명확한 그룹 정의
fprintf('분류 기준 재정의:\n');
fprintf('  고성과자: %s\n', strjoin(config.high_performers, ', '));
fprintf('  저성과자: %s\n', strjoin(config.low_performers, ', '));
fprintf('  분석 제외: %s\n', strjoin(config.excluded_from_analysis, ', '));

% 원본 역량 데이터 품질 확인
X_raw = table2array(matched_comp);
fprintf('\n원본 데이터 확인:\n');
fprintf('  샘플 수: %d\n', size(X_raw, 1));
fprintf('  역량 수: %d\n', size(X_raw, 2));
fprintf('  결측값 비율: %.1f%%\n', sum(isnan(X_raw(:)))/numel(X_raw)*100);

% 유능한 불연성을 제외하고 레이블 생성
y_binary = NaN(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    type_name = matched_talent_types{i};
    if any(strcmp(type_name, config.high_performers))
        y_binary(i) = 1;  % 고성과자
    elseif any(strcmp(type_name, config.low_performers))
        y_binary(i) = 0;  % 저성과자
    elseif any(strcmp(type_name, config.excluded_from_analysis))
        y_binary(i) = NaN;  % 분석에서 제외
    end
end

% % 유능한 불연성 제외 후 데이터 확인
excluded_count = sum(strcmp(matched_talent_types, '유능한 불연성'));
fprintf('\n유능한 불연성 %d명을 분석에서 제외\n', excluded_count);

%% STEP 9-1: 마할라노비스 거리 기반 그룹 분리 타당성 검증
fprintf('\n=== STEP 9-1: 마할라노비스 거리 기반 그룹 분리 타당성 검증 ===\n');

try
    % 분석 대상 데이터 준비 (NaN 제외)
    valid_idx = ~isnan(y_binary);
    X_for_mahal = X_raw(valid_idx, :);
    y_for_mahal = y_binary(valid_idx);

    % 완전한 케이스만 사용
    complete_cases = ~any(isnan(X_for_mahal), 2);
    X_complete = X_for_mahal(complete_cases, :);
    y_complete = y_for_mahal(complete_cases);

    if sum(complete_cases) < 10
        warning('완전한 케이스가 너무 적습니다 (%d개). 마할라노비스 거리 분석을 건너뜁니다.', sum(complete_cases));
    else
        % 고성과자와 저성과자 그룹 분리
        high_perf_idx = y_complete == 1;
        low_perf_idx = y_complete == 0;

        X_high = X_complete(high_perf_idx, :);
        X_low = X_complete(low_perf_idx, :);

        fprintf('\n그룹별 샘플 수:\n');
        fprintf('  고성과자: %d명\n', size(X_high, 1));
        fprintf('  저성과자: %d명\n', size(X_low, 1));

        if size(X_high, 1) >= 3 && size(X_low, 1) >= 3
            % 전체 공분산 행렬 계산 (pooled covariance)
            n_high = size(X_high, 1);
            n_low = size(X_low, 1);

            % 각 그룹의 공분산 행렬
            cov_high = cov(X_high);
            cov_low = cov(X_low);

            % 공통 공분산 행렬 (pooled covariance)
            pooled_cov = ((n_high - 1) * cov_high + (n_low - 1) * cov_low) / (n_high + n_low - 2);

            % 그룹 중심점 계산
            mean_high = mean(X_high);
            mean_low = mean(X_low);

            % 마할라노비스 거리 계산
            mean_diff = mean_high - mean_low;

            % 마할라노비스 거리 계산 (안전한 방법)
            pooled_cov_reg = pooled_cov + eye(size(pooled_cov)) * 1e-6;
            try
                % Cholesky 분해 사용
                L = chol(pooled_cov_reg, 'lower');
                v = L \ mean_diff';
                mahal_distance_squared = v' * v;
                mahal_distance = sqrt(mahal_distance_squared);

                fprintf('\n마할라노비스 거리 분석 결과:\n');
                fprintf('  마할라노비스 거리: %.4f\n', mahal_distance);
                fprintf('  거리 제곱: %.4f\n', mahal_distance_squared);

                % 해석 기준
                fprintf('\n해석 기준:\n');
                if mahal_distance >= 3.0
                    fprintf('  ✓ 매우 우수한 그룹 분리 (D² ≥ 3.0)\n');
                    separation_quality = 'excellent';
                elseif mahal_distance >= 2.0
                    fprintf('  ✓ 우수한 그룹 분리 (D² ≥ 2.0)\n');
                    separation_quality = 'good';
                elseif mahal_distance >= 1.5
                    fprintf('  △ 보통 수준의 그룹 분리 (D² ≥ 1.5)\n');
                    separation_quality = 'moderate';
                elseif mahal_distance >= 1.0
                    fprintf('  ⚠ 약한 그룹 분리 (D² ≥ 1.0)\n');
                    separation_quality = 'weak';
                    warning('그룹 간 분리가 약합니다. 분류 모델의 성능이 제한적일 수 있습니다.');
                else
                    fprintf('  ✗ 매우 약한 그룹 분리 (D² < 1.0)\n');
                    separation_quality = 'very_weak';
                    warning('그룹 간 분리가 매우 약합니다. 분류 모델 적용을 재검토해야 합니다.');
                end

                % 효과 크기 계산 (Cohen's d equivalent for multivariate)
                effect_size = mahal_distance / sqrt(size(X_complete, 2));
                fprintf('  다변량 효과 크기: %.4f\n', effect_size);

                if effect_size >= 0.8
                    fprintf('  → 큰 효과 크기 (≥ 0.8)\n');
                elseif effect_size >= 0.5
                    fprintf('  → 중간 효과 크기 (≥ 0.5)\n');
                elseif effect_size >= 0.2
                    fprintf('  → 작은 효과 크기 (≥ 0.2)\n');
                else
                    fprintf('  → 매우 작은 효과 크기 (< 0.2)\n');
                end

                % 통계적 유의성 검정 (Hotelling's T² test)
                n_total = n_high + n_low;
                hotelling_t2 = (n_high * n_low / n_total) * mahal_distance_squared;

                % F-통계량으로 변환
                p_features = size(X_complete, 2);
                f_stat = ((n_total - p_features - 1) / ((n_total - 2) * p_features)) * hotelling_t2;
                df1 = p_features;
                df2 = n_total - p_features - 1;

                % F-분포를 이용한 p-값 계산 (근사치)
                if df2 > 0
                    p_value_approx = 1 - fcdf(f_stat, df1, df2);
                    fprintf('\n통계적 유의성 검정 (Hotelling''s T²):\n');
                    fprintf('  F-통계량: %.4f (df1=%d, df2=%d)\n', f_stat, df1, df2);
                    fprintf('  p-값 (근사): %.6f\n', p_value_approx);

                    if p_value_approx < 0.001
                        fprintf('  ✓ 매우 유의한 그룹 차이 (p < 0.001)\n');
                    elseif p_value_approx < 0.01
                        fprintf('  ✓ 유의한 그룹 차이 (p < 0.01)\n');
                    elseif p_value_approx < 0.05
                        fprintf('  ✓ 유의한 그룹 차이 (p < 0.05)\n');
                    else
                        fprintf('  ⚠ 그룹 차이가 유의하지 않음 (p ≥ 0.05)\n');
                    end
                end

                % 개별 변수별 기여도 분석
                fprintf('\n변수별 그룹 분리 기여도 (상위 5개):\n');
                try
                    % 안전한 diag 계산
                    L = chol(pooled_cov_reg, 'lower');
                    inv_diag = 1 ./ sum(L.^2, 1);  % 대각원소 근사
                    individual_contributions = abs(mean_diff .* inv_diag);
                catch
                    % 대체 방법: pinv의 대각원소
                    individual_contributions = abs(mean_diff .* diag(pinv(pooled_cov_reg))');
                end
                [sorted_contrib, sorted_idx] = sort(individual_contributions, 'descend');

                for i = 1:min(5, length(sorted_contrib))
                    var_idx = sorted_idx(i);
                    fprintf('  %s: %.4f\n', valid_comp_cols{var_idx}, sorted_contrib(i));
                end

                % 결과 저장
                mahalanobis_results = struct();
                mahalanobis_results.distance = mahal_distance;
                mahalanobis_results.distance_squared = mahal_distance_squared;
                mahalanobis_results.separation_quality = separation_quality;
                mahalanobis_results.effect_size = effect_size;
                mahalanobis_results.n_high_performers = n_high;
                mahalanobis_results.n_low_performers = n_low;
                mahalanobis_results.variable_contributions = individual_contributions;

                if exist('p_value_approx', 'var')
                    mahalanobis_results.p_value = p_value_approx;
                    mahalanobis_results.f_statistic = f_stat;
                end

                % 권장사항 출력
                fprintf('\n분석 권장사항:\n');
                if strcmp(separation_quality, 'excellent') || strcmp(separation_quality, 'good')
                    fprintf('  ✓ 그룹 분리가 우수합니다. 분류 모델 적용에 적합합니다.\n');
                elseif strcmp(separation_quality, 'moderate')
                    fprintf('  △ 보통 수준의 분리입니다. 모델 성능을 주의깊게 모니터링하세요.\n');
                else
                    fprintf('  ⚠ 그룹 분리가 약합니다. 다음을 고려하세요:\n');
                    fprintf('    - 추가적인 특성 엔지니어링\n');
                    fprintf('    - 더 정교한 분류 기준 재검토\n');
                    fprintf('    - 비선형 모델 적용 검토\n');
                end

            catch chol_error
                warning('Cholesky 분해 실패: %s. pinv 사용.', chol_error.message);
                % pinv로 대체
                try
                    mahal_distance_squared = mean_diff * pinv(pooled_cov_reg) * mean_diff';
                    mahal_distance = sqrt(mahal_distance_squared);
                catch
                    warning('모든 마할라노비스 계산 실패. 유클리드 거리 사용.');
                    mahal_distance = sqrt(sum(mean_diff.^2));
                    mahal_distance_squared = mahal_distance^2;
                end
                fprintf('  → 차원 축소나 정규화를 고려하세요.\n');
            end
        else
            warning('각 그룹에 최소 3개 이상의 샘플이 필요합니다.');
        end
    end

catch mahal_error
    warning('마할라노비스 거리 분석 중 오류 발생: %s', mahal_error.message);
    fprintf('→ 다음 단계로 진행합니다.\n');
end

fprintf('\n【2단계 결측값 처리】\n');

% Step 1: 고품질 샘플만 선택 (결측률 50% 미만으로 완화)
missing_rate_per_sample = sum(isnan(X_raw), 2) / size(X_raw, 2);
quality_threshold = 0.5;  % 30% → 50%로 완화하여 데이터 손실 최소화
quality_samples = missing_rate_per_sample < quality_threshold;

fprintf('Step 1 - 고품질 샘플 선택 (결측률 임계값 완화):\n');
fprintf('  결측률 50%% 이상 제거: %d명 제외\n', sum(~quality_samples));
fprintf('  남은 샘플: %d명\n', sum(quality_samples));
fprintf('  → 30%% 임계값 대비 추가 보존: %d명\n', sum(quality_samples) - sum(missing_rate_per_sample < 0.3));

% 제거된 샘플의 인재유형 분포 출력
if sum(~quality_samples) > 0
    fprintf('\n   결측률 필터링으로 제거된 인원의 인재유형 분포:\n');
    removed_types = matched_talent_types(~quality_samples);
    [removed_unique, ~, removed_idx] = unique(removed_types);
    removed_counts = accumarray(removed_idx, 1);
    for i = 1:length(removed_unique)
        fprintf('    • %-20s: %2d명\n', removed_unique{i}, removed_counts(i));
    end
end

% 유능한 불연성 제외 + 품질 필터 적용
binary_valid_idx = ~isnan(y_binary);
final_idx = binary_valid_idx & quality_samples;
X_quality = X_raw(final_idx, :);
y_quality = y_binary(final_idx);

% Step 2: 결측치 대체 (평균값 대체 방식)
% 기존: 완전한 케이스만 사용 → 데이터 손실 과다
% 개선: 역량별 평균값으로 결측치 대체
fprintf('\nStep 2 - 결측치 대체 (평균값 대체):\n');

X_imputed = X_quality;
n_missing_samples = 0;
n_missing_total = 0;

for col = 1:size(X_quality, 2)
    col_data = X_quality(:, col);
    missing_idx = isnan(col_data);
    n_missing = sum(missing_idx);

    if n_missing > 0
        n_missing_samples = n_missing_samples + 1;
        n_missing_total = n_missing_total + n_missing;

        % 해당 컬럼의 평균값으로 대체 (그룹별 평균 사용)
        for grp = [0, 1]  % 저성과자(0), 고성과자(1)
            grp_idx = (y_quality == grp);
            grp_missing = missing_idx & grp_idx;

            if sum(grp_missing) > 0
                % 그룹 내 평균값 계산
                grp_mean = nanmean(col_data(grp_idx));

                % 그룹 평균이 없으면 전체 평균 사용
                if isnan(grp_mean)
                    grp_mean = nanmean(col_data);
                end

                % 결측치 대체
                X_imputed(grp_missing, col) = grp_mean;
            end
        end
    end
end

fprintf('  결측치가 있던 역량: %d개\n', n_missing_samples);
fprintf('  대체된 총 결측치 수: %d개\n', n_missing_total);
fprintf('  최종 분석 샘플: %d명 (제거 없음)\n', size(X_imputed, 1));

% 최종 데이터 설정
X_final = X_imputed;
y_final = y_quality;

fprintf('  - 고성과자: %d명\n', sum(y_final == 1));
fprintf('  - 저성과자: %d명\n', sum(y_final == 0));
fprintf('  - 샘플/변수 비율: %.1f\n', length(y_final)/size(X_final, 2));

% complete_cases를 모두 true로 설정 (모든 샘플 유지)
complete_cases = true(size(y_quality));

y_weight = y_final;

% 클래스 분포 확인
n_high = sum(y_weight == 1);
n_low = sum(y_weight == 0);
total_binary = length(y_weight);

fprintf('\n최종 이진분류 데이터 분포:\n');
fprintf('  고성과자 (1): %d명 (%.1f%%)\n', n_high, n_high/total_binary*100);
fprintf('  저성과자 (0): %d명 (%.1f%%)\n', n_low, n_low/total_binary*100);
fprintf('  불균형 비율: %.2f:1\n', max(n_high, n_low)/min(n_high, n_low));

% 최종 분석 대상 인재유형별 분포 출력
fprintf('\n【최종 분석 대상 인재유형 분포】\n');
fprintf('────────────────────────────────────────────\n');
final_talent_types_temp = matched_talent_types(final_idx);
final_talent_types_temp = final_talent_types_temp(complete_cases);
[final_unique_types, ~, final_type_indices] = unique(final_talent_types_temp);
final_type_counts = accumarray(final_type_indices, 1);

fprintf('총 %d명 (결측치 필터링 완료):\n\n', length(final_talent_types_temp));

% 고성과자 그룹
fprintf('✅ 고성과자 그룹:\n');
for i = 1:length(config.high_performers)
    type_name = config.high_performers{i};
    type_idx = strcmp(final_unique_types, type_name);
    if any(type_idx)
        count = final_type_counts(type_idx);
        fprintf('  • %-20s: %3d명 (%5.1f%%)\n', type_name, count, count/length(final_talent_types_temp)*100);
    else
        fprintf('  • %-20s: %3d명 (%5.1f%%)\n', type_name, 0, 0.0);
    end
end

% 저성과자 그룹
fprintf('\n❌ 저성과자 그룹:\n');
for i = 1:length(config.low_performers)
    type_name = config.low_performers{i};
    type_idx = strcmp(final_unique_types, type_name);
    if any(type_idx)
        count = final_type_counts(type_idx);
        fprintf('  • %-20s: %3d명 (%5.1f%%)\n', type_name, count, count/length(final_talent_types_temp)*100);
    else
        fprintf('  • %-20s: %3d명 (%5.1f%%)\n', type_name, 0, 0.0);
    end
end

% 제외된 그룹
if ~isempty(config.excluded_from_analysis)
    fprintf('\n 분석 제외 그룹 (이미 제거됨):\n');
    for i = 1:length(config.excluded_from_analysis)
        fprintf('  • %-20s: 제외됨\n', config.excluded_from_analysis{i});
    end
end
fprintf('────────────────────────────────────────────\n');

% 클래스 가중치 계산 (inverse frequency weighting)
class_weights = length(y_weight) ./ (2 * [sum(y_weight==0), sum(y_weight==1)]);
sample_weights = zeros(size(y_weight));
sample_weights(y_weight==0) = class_weights(1);  % 저성과자
sample_weights(y_weight==1) = class_weights(2);  % 고성과자

fprintf('  클래스 가중치 - 저성과자: %.3f, 고성과자: %.3f\n', class_weights(1), class_weights(2));

% 비용 행렬 정의 (저성과자→고성과자 오분류 비용 1.5배)
cost_matrix = [0 1; 1.5 0];  % [TN FP; FN TP]
fprintf('  비용 행렬: 저성과자→고성과자 오분류 비용 1.5배 적용\n');

%% STEP 9.5: 나이 및 성별 효과 통제
fprintf('\n【STEP 9.5】 인구통계학적 변수 효과 통제\n');
fprintf('────────────────────────────────────────────\n\n');

% '만 나이'와 '성별' 변수 추출
age_col_idx = find(strcmp(matched_hr.Properties.VariableNames, '만 나이'), 1);
gender_col_idx = find(strcmp(matched_hr.Properties.VariableNames, '성별'), 1);

% 나이 효과 분석
if ~isempty(age_col_idx)
    try
        ages = matched_hr{:, age_col_idx};
        % 나이 데이터가 숫자인지 확인
        if isnumeric(ages)
            ages_final = ages(final_idx);
            ages_final = ages_final(complete_cases);
            
            % 유효한 나이 데이터만 사용 (0-100 범위)
            valid_age_idx = ages_final > 0 & ages_final < 100 & ~isnan(ages_final);
            if sum(valid_age_idx) > 5
                ages_valid = ages_final(valid_age_idx);
                y_valid = y_final(valid_age_idx);
                
                % 나이-성과 상관관계
                age_perf_corr = corr(ages_valid, y_valid, 'Type', 'Spearman');
                fprintf('나이-성과 상관계수: %.3f\n', age_perf_corr);
            else
                fprintf('⚠ 유효한 나이 데이터가 부족합니다.\n');
                age_perf_corr = NaN;
            end
        else
            fprintf('⚠ 나이 데이터가 숫자가 아닙니다.\n');
            age_perf_corr = NaN;
        end
    catch ME
        fprintf('⚠ 나이 분석 중 에러 발생: %s\n', ME.message);
        age_perf_corr = NaN;
    end
else
    fprintf('⚠ "만 나이" 컬럼을 찾을 수 없습니다.\n');
    age_perf_corr = NaN;
end

% 성별 효과 분석
if ~isempty(gender_col_idx)
    try
        genders = matched_hr{:, gender_col_idx};
        genders_final = genders(final_idx);
        genders_final = genders_final(complete_cases);
        
        % 성별 더미변수 생성 (남성=1, 여성=0으로 가정)
        if iscell(genders_final)
            % 셀 배열인 경우 문자열 비교
            gender_dummy = strcmp(genders_final, '남성') | strcmp(genders_final, '남') | ...
                          strcmp(genders_final, 'M') | strcmp(genders_final, 'male');
        elseif isnumeric(genders_final)
            % 숫자인 경우 (1=남성, 0=여성으로 가정)
            gender_dummy = genders_final == 1;
        else
            fprintf('⚠ 성별 데이터 타입을 인식할 수 없습니다.\n');
            p_gender = NaN;
        end
        
        % 유효한 성별 데이터가 있는지 확인
        if exist('gender_dummy', 'var') && sum(gender_dummy) > 0 && sum(~gender_dummy) > 0
            % 성별별 성과 차이 t-test
            try
                [h, p_gender] = ttest2(y_final(gender_dummy==1), y_final(gender_dummy==0));
                fprintf('성별 효과 p-value: %.4f\n', p_gender);
                fprintf('  남성: %d명, 여성: %d명\n', sum(gender_dummy), sum(~gender_dummy));
            catch ME
                fprintf('⚠ 성별 t-test 수행 실패: %s\n', ME.message);
                p_gender = NaN;
            end
        else
            fprintf('⚠ 성별 그룹이 불균형하거나 데이터가 부족합니다.\n');
            p_gender = NaN;
        end
    catch ME
        fprintf('⚠ 성별 분석 중 에러 발생: %s\n', ME.message);
        p_gender = NaN;
    end
else
    fprintf('⚠ "성별" 컬럼을 찾을 수 없습니다.\n');
    p_gender = NaN;
end

% 유의미한 효과가 있을 경우 조정 모델 실행
if (exist('age_perf_corr', 'var') && ~isnan(age_perf_corr) && abs(age_perf_corr) > 0.3) || ...
   (exist('p_gender', 'var') && ~isnan(p_gender) && p_gender < 0.05)
    fprintf('\n▶ 인구통계 조정 모델 실행\n');
    
    try
        % 조정 변수 포함한 feature matrix 구성
        X_adjusted = X_final;
        adjusted_feature_names = feature_names;
        
        if exist('ages_valid', 'var') && ~isnan(age_perf_corr) && abs(age_perf_corr) > 0.3
            % 나이 변수 추가 (zscore 적용)
            ages_for_model = ages_valid;
            X_adjusted = [X_adjusted, zscore(ages_for_model)];
            adjusted_feature_names = [adjusted_feature_names, {'Age'}];
            fprintf('  - 나이 변수 추가됨\n');
        end
        
        if exist('gender_dummy', 'var') && ~isnan(p_gender) && p_gender < 0.05
            % 성별 변수 추가
            X_adjusted = [X_adjusted, double(gender_dummy)];
            adjusted_feature_names = [adjusted_feature_names, {'Gender'}];
            fprintf('  - 성별 변수 추가됨\n');
        end
        
        % 조정 모델 학습 (optimal_lambda가 정의되어 있다면)
        if exist('optimal_lambda', 'var')
            demo_adjusted_mdl = fitclinear(X_adjusted, y_final, ...
                'Learner', 'logistic', ...
                'Regularization', 'ridge', ...
                'Lambda', optimal_lambda);
        else
            demo_adjusted_mdl = fitclinear(X_adjusted, y_final, ...
                'Learner', 'logistic', ...
                'Regularization', 'ridge');
        end
        
        fprintf('  인구통계 변수 조정 완료\n');
    catch ME
        fprintf('  ⚠ 인구통계 조정 모델 학습 실패: %s\n', ME.message);
    end
else
    fprintf('▶ 인구통계 변수 효과가 미미하여 조정하지 않음\n');
end

fprintf('\n✅ 인구통계학적 변수 효과 통제 완료\n');

%% STEP 9.6: 성별 효과 비교 시각화
fprintf('\n【STEP 9.6】 성별 효과 비교 시각화\n');
fprintf('────────────────────────────────────────────\n');

% 성별 데이터 존재 여부 확인
% 성별 컬럼 찾기 (다양한 가능한 이름들 시도)
gender_col_candidates = {'성별', 'gender', 'Gender', 'GENDER', '성', 'sex', 'Sex', 'SEX'};
gender_col_idx = [];
gender_col_name = '';

for i = 1:length(gender_col_candidates)
    if exist('matched_hr', 'var') && any(strcmp(matched_hr.Properties.VariableNames, gender_col_candidates{i}))
        gender_col_idx = find(strcmp(matched_hr.Properties.VariableNames, gender_col_candidates{i}), 1);
        gender_col_name = gender_col_candidates{i};
        break;
    end
end

if ~isempty(gender_col_idx)
    fprintf('▶ 성별별 역량 프로파일 비교 분석 시작 (컬럼: %s)\n', gender_col_name);

    % 성별 데이터 추출 및 전처리
    gender_data = matched_hr{:, gender_col_idx};

    % 성별 데이터 정리 (다양한 형태 통일)
    if iscell(gender_data)
        gender_clean = cellfun(@(x) string(x), gender_data, 'UniformOutput', false);
        gender_clean = string(gender_clean);
    elseif isnumeric(gender_data)
        gender_clean = string(gender_data);
    else
        gender_clean = string(gender_data);
    end

    % 성별 표준화 (남성=1, 여성=0)
    male_indicators = contains(lower(gender_clean), {'남', 'male', 'm', '1'}) & ...
                     ~contains(lower(gender_clean), {'여', 'female', 'f'});
    female_indicators = contains(lower(gender_clean), {'여', 'female', 'f', '2'}) & ...
                       ~contains(lower(gender_clean), {'남', 'male', 'm'});

    % 유효한 성별 데이터가 있는 경우만 분석
    valid_gender_mask = male_indicators | female_indicators;

    if sum(valid_gender_mask) > 20 % 최소 샘플 수 확인
        gender_binary = nan(size(gender_clean));
        gender_binary(male_indicators) = 1;  % 남성
        gender_binary(female_indicators) = 0; % 여성

        % 성별별 샘플 수
        n_male = sum(gender_binary == 1, 'omitnan');
        n_female = sum(gender_binary == 0, 'omitnan');

        fprintf('  - 남성: %d명, 여성: %d명\n', n_male, n_female);

        if n_male >= 5 && n_female >= 5 % 최소 그룹 크기 확인

            % 1. 성별별 역량 프로파일 레이더 차트
            fprintf('▶ 성별별 역량 프로파일 레이더 차트 생성\n');

            % 주요 역량 선택 (상위 12개) - 기존에 정의된 top_comp_idx 사용
            if exist('top_comp_idx', 'var') && ~isempty(top_comp_idx)
                radar_comp_idx = top_comp_idx;
                radar_comp_names = top_comp_names;
            else
                % top_comp_idx가 없는 경우 처음 12개 사용
                n_top_comps = min(12, length(valid_comp_cols));
                radar_comp_idx = 1:n_top_comps;
                radar_comp_names = valid_comp_cols(1:n_top_comps);
            end
            n_radar_comps = length(radar_comp_idx);

            % 성별별 평균 계산
            male_mask = gender_binary == 1 & valid_gender_mask;
            female_mask = gender_binary == 0 & valid_gender_mask;

            male_profile = nanmean(matched_comp{male_mask, radar_comp_idx}, 1);
            female_profile = nanmean(matched_comp{female_mask, radar_comp_idx}, 1);

            % 레이더 차트 생성
            fig_radar = figure('Position', [150, 150, 800, 800], ...
                              'Name', '성별별 역량 프로파일 비교', 'Color', 'white');

            % 레이더 차트 함수 (인라인)
            angles = linspace(0, 2*pi, n_radar_comps+1);

            % 데이터 정규화 (0-1 범위)
            all_data = [male_profile; female_profile];
            data_min = min(all_data(:));
            data_max = max(all_data(:));
            data_range = data_max - data_min;

            if data_range > 0
                male_norm = (male_profile - data_min) / data_range;
                female_norm = (female_profile - data_min) / data_range;
            else
                male_norm = ones(size(male_profile)) * 0.5;
                female_norm = ones(size(female_profile)) * 0.5;
            end

            % 순환을 위해 첫 번째 값 추가
            male_plot = [male_norm, male_norm(1)];
            female_plot = [female_norm, female_norm(1)];

            % 좌표 변환
            [x_male, y_male] = pol2cart(angles, male_plot);
            [x_female, y_female] = pol2cart(angles, female_plot);

            % 그리드 그리기
            hold on;
            for r = 0.2:0.2:1
                [gx, gy] = pol2cart(angles, r*ones(size(angles)));
                plot(gx, gy, 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
            end

            % 축 그리기
            for i = 1:n_radar_comps
                plot([0, cos(angles(i))], [0, sin(angles(i))], 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
            end

            % 데이터 플롯
            plot(x_male, y_male, 'b-', 'LineWidth', 3, 'DisplayName', sprintf('남성 (n=%d)', n_male));
            plot(x_female, y_female, 'r-', 'LineWidth', 3, 'DisplayName', sprintf('여성 (n=%d)', n_female));

            % 마커 추가
            scatter(x_male(1:end-1), y_male(1:end-1), 60, 'b', 'filled');
            scatter(x_female(1:end-1), y_female(1:end-1), 60, 'r', 'filled');

            % 레이블 추가
            for i = 1:n_radar_comps
                label_r = 1.15;
                x_label = label_r * cos(angles(i));
                y_label = label_r * sin(angles(i));

                % 텍스트 정렬 조정
                if abs(x_label) < 0.1
                    ha = 'center';
                elseif x_label > 0
                    ha = 'left';
                else
                    ha = 'right';
                end

                if abs(y_label) < 0.1
                    va = 'middle';
                elseif y_label > 0
                    va = 'bottom';
                else
                    va = 'top';
                end

                text(x_label, y_label, radar_comp_names{i}, 'FontSize', 10, ...
                     'HorizontalAlignment', ha, 'VerticalAlignment', va, ...
                     'Interpreter', 'none');
            end

            axis equal;
            axis off;
            title('성별별 역량 프로파일 비교 (정규화)', 'FontSize', 14, 'FontWeight', 'bold');
            legend('Location', 'best');

            % 그래프 저장
            radar_file = get_managed_filepath(config.output_dir, 'gender_competency_radar.png', config);
            saveas(fig_radar, radar_file);
            fprintf('  ✓ 레이더 차트 저장: %s\n', radar_file);

            % 2. 역량별 성별 차이 막대그래프 (효과크기 포함)
            fprintf('▶ 역량별 성별 차이 분석 및 시각화\n');

            % 모든 역량에 대해 t-test 수행
            n_all_comps = length(valid_comp_cols);
            gender_diff_results = table();
            gender_diff_results.Competency = valid_comp_cols';
            gender_diff_results.Male_Mean = zeros(n_all_comps, 1);
            gender_diff_results.Female_Mean = zeros(n_all_comps, 1);
            gender_diff_results.Mean_Diff = zeros(n_all_comps, 1);
            gender_diff_results.Cohen_d = zeros(n_all_comps, 1);
            gender_diff_results.P_Value = ones(n_all_comps, 1);
            gender_diff_results.Significant = false(n_all_comps, 1);

            alpha = 0.05;
            bonferroni_alpha = alpha / n_all_comps; % Bonferroni 보정

            for i = 1:n_all_comps
                comp_data = matched_comp{:, i};

                male_scores = comp_data(male_mask);
                female_scores = comp_data(female_mask);

                % 결측치 제거
                male_scores = male_scores(~isnan(male_scores));
                female_scores = female_scores(~isnan(female_scores));

                if length(male_scores) >= 3 && length(female_scores) >= 3
                    % 기본 통계
                    gender_diff_results.Male_Mean(i) = mean(male_scores);
                    gender_diff_results.Female_Mean(i) = mean(female_scores);
                    gender_diff_results.Mean_Diff(i) = gender_diff_results.Male_Mean(i) - gender_diff_results.Female_Mean(i);

                    % t-test
                    try
                        [h, p] = ttest2(male_scores, female_scores);
                        gender_diff_results.P_Value(i) = p;
                        gender_diff_results.Significant(i) = p < bonferroni_alpha;

                        % Cohen's d 계산
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

            % 유의미한 차이가 있는 역량들
            sig_comps = gender_diff_results(gender_diff_results.Significant, :);
            fprintf('  - Bonferroni 보정 후 유의미한 성별 차이: %d개 역량\n', height(sig_comps));

            % 효과크기가 큰 역량들 (|Cohen's d| > 0.5)
            large_effect = abs(gender_diff_results.Cohen_d) > 0.5;
            fprintf('  - 큰 효과크기 (|d| > 0.5): %d개 역량\n', sum(large_effect));

            % 막대그래프 생성 (상위 20개 효과크기 기준)
            [~, sort_idx] = sort(abs(gender_diff_results.Cohen_d), 'descend');
            top_effects = sort_idx(1:min(20, length(sort_idx)));

            fig_bar = figure('Position', [200, 200, 1200, 600], ...
                            'Name', '성별별 역량 차이 (효과크기)', 'Color', 'white');

            effect_sizes = gender_diff_results.Cohen_d(top_effects);
            comp_names = gender_diff_results.Competency(top_effects);
            is_significant = gender_diff_results.Significant(top_effects);

            % 색상 설정 (유의미한 차이는 진한색, 아닌 것은 연한색)
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
            title('역량별 성별 차이 (상위 20개, *: p < 0.05 Bonferroni 보정)', 'FontSize', 12);
            grid on;

            % 범례 추가
            text(0.02, 0.98, sprintf('파랑: 남성 > 여성\n빨강: 여성 > 남성\n진한색: 통계적 유의\n연한색: 비유의'), ...
                 'Units', 'normalized', 'VerticalAlignment', 'top', ...
                 'BackgroundColor', 'white', 'EdgeColor', 'black', 'Interpreter', 'none');

            % 그래프 저장
            bar_file = get_managed_filepath(config.output_dir, 'gender_competency_differences.png', config);
            saveas(fig_bar, bar_file);
            fprintf('  ✓ 막대그래프 저장: %s\n', bar_file);

            % 3. 성별×성과 상호작용 분석
            fprintf('▶ 성별×성과 상호작용 분석\n');

            % 성과 그룹 정의 (matched_talent_types 사용)
            high_perf_mask = false(size(matched_talent_types));
            low_perf_mask = false(size(matched_talent_types));

            for i = 1:length(config.high_performers)
                high_perf_mask = high_perf_mask | strcmp(matched_talent_types, config.high_performers{i});
            end

            for i = 1:length(config.low_performers)
                low_perf_mask = low_perf_mask | strcmp(matched_talent_types, config.low_performers{i});
            end

            % 4개 그룹 정의
            group1_mask = male_mask & high_perf_mask;   % 고성과 남성
            group2_mask = female_mask & high_perf_mask; % 고성과 여성
            group3_mask = male_mask & low_perf_mask;    % 저성과 남성
            group4_mask = female_mask & low_perf_mask;  % 저성과 여성

            group_sizes = [sum(group1_mask), sum(group2_mask), sum(group3_mask), sum(group4_mask)];
            group_labels = {'고성과 남성', '고성과 여성', '저성과 남성', '저성과 여성'};

            fprintf('  - 그룹별 샘플 수: %s\n', mat2str(group_sizes));

            if all(group_sizes >= 3) % 모든 그룹에 최소 3명씩 있어야 분석 가능
                % 상호작용 시각화 (주요 역량 5개)
                fig_interact = figure('Position', [250, 250, 1000, 600], ...
                                     'Name', '성별×성과 상호작용', 'Color', 'white');

                n_interact_comps = min(5, n_all_comps);
                for comp_idx = 1:n_interact_comps
                    subplot(2, 3, comp_idx);

                    comp_data = matched_comp{:, comp_idx};

                    % 4개 그룹별 평균 계산
                    group_means = [
                        nanmean(comp_data(group1_mask));
                        nanmean(comp_data(group2_mask));
                        nanmean(comp_data(group3_mask));
                        nanmean(comp_data(group4_mask))
                    ];

                    % 2x2 패턴으로 재배열
                    interaction_data = reshape(group_means, 2, 2)'; % [고성과, 저성과] x [남성, 여성]

                    % 라인 플롯
                    plot([1, 2], interaction_data(1, :), 'b-o', 'LineWidth', 2, 'MarkerSize', 8, 'DisplayName', '고성과');
                    hold on;
                    plot([1, 2], interaction_data(2, :), 'r-s', 'LineWidth', 2, 'MarkerSize', 8, 'DisplayName', '저성과');

                    set(gca, 'XTick', [1, 2], 'XTickLabel', {'남성', '여성'});
                    ylabel('평균 점수');
                    title(valid_comp_cols{comp_idx}, 'Interpreter', 'none');
                    legend('Location', 'best');
                    grid on;
                end

                sgtitle('성별×성과 상호작용 분석 (주요 역량)', 'FontSize', 14, 'FontWeight', 'bold');

                % 상호작용 그래프 저장
                interact_file = get_managed_filepath(config.output_dir, 'gender_performance_interaction.png', config);
                saveas(fig_interact, interact_file);
                fprintf('  ✓ 상호작용 그래프 저장: %s\n', interact_file);

            else
                fprintf('  ⚠ 일부 그룹의 샘플 수가 부족하여 상호작용 분석을 건너뜁니다.\n');
            end

            % 4. 성별 편향 진단
            fprintf('▶ 성별 편향 진단\n');

            % Disparate Impact Ratio 계산 (고성과자 비율 기준)
            male_high_perf_rate = sum(male_mask & high_perf_mask) / sum(male_mask);
            female_high_perf_rate = sum(female_mask & high_perf_mask) / sum(female_mask);

            if female_high_perf_rate > 0
                disparate_impact_ratio = male_high_perf_rate / female_high_perf_rate;
            else
                disparate_impact_ratio = Inf;
            end

            fprintf('  - 남성 고성과자 비율: %.1f%%\n', male_high_perf_rate * 100);
            fprintf('  - 여성 고성과자 비율: %.1f%%\n', female_high_perf_rate * 100);
            fprintf('  - Disparate Impact Ratio: %.2f\n', disparate_impact_ratio);

            % 공정성 판단 (0.8-1.2 범위가 일반적 기준)
            if disparate_impact_ratio >= 0.8 && disparate_impact_ratio <= 1.2
                fairness_status = '공정';
            else
                fairness_status = '편향 가능성';
            end

            fprintf('  - 공정성 평가: %s\n', fairness_status);

            % 가장 큰 성별 차이를 보이는 역량 식별
            [max_effect, max_idx] = max(abs(gender_diff_results.Cohen_d));
            most_biased_comp = gender_diff_results.Competency{max_idx};

            fprintf('  - 최대 성별 차이 역량: %s (d = %.3f)\n', most_biased_comp, gender_diff_results.Cohen_d(max_idx));

            % 결과를 변수로 저장 (나중에 Excel 출력용)
            gender_analysis_results = struct();
            gender_analysis_results.sample_sizes = struct('male', n_male, 'female', n_female);
            gender_analysis_results.competency_differences = gender_diff_results;
            gender_analysis_results.fairness = struct('disparate_impact_ratio', disparate_impact_ratio, ...
                                                     'status', fairness_status);
            gender_analysis_results.group_analysis = struct('sizes', group_sizes, 'labels', {group_labels});

        else
            fprintf('  ⚠ 성별 그룹별 샘플 수가 부족합니다 (남성: %d명, 여성: %d명)\n', n_male, n_female);
        end

    else
        fprintf('  ⚠ 유효한 성별 데이터가 부족합니다 (%d명)\n', sum(valid_gender_mask));
    end

else
    fprintf('  ⚠ 성별 컬럼을 찾을 수 없어 성별 효과 분석을 건너뜁니다.\n');
end

fprintf('\n✅ 성별 효과 비교 시각화 완료\n');

%% 4.2 모든 역량 feature 전처리 (동질성 고려)
fprintf('\n【STEP 10】 역량 feature 전처리 (동질성 고려)\n');
fprintf('────────────────────────────────────────────\n');

% 동질성이 너무 높은 역량 제외 (CV < 0.1)
if exist('range_stats', 'var')
    very_high_homo = range_stats.CV < 0.1;
    fprintf('매우 높은 동질성 역량 제외: %d개\n', sum(very_high_homo));

    if any(very_high_homo)
        exclude_idx = very_high_homo;
        X_filtered = X_imputed(:, ~exclude_idx);
        feature_names_filtered = valid_comp_cols(~exclude_idx);

        fprintf('  최종 사용 역량: %d개 → %d개\n', length(valid_comp_cols), length(feature_names_filtered));

        X_imputed = X_filtered;
        feature_names = feature_names_filtered;
    else
        feature_names = valid_comp_cols;
    end
else
    % range_stats가 없는 경우 기본 처리
    feature_names = valid_comp_cols;
end

n_features = size(X_imputed, 2);
fprintf('  활용 역량 feature 수: %d개\n', n_features);

% 표준화는 LOO-CV 내부에서 수행 (데이터 누수 방지)
fprintf('  표준화는 교차검증 내부에서 수행됩니다 (데이터 누수 방지)\n');

%% 4.3 Leave-One-Out 교차검증으로 최적 Lambda 찾기
fprintf('\n【STEP 11】 Leave-One-Out 교차검증으로 최적 Lambda 찾기\n');
fprintf('────────────────────────────────────────────\n');

% Lambda 파라미터 범위 설정
lambda_range = logspace(-3, 0, 10);  % 0.001 ~ 1.0
fprintf('Lambda 후보값: [%.4f ~ %.4f], %d개 지점\n', min(lambda_range), max(lambda_range), length(lambda_range));

% Leave-One-Out 교차검증 수행
cv_scores = zeros(length(lambda_range), 1);
cv_aucs = zeros(length(lambda_range), 1);

fprintf('\nLOO-CV 진행상황:\n');
for lambda_idx = 1:length(lambda_range)
    current_lambda = lambda_range(lambda_idx);

    % LOO-CV를 위한 예측값 저장
    loo_predictions = zeros(length(y_weight), 1);
    loo_probabilities = zeros(length(y_weight), 1);

    % Leave-One-Out 루프
    for i = 1:length(y_weight)
        % 훈련 데이터 (i번째 샘플 제외)
        train_idx = true(length(y_weight), 1);
        train_idx(i) = false;

        X_train = X_imputed(train_idx, :);  % 원본 데이터 사용
        y_train = y_weight(train_idx);
        X_test = X_imputed(i, :);  % 원본 데이터 사용

        % ★ 훈련 세트로만 표준화 (데이터 누수 방지)
        mu = mean(X_train, 1);
        sigma = std(X_train, 0, 1);
        sigma(sigma == 0) = 1;  % 0 방지

        X_train_z = (X_train - mu) ./ sigma;
        X_test_z = (X_test - mu) ./ sigma;

        % 가중치 계산 (클래스 불균형 + 비용 반영)
        w = zeros(size(y_train));
        n_class0 = sum(y_train == 0);
        n_class1 = sum(y_train == 1);

        % 역빈도 가중치 * 비용 행렬
        w(y_train == 0) = (length(y_train)/(2*n_class0)) * cost_matrix(1,2);
        w(y_train == 1) = (length(y_train)/(2*n_class1)) * cost_matrix(2,1);

        try
            mdl = fitclinear(X_train_z, y_train, ...
                'Learner', 'logistic', ...
                'Regularization', 'ridge', ...
                'Lambda', current_lambda, ...
                'Solver', 'lbfgs', ...
                'Weights', w);  % Cost 대신 Weights 사용

            [pred_label, pred_score] = predict(mdl, X_test_z);
            loo_predictions(i) = pred_label;
            loo_probabilities(i) = pred_score(2);
        catch
            % 폴백
            loo_predictions(i) = mode(y_train);
            loo_probabilities(i) = mean(y_train);
        end
    end

    % 성능 평가
    accuracy = mean(loo_predictions == y_weight);

    % AUC 계산
    try
        [~, ~, ~, auc] = perfcurve(y_weight, loo_probabilities, 1);
        cv_aucs(lambda_idx) = auc;
    catch
        cv_aucs(lambda_idx) = 0.5;  % 기본값
    end

    cv_scores(lambda_idx) = accuracy;

    fprintf('  λ=%.4f: 정확도=%.3f, AUC=%.3f\n', current_lambda, accuracy, cv_aucs(lambda_idx));
end

% 최적 Lambda 선택 (AUC 기준)
[best_auc, best_idx] = max(cv_aucs);
optimal_lambda = lambda_range(best_idx);

fprintf('\n최적 Lambda 선택:\n');
fprintf('  최적 λ: %.4f\n', optimal_lambda);
fprintf('  최고 AUC: %.3f\n', best_auc);
fprintf('  해당 정확도: %.3f\n', cv_scores(best_idx));

%% 4.4 최적 Lambda로 최종 모델 학습 및 가중치 추출
fprintf('\n【STEP 12】 최종 Cost-Sensitive 모델 학습\n');
fprintf('────────────────────────────────────────────\n');

% 전체 데이터로 표준화 (최종 모델용)
mu_final = mean(X_imputed, 1);
sigma_final = std(X_imputed, 0, 1);
sigma_final(sigma_final == 0) = 1;
X_normalized = (X_imputed - mu_final) ./ sigma_final;

% 가중치 계산
sample_weights = zeros(size(y_weight));
n0 = sum(y_weight == 0);
n1 = sum(y_weight == 1);
sample_weights(y_weight == 0) = (length(y_weight)/(2*n0)) * cost_matrix(1,2);
sample_weights(y_weight == 1) = (length(y_weight)/(2*n1)) * cost_matrix(2,1);

try
    final_mdl = fitclinear(X_normalized, y_weight, ...
        'Learner', 'logistic', ...
        'Regularization', 'ridge', ...
        'Lambda', optimal_lambda, ...
        'Solver', 'lbfgs', ...
        'Weights', sample_weights);  % Cost 제거, Weights만 사용

    coefficients = final_mdl.Beta;
    intercept = final_mdl.Bias;

    fprintf('  ✓ Cost-Sensitive 로지스틱 회귀 학습 성공\n');
    fprintf('  절편: %.4f\n', intercept);

    % 양수 계수만 사용하여 가중치 변환
    positive_coefs = max(0, coefficients);
    final_weights = positive_coefs / sum(positive_coefs) * 100;  % 백분율로 변환

    fprintf('  양수 계수 개수: %d/%d\n', sum(positive_coefs > 0), length(coefficients));
    fprintf('  가중치 변환 완료 (백분율)\n');

catch ME
    warning('모델 학습 실패: %s. 상관계수로 대체합니다.', ME.message);
    correlations = zeros(n_features, 1);
    for i = 1:n_features
        correlations(i) = corr(X_normalized(:,i), y_weight, 'rows', 'complete');
    end
    coefficients = correlations;  % ★ 중요: coefficients 변수 설정
    intercept = 0;
end

%% 4.5 모델 성능 평가 및 검증
fprintf('\n【STEP 13】 모델 성능 평가 및 검증\n');
fprintf('────────────────────────────────────────────\n');

% 종합점수 계산 (가중치 적용)
weighted_scores = X_normalized * (final_weights / 100);  % 백분율을 다시 비율로 변환

% 고성과자와 저성과자의 종합점수 비교
high_idx = y_weight == 1;
low_idx = y_weight == 0;
high_scores = weighted_scores(high_idx);
low_scores = weighted_scores(low_idx);

fprintf('종합점수 검증:\n');
fprintf('  고성과자 평균: %.3f ± %.3f (n=%d)\n', mean(high_scores), std(high_scores), length(high_scores));
fprintf('  저성과자 평균: %.3f ± %.3f (n=%d)\n', mean(low_scores), std(low_scores), length(low_scores));
fprintf('  점수 차이: %.3f\n', mean(high_scores) - mean(low_scores));

% 통계적 유의성 검정
[~, ttest_p, ~, ttest_stats] = ttest2(high_scores, low_scores);
fprintf('  t-test: t=%.3f, p=%.6f\n', ttest_stats.tstat, ttest_p);

% Effect Size (Cohen's d) 계산
pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                  (length(low_scores)-1)*var(low_scores)) / ...
                  (length(high_scores) + length(low_scores) - 2));
cohens_d = (mean(high_scores) - mean(low_scores)) / pooled_std;

fprintf('  Cohen''s d: %.3f', cohens_d);
if cohens_d < 0.2
    fprintf(' (작은 효과)\n');
elseif cohens_d < 0.5
    fprintf(' (중간 효과)\n');
elseif cohens_d < 0.8
    fprintf(' (큰 효과)\n');
else
    fprintf(' (매우 큰 효과)\n');
end

% ROC 분석
[X_roc, Y_roc, T_roc, AUC] = perfcurve(y_weight, weighted_scores, 1);
fprintf('  분류 성능 (AUC): %.3f\n', AUC);

% 최적 임계값 찾기 (Youden's J statistic)
J = Y_roc - X_roc;
[~, opt_idx] = max(J);
optimal_threshold = T_roc(opt_idx);
fprintf('  최적 임계값: %.3f (민감도=%.3f, 특이도=%.3f)\n', ...
        optimal_threshold, Y_roc(opt_idx), 1-X_roc(opt_idx));

%% 4.6 가중치 결과 분석 및 저장
fprintf('\n【STEP 14】 가중치 결과 분석 및 저장\n');
fprintf('────────────────────────────────────────────\n');

% 가중치 결과 테이블 생성
weight_results = table();
weight_results.Feature = feature_names';
weight_results.Weight_Percent = final_weights;
weight_results.Raw_Coefficient = coefficients;

% 기여도가 있는 역량만 필터링 (0.1% 이상)
significant_idx = final_weights > 0.1;
weight_results_significant = weight_results(significant_idx, :);
weight_results_significant = sortrows(weight_results_significant, 'Weight_Percent', 'descend');

fprintf('주요 역량 가중치 (기여도 0.1%% 이상):\n');
fprintf('순위 | %-25s | 가중치(%%) | 원계수\n', '역량명');
fprintf('%s\n', repmat('-', 70, 1));

for i = 1:min(15, height(weight_results_significant))
    fprintf('%2d   | %-25s | %8.2f | %8.4f\n', ...
            i, weight_results_significant.Feature{i}, ...
            weight_results_significant.Weight_Percent(i), ...
            weight_results_significant.Raw_Coefficient(i));
end

% 가중치 파일 저장
result_data = struct();
result_data.final_weights = final_weights;
result_data.feature_names = feature_names;
result_data.optimal_lambda = optimal_lambda;
result_data.optimal_threshold = optimal_threshold;
result_data.model_performance = struct('AUC', AUC, 'cohens_d', cohens_d, ...
                                      'accuracy', cv_scores(best_idx));
result_data.cost_matrix = cost_matrix;
result_data.class_weights = class_weights;

% 가중치 파일 저장 (백업 시스템 적용)
weight_filepath = get_managed_filepath(config.output_dir, 'cost_sensitive_weights.mat', config);
backup_and_prepare_file(weight_filepath, config);
save(weight_filepath, 'result_data', 'weight_results_significant');

fprintf('\n가중치 저장 완료: %s\n', weight_filepath);

%% 4.7 종합 시각화
fprintf('\n【STEP 15】 Cost-Sensitive Learning 결과 시각화\n');
fprintf('────────────────────────────────────────────\n');

% 종합 시각화 생성
fig = figure('Position', [100, 100, 1400, 1000], 'Color', 'white');

% 1. Lambda 최적화 과정
subplot(2, 3, 1);
yyaxis left
plot(lambda_range, cv_scores, 'o-', 'LineWidth', 2, 'Color', [0.2, 0.4, 0.8]);
ylabel('정확도', 'Color', [0.2, 0.4, 0.8]);
ylim([min(cv_scores)-0.02, max(cv_scores)+0.02]);

yyaxis right
plot(lambda_range, cv_aucs, 's-', 'LineWidth', 2, 'Color', [0.8, 0.3, 0.2]);
ylabel('AUC', 'Color', [0.8, 0.3, 0.2]);
ylim([min(cv_aucs)-0.02, max(cv_aucs)+0.02]);

% 최적점 표시
hold on;
yyaxis right;
plot(optimal_lambda, best_auc, 'r*', 'MarkerSize', 12, 'LineWidth', 3);

set(gca, 'XScale', 'log');
xlabel('Lambda (정규화 강도)');
title('LOO-CV Lambda 최적화');
grid on;

% 2. 주요 가중치 (상위 12개)
subplot(2, 3, 2);
top_n = min(12, height(weight_results_significant));
top_weights = weight_results_significant(1:top_n, :);

barh(1:top_n, top_weights.Weight_Percent, 'FaceColor', [0.3, 0.7, 0.4]);
set(gca, 'YTick', 1:top_n, 'YTickLabel', top_weights.Feature, 'FontSize', 9);
xlabel('가중치 (%)');
title('주요 역량 가중치 (상위 12개)');
grid on;

% 3. 종합점수 분포 비교
subplot(2, 3, 3);
bin_edges = linspace(min(weighted_scores), max(weighted_scores), 15);
histogram(low_scores, bin_edges, 'FaceColor', [0.8, 0.3, 0.3], 'FaceAlpha', 0.7, ...
          'DisplayName', sprintf('저성과자 (n=%d)', length(low_scores)));
hold on;
histogram(high_scores, bin_edges, 'FaceColor', [0.3, 0.7, 0.3], 'FaceAlpha', 0.7, ...
          'DisplayName', sprintf('고성과자 (n=%d)', length(high_scores)));

% 최적 임계값 표시
line([optimal_threshold, optimal_threshold], ylim, 'Color', 'k', 'LineStyle', '--', ...
     'LineWidth', 2, 'DisplayName', sprintf('최적 임계값 (%.3f)', optimal_threshold));

xlabel('종합점수');
ylabel('빈도');
title('고성과자 vs 저성과자 점수 분포');
legend('Location', 'best');
grid on;

% 4. ROC 곡선
subplot(2, 3, 4);
plot(X_roc, Y_roc, 'LineWidth', 3, 'Color', [0.2, 0.4, 0.8]);
hold on;
plot([0, 1], [0, 1], '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);

% 최적점 표시
plot(X_roc(opt_idx), Y_roc(opt_idx), 'ro', 'MarkerSize', 10, 'LineWidth', 2);

xlabel('위양성률 (1-특이도)');
ylabel('민감도');
title(sprintf('ROC 곡선 (AUC=%.3f)', AUC));
legend({'ROC 곡선', '무작위', '최적점'}, 'Location', 'southeast');
grid on;

% 5. 클래스별 가중치 기여도
subplot(2, 3, 5);
positive_weights = final_weights(final_weights > 0);
pie_data = [sum(positive_weights), 100 - sum(positive_weights)];
pie_labels = {sprintf('활성 역량\n(%.1f%%)', pie_data(1)), ...
              sprintf('비활성 역량\n(%.1f%%)', pie_data(2))};

pie(pie_data, pie_labels);
% title('역량 활용도');
colormap([0.3, 0.7, 0.4; 0.8, 0.8, 0.8]);

% 6. 성능 지표 요약
subplot(2, 3, 6);
axis off;

% 성능 지표 텍스트
perf_text = {
    sprintf('◆ Cost-Sensitive Learning 결과 ◆');
    '';
    sprintf('최적 Lambda: %.4f', optimal_lambda);
    sprintf('교차검증 AUC: %.3f', best_auc);
    sprintf('교차검증 정확도: %.3f', cv_scores(best_idx));
    '';
    sprintf('Cohen''s d: %.3f', cohens_d);
    sprintf('최적 임계값: %.3f', optimal_threshold);
    sprintf('활성 역량 수: %d/%d', sum(final_weights > 0.1), length(final_weights));
    '';
    sprintf('클래스 가중치:');
    sprintf('  저성과자: %.3f', class_weights(1));
    sprintf('  고성과자: %.3f', class_weights(2));
    '';
    sprintf('비용 행렬: [0, 1; 1.5, 0]');
};

text(0.05, 0.95, perf_text, 'Units', 'normalized', 'VerticalAlignment', 'top', ...
     'FontSize', 10, 'FontWeight', 'bold');

sgtitle('Cost-Sensitive Learning 기반 고성과자 예측 시스템 분석 결과', ...
        'FontSize', 14, 'FontWeight', 'bold');

% 그래프 저장 (백업 시스템 적용)
chart_filepath = get_managed_filepath(config.output_dir, 'cost_sensitive_analysis.png', config);
backup_and_prepare_file(chart_filepath, config);
saveas(fig, chart_filepath);

fprintf('  ✓ 시각화 차트 저장: %s\n', chart_filepath);
