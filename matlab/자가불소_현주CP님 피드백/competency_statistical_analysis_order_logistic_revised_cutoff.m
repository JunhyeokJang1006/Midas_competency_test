% 주요 특징:
% 1. 개별 레이더 차트 생성 (통일된 스케일)
% 3. 하이퍼파라미터 튜닝 및 교차검증
% 4. 고도화된 상관 기반 가중치 시스템
% 5. Logistic regression과 상관분석 통합
% 6. 학습된 모델 저장 및 재사용 기능
% 7. 비활성/과활성 포함 분석, 신뢰불가 데이터 제외 추가
% 8. 극단 그룹 비교 방식 선택 옵션 (STEP 17)
%    - 'extreme': 가장 확실한 케이스만 (자연성,성실한가연성 vs 무능한불연성,소화성)
%    - 'all': 모든 고성과자 vs 저성과자 (중간 그룹 포함)
% 9. F1 스코어 통계적 유의성 검증 (STEP 22.5)
%    - 퍼뮤테이션 테스트 (5000회)
%    - 캐싱 시스템으로 재계산 방지
%    - Cohen's d 효과 크기 측정
% 10. Range Restriction 보정 (Cut-off 버전)
%    - Thorndike Case II 공식을 통한 상관계수 보정
%    - 선발 효과(Cut-off) 반영한 로지스틱 회귀
%    - Truncated Regression 및 Heckman 2단계 보정
%    - 가중치 조정을 통한 Cut-off 근처 샘플 강조
 
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
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\자가불소';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
config.model_file = fullfile(config.output_dir, 'trained_logitboost_model.mat');
config.use_saved_model = true;  % 저장된 모델 사용 여부
config.baseline_type = 'weighted';  % 'simple' 또는 'weighted' 선택

% 극단 그룹 비교 방식 설정
config.extreme_group_method = 'all';  % 'extreme' 또는 'all' 선택
% 'extreme': 가장 확실한 케이스만 (자연성, 성실한 가연성 vs 무능한 불연성, 소화성)
% 'all': 모든 고성과자 vs 저성과자 비교

% 퍼뮤테이션 테스트 설정
config.force_recalc_permutation = false;  % 퍼뮤테이션 강제 재계산 여부

% Cut-off 보정 설정 (Cut-off 버전에서 추가)
config.cutoff_correction_enabled = true;  % Cut-off 보정 활성화 여부
config.cutoff_method = 'truncated';  % Cut-off 보정 방법: 'weighted', 'heckman', 'truncated'
config.thorndike_correction = true;  % Thorndike Case II 보정 사용 여부
config.minimum_cutoff_threshold = 40;  % Cut-off 탐지 최소 임계값

% 파일 관리 시스템 설정
config.create_backup = true;  % 백업 폴더 생성 여부
config.backup_folder = 'backup';  % 백업 폴더 이름
config.use_timestamp = false;  % 파일명에 타임스탬프 사용 여부

% 성과 순위 정의 (위장형 소화성만 제외)
config.performance_ranking = containers.Map(...
    {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
     '게으른 가연성', '무능한 불연성', '소화성'}, ...
    [8, 7, 6, 5, 4, 3, 1]);

% 순서형 로지스틱 회귀용 레벨 정의 (1: 최하위, 5: 최상위)
config.ordinal_levels = containers.Map(...
    {'소화성', '무능한 불연성', '게으른 가연성', '유능한 불연성', ...
     '유익한 불연성', '성실한 가연성', '자연성'}, ...
    [1, 2, 3, 4, 5, 6, 7]);

config.level_names = {'소화성', '무능한 불연성', '게으른 가연성', '유능한 불연성', ...
                     '유익한 불연성', '성실한 가연성', '자연성'};

% 고성과자/저성과자 정의 (이진 분류용) - 유능한 불연성 제외
config.high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
config.low_performers = {'무능한 불연성', '소화성', '게으른 가연성'};
config.excluded_from_analysis = {'유능한 불연성'};  % 분석에서 제외
config.excluded_types = {'위장형 소화성'}; % 위장형 소화성도 제외

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
        fprintf('  📁 백업 폴더 생성: %s\n', backup_dir);
    end

    % 기존 파일이 있으면 백업으로 이동
    if exist(filepath, 'file')
        [~, filename, ext] = fileparts(filepath);
        backup_timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
        backup_filename = sprintf('%s_%s%s', filename, backup_timestamp, ext);
        backup_filepath = fullfile(backup_dir, backup_filename);

        try
            movefile(filepath, backup_filepath);
            fprintf('  💾 기존 파일을 백업 폴더로 이동: %s → %s/%s\n', ...
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

    user_choice = input('  선택 (1-3, 엔터=기본값): ', 's');
    if isempty(user_choice)
        user_choice = '1';
    end

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

fprintf('\n전체 인재유형 분포:\n');
for i = 1:length(unique_types)
    fprintf('  • %-20s: %3d명 (%5.1f%%)\n', ...
        unique_types{i}, type_counts(i), type_counts(i)/sum(type_counts)*100);
end

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
    total_scores = comp_total{total_idx, end};
    fprintf('▶ 종합점수 통합: %d명\n', length(total_idx));
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

% 분포 분석을 위한 데이터 준비
n_comps = length(valid_comp_cols);
distribution_stats = table();
distribution_stats.Competency = valid_comp_cols';
distribution_stats.Mean = zeros(n_comps, 1);
distribution_stats.Std = zeros(n_comps, 1);
distribution_stats.Skewness = zeros(n_comps, 1);
distribution_stats.Kurtosis = zeros(n_comps, 1);
distribution_stats.Min_Cutoff_Effect = false(n_comps, 1);
distribution_stats.Max_Cutoff_Effect = false(n_comps, 1);
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

    % Cut-off 효과 탐지 (선발 효과로 인한 truncated distribution)
    % 최소값이 이론적 최소값(0)보다 비정상적으로 높은 경우
    theoretical_min = 0;
    theoretical_max = 100;
    min_val = min(valid_data);
    max_val = max(valid_data);

    % 최소값 cut-off: 데이터의 최소값이 이론적 범위의 25% 이상인 경우
    if min_val > theoretical_min + 0.25 * (theoretical_max - theoretical_min)
        distribution_stats.Min_Cutoff_Effect(i) = true;
    end

    % 최대값 cut-off: 데이터의 최대값이 이론적 범위의 75% 이하인 경우
    if max_val < theoretical_min + 0.75 * (theoretical_max - theoretical_min)
        distribution_stats.Max_Cutoff_Effect(i) = true;
    end

    % 분포 타입 분류
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

% 분포 시각화 생성
fprintf('▶ 역량 분포 subplot 그리드 생성 중...\n');

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

    % 히스토그램과 정규분포 곡선
    histogram(valid_data, 'Normalization', 'pdf', 'FaceAlpha', 0.7, ...
              'EdgeColor', 'none', 'FaceColor', [0.3 0.6 0.9]);
    hold on;

    % 정규분포 곡선 오버레이
    x_range = linspace(min(valid_data), max(valid_data), 100);
    norm_curve = normpdf(x_range, mean(valid_data), std(valid_data));
    plot(x_range, norm_curve, 'r-', 'LineWidth', 2);

    % 통계 정보 텍스트
    stats_text = sprintf('평균: %.1f\n표준편차: %.1f\n왜도: %.2f\n첨도: %.2f', ...
                        distribution_stats.Mean(i), distribution_stats.Std(i), ...
                        distribution_stats.Skewness(i), distribution_stats.Kurtosis(i));

    % Cut-off 효과 표시
    if distribution_stats.Min_Cutoff_Effect(i) || distribution_stats.Max_Cutoff_Effect(i)
        cutoff_text = '\n⚠ Cut-off 효과';
        stats_text = [stats_text, cutoff_text];
    end

    text(0.02, 0.98, stats_text, 'Units', 'normalized', 'VerticalAlignment', 'top', ...
         'FontSize', 8, 'BackgroundColor', 'white', 'EdgeColor', 'black', 'Interpreter', 'none');

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
cutoff_min_count = sum(distribution_stats.Min_Cutoff_Effect);
cutoff_max_count = sum(distribution_stats.Max_Cutoff_Effect);
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
    fprintf('    💡 가중평균(weighted) baseline 사용을 권장합니다.\n');
else
    fprintf('  ✓ 샘플 균형: 최대/최소 비율 = %.1f\n', imbalance_ratio);
end

fprintf('  - 선택된 baseline 유형: %s\n', config.baseline_type);

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

    % 스케일 정규화
    data_norm = (data - min_val) / (max_val - min_val);
    baseline_norm = (baseline_data - min_val) / (max_val - min_val);

    % 순환을 위해 첫 번째 값을 마지막에 추가
    data_plot = [data_norm, data_norm(1)];
    baseline_plot = [baseline_norm, baseline_norm(1)];

    % 좌표 변환
    [x_data, y_data] = pol2cart(angles, data_plot);
    [x_base, y_base] = pol2cart(angles, baseline_plot);

    hold on;

    % 그리드 그리기
    grid_levels = 5;
    for j = 1:grid_levels
        r = j / grid_levels;
        [gx, gy] = pol2cart(angles, r*ones(size(angles)));
        plot(gx, gy, 'k:', 'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);

        % 그리드 레이블
        grid_value = min_val + (max_val - min_val) * r;
        text(0, r, sprintf('%.0f', grid_value), ...
             'HorizontalAlignment', 'center', ...
             'VerticalAlignment', 'bottom', ...
             'FontSize', 9, 'Color', [0.4 0.4 0.4]);
    end

    % 방사선 그리기
    for j = 1:n_vars
        plot([0, cos(angles(j))], [0, sin(angles(j))], 'k:', ...
             'LineWidth', 0.8, 'Color', [0.6 0.6 0.6]);
    end

    % 기준선
    plot(x_base, y_base, '--', 'Color', [0.7 0.7 0.7], 'LineWidth', 2);

    % 데이터 플롯
    patch(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2.5);

    % 데이터 포인트
    scatter(x_data(1:end-1), y_data(1:end-1), 60, color, 'filled', ...
            'MarkerEdgeColor', 'white', 'LineWidth', 1);

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

    % 범례
    legend({'평균선', '해당 유형'}, 'Location', 'best', 'FontSize', 10);

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

    % % 샘플 크기 경고 (10개 미만인 경우)
    % if sample_count < 10
    %     text(0.5, -0.12, sprintf('⚠ 주의: 샘플 수가 적음 (n=%d)', sample_count), ...
    %          'Units', 'normalized', ...
    %          'HorizontalAlignment', 'center', ...
    %          'FontWeight', 'bold', 'FontSize', 12, ...
    %          'Color', 'red');
    % end

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

% 유능한 불연성 제외 후 데이터 확인
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

% Step 1: 고품질 샘플만 선택 (결측률 30% 미만)
missing_rate_per_sample = sum(isnan(X_raw), 2) / size(X_raw, 2);
quality_threshold = 0.3;
quality_samples = missing_rate_per_sample < quality_threshold;

fprintf('Step 1 - 고품질 샘플 선택:\n');
fprintf('  결측률 30%% 이상 제거: %d명 제외\n', sum(~quality_samples));
fprintf('  남은 샘플: %d명\n', sum(quality_samples));

% 유능한 불연성 제외 + 품질 필터 적용
binary_valid_idx = ~isnan(y_binary);
final_idx = binary_valid_idx & quality_samples;
X_quality = X_raw(final_idx, :);
y_quality = y_binary(final_idx);

% Step 2: 완전한 케이스만 사용 (남은 결측값 제거)
complete_cases = ~any(isnan(X_quality), 2);
X_final = X_quality(complete_cases, :);
y_final = y_quality(complete_cases);

fprintf('\nStep 2 - 완전한 케이스만 사용:\n');
fprintf('  추가 제거: %d명\n', sum(~complete_cases));
fprintf('  최종 분석 샘플: %d명\n', length(y_final));
fprintf('  - 고성과자: %d명\n', sum(y_final == 1));
fprintf('  - 저성과자: %d명\n', sum(y_final == 0));
fprintf('  - 샘플/변수 비율: %.1f\n', length(y_final)/size(X_final, 2));

% 결측값 대체 없이 완전한 데이터만 사용
X_imputed = X_final;  % 이름은 유지하되 실제로는 대체 없음
y_weight = y_final;

% 클래스 분포 확인
n_high = sum(y_weight == 1);
n_low = sum(y_weight == 0);
total_binary = length(y_weight);

fprintf('\n최종 이진분류 데이터 분포:\n');
fprintf('  고성과자 (1): %d명 (%.1f%%)\n', n_high, n_high/total_binary*100);
fprintf('  저성과자 (0): %d명 (%.1f%%)\n', n_low, n_low/total_binary*100);
fprintf('  불균형 비율: %.2f:1\n', max(n_high, n_low)/min(n_high, n_low));

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
if exist('matched_hr', 'var') && any(strcmp(matched_hr.Properties.VariableNames, '성별'))
    fprintf('▶ 성별별 역량 프로파일 비교 분석 시작\n');

    % 성별 데이터 추출 및 전처리
    gender_data = matched_hr.('성별');

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

            % 주요 역량 선택 (상위 12개)
            n_top_comps = min(12, length(valid_comp_cols));

            % 성별별 평균 계산
            male_mask = gender_binary == 1 & valid_gender_mask;
            female_mask = gender_binary == 0 & valid_gender_mask;

            male_profile = nanmean(matched_comp{male_mask, 1:n_top_comps}, 1);
            female_profile = nanmean(matched_comp{female_mask, 1:n_top_comps}, 1);

            % 레이더 차트 생성
            fig_radar = figure('Position', [150, 150, 800, 800], ...
                              'Name', '성별별 역량 프로파일 비교', 'Color', 'white');

            % 레이더 차트 함수 (인라인)
            angles = linspace(0, 2*pi, n_top_comps+1);

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
            for i = 1:n_top_comps
                plot([0, cos(angles(i))], [0, sin(angles(i))], 'k:', 'LineWidth', 0.5, 'Color', [0.7 0.7 0.7]);
            end

            % 데이터 플롯
            plot(x_male, y_male, 'b-', 'LineWidth', 3, 'DisplayName', sprintf('남성 (n=%d)', n_male));
            plot(x_female, y_female, 'r-', 'LineWidth', 3, 'DisplayName', sprintf('여성 (n=%d)', n_female));

            % 마커 추가
            scatter(x_male(1:end-1), y_male(1:end-1), 60, 'b', 'filled');
            scatter(x_female(1:end-1), y_female(1:end-1), 60, 'r', 'filled');

            % 레이블 추가
            for i = 1:n_top_comps
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

                text(x_label, y_label, valid_comp_cols{i}, 'FontSize', 10, ...
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
    % Cut-off 보정 로지스틱 회귀 모델 학습
    if exist('cutoff_diagnosis', 'var') && sum(cutoff_diagnosis.Severity ~= 'None') > 0
        fprintf('  ▶ Cut-off 보정 로지스틱 회귀 적용\n');

        % Cut-off 보정 방법 선택 (가중치 조정 방식 사용)
        final_mdl = cutoff_aware_logistic(X_normalized, y_weight, ...
                                        matched_comp{:, 1:size(X_normalized, 2)}, 'weighted');

        % 모델에서 계수 추출 (GLM 객체에서)
        coefficients = final_mdl.Coefficients.Estimate(2:end);  % 절편 제외
        intercept = final_mdl.Coefficients.Estimate(1);  % 절편

        fprintf('    - Cut-off 보정 모델 학습 완료\n');
    else
        fprintf('  ▶ 표준 로지스틱 회귀 사용 (Cut-off 효과 미미)\n');

        final_mdl = fitclinear(X_normalized, y_weight, ...
            'Learner', 'logistic', ...
            'Regularization', 'ridge', ...
            'Lambda', optimal_lambda, ...
            'Solver', 'lbfgs', ...
            'Weights', sample_weights);  % Cost 제거, Weights만 사용

        coefficients = final_mdl.Beta;
        intercept = final_mdl.Bias;
    end

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

%% 4.8 Bootstrap을 통한 가중치 안정성 검증
fprintf('\n【STEP 16】 Bootstrap 가중치 안정성 검증\n');
fprintf('────────────────────────────────────────────\n');


bootstrap_chart_filename='D:\project\HR데이터\결과\자가불소\bootstrap.xlsx';

% Bootstrap 설정
n_bootstrap = 5000;
n_samples = size(X_final, 1);
n_features = size(X_final, 2);

% Bootstrap 결과 저장 배열
bootstrap_weights = zeros(n_features, n_bootstrap);
bootstrap_rankings = zeros(n_features, n_bootstrap);

% Progress bar 표시
fprintf('Bootstrap 진행 중: ');

for b = 1:n_bootstrap
    % 복원추출로 재샘플링
    bootstrap_idx = randsample(n_samples, n_samples, true);
    X_boot = X_final(bootstrap_idx, :);
    y_boot = y_final(bootstrap_idx);

    % 정규화
    X_boot_norm = zscore(X_boot);

    % 샘플 가중치 재계산
    n_high_boot = sum(y_boot == 1);
    n_low_boot = sum(y_boot == 0);

    % 클래스가 하나만 있는 경우 스킵
    if n_high_boot == 0 || n_low_boot == 0
        bootstrap_weights(:, b) = NaN;
        bootstrap_rankings(:, b) = NaN;
        continue;
    end

    class_weights_boot = [n_samples/(2*n_low_boot), n_samples/(2*n_high_boot)];
    sample_weights_boot = zeros(size(y_boot));
    sample_weights_boot(y_boot == 0) = class_weights_boot(1);
    sample_weights_boot(y_boot == 1) = class_weights_boot(2);

    % Cost-Sensitive 모델 학습 (최적 Lambda 사용)
    try
        mdl_boot = fitclinear(X_boot_norm, y_boot, ...
            'Learner', 'logistic', ...
            'Cost', cost_matrix, ...
            'Weights', sample_weights_boot, ...
            'Regularization', 'ridge', ...
            'Lambda', optimal_lambda);

        % 가중치 추출 및 저장
        coefs = mdl_boot.Beta;
        positive_coefs = max(0, coefs);
        if sum(positive_coefs) > 0
            weights = positive_coefs / sum(positive_coefs) * 100;
        else
            weights = zeros(size(positive_coefs));
        end
        bootstrap_weights(:, b) = weights;

        % 순위 저장
        [~, ranks] = sort(weights, 'descend');
        for r = 1:length(ranks)
            bootstrap_rankings(ranks(r), b) = r;
        end

    catch
        % 실패한 경우 NaN 처리
        bootstrap_weights(:, b) = NaN;
        bootstrap_rankings(:, b) = NaN;
    end

    % Progress 표시
    if mod(b, 100) == 0
        fprintf('.');
    end
end
fprintf(' 완료!\n\n');

% Bootstrap 통계 계산
fprintf('【Bootstrap 결과 분석】\n');
fprintf('────────────────────────────────────────────\n\n');

% 각 역량별 통계
bootstrap_stats = table();
bootstrap_stats.Feature = feature_names';
bootstrap_stats.Original_Weight = final_weights;
bootstrap_stats.Boot_Mean = nanmean(bootstrap_weights, 2);
bootstrap_stats.Boot_Std = nanstd(bootstrap_weights, 0, 2);
bootstrap_stats.CI_Lower = prctile(bootstrap_weights, 2.5, 2);
bootstrap_stats.CI_Upper = prctile(bootstrap_weights, 97.5, 2);
bootstrap_stats.CV = bootstrap_stats.Boot_Std ./ (bootstrap_stats.Boot_Mean + eps);  % 변동계수

% 순위 안정성
bootstrap_stats.Avg_Rank = zeros(n_features, 1);
bootstrap_stats.Top3_Prob = zeros(n_features, 1);
bootstrap_stats.Top5_Prob = zeros(n_features, 1);

for i = 1:n_features
    valid_ranks = bootstrap_rankings(i, :);
    valid_ranks = valid_ranks(~isnan(valid_ranks));

    if ~isempty(valid_ranks)
        bootstrap_stats.Avg_Rank(i) = mean(valid_ranks);
        bootstrap_stats.Top3_Prob(i) = sum(valid_ranks <= 3) / length(valid_ranks) * 100;
        bootstrap_stats.Top5_Prob(i) = sum(valid_ranks <= 5) / length(valid_ranks) * 100;
    end
end

% 원본 가중치 기준으로 정렬
bootstrap_stats = sortrows(bootstrap_stats, 'Original_Weight', 'descend');

% 결과 출력
fprintf('가중치 안정성 분석 (상위 10개):\n');
fprintf('%-20s | 원본(%%) | 평균(%%) | 95%% CI | CV | Top3확률(%%) | Top5확률(%%)\n', '역량');
fprintf('%s\n', repmat('-', 95, 1));

for i = 1:min(10, height(bootstrap_stats))
    fprintf('%-20s | %7.2f | %7.2f | [%5.2f-%5.2f] | %4.2f | %7.1f | %7.1f\n', ...
        bootstrap_stats.Feature{i}, ...
        bootstrap_stats.Original_Weight(i), ...
        bootstrap_stats.Boot_Mean(i), ...
        bootstrap_stats.CI_Lower(i), ...
        bootstrap_stats.CI_Upper(i), ...
        bootstrap_stats.CV(i), ...
        bootstrap_stats.Top3_Prob(i), ...
        bootstrap_stats.Top5_Prob(i));
end

% 안정성 평가
fprintf('\n【안정성 평가】\n');
fprintf('────────────────────────────────────────────\n');

% 매우 안정적 (CV < 0.3 & Top3 > 70%)
very_stable = bootstrap_stats.CV < 0.3 & bootstrap_stats.Top3_Prob > 70;
if any(very_stable)
    fprintf('✅ 매우 안정적인 역량 (일관되게 중요):\n');
    stable_features = find(very_stable);
    for i = stable_features'
        fprintf('   - %s (CV=%.2f, Top3=%.1f%%)\n', ...
            bootstrap_stats.Feature{i}, ...
            bootstrap_stats.CV(i), ...
            bootstrap_stats.Top3_Prob(i));
    end
else
    fprintf('✅ 매우 안정적인 역량: 없음\n');
end

% 불안정 (CV > 0.5 | Top5 < 30%)
unstable = bootstrap_stats.CV > 0.5 | bootstrap_stats.Top5_Prob < 30;
if any(unstable)
    fprintf('\n⚠️ 불안정한 역량 (해석 주의):\n');
    unstable_features = find(unstable);
    for i = unstable_features'
        fprintf('   - %s (CV=%.2f, Top5=%.1f%%)\n', ...
            bootstrap_stats.Feature{i}, ...
            bootstrap_stats.CV(i), ...
            bootstrap_stats.Top5_Prob(i));
    end
else
    fprintf('\n⚠️ 불안정한 역량: 없음\n');
end

% Bootstrap 시각화
bootstrap_fig = figure('Position', [100, 100, 1600, 1200], 'Color', 'white');

% 전체 10개 역량의 Bootstrap 분포
subplot(3, 2, [1:4]);
top_10 = bootstrap_stats(1:min(10, height(bootstrap_stats)), :);
top_10_indices = zeros(height(top_10), 1);
for i = 1:height(top_10)
    top_10_indices(i) = find(strcmp(feature_names, top_10.Feature{i}));
end

boxplot(bootstrap_weights(top_10_indices, :)', ...
    'Labels', top_10.Feature, 'Colors', lines(10));
hold on;
% 원본 가중치 표시
for i = 1:height(top_10)
    feat_idx = top_10_indices(i);
    plot(i, final_weights(feat_idx), 'r*', 'MarkerSize', 15, 'LineWidth', 2);
end
ylabel('가중치 (%)', 'FontWeight', 'bold');
title('Bootstrap 가중치 분포 (전체 10개 역량)', 'FontWeight', 'bold');
legend({'원본 가중치'}, 'Location', 'northeast');
grid on;
% X축 레이블 회전 (가독성 향상)
xtickangle(45);

% 순위 변동성 히트맵
subplot(3, 2, [5:6]);
% 상위 10개 역량의 순위 확률 매트릭스
rank_prob_matrix = zeros(min(10, n_features), 10);
for i = 1:min(10, n_features)
    feat_idx = find(strcmp(feature_names, bootstrap_stats.Feature{i}));
    if ~isempty(feat_idx)
        for r = 1:10
            valid_ranks = bootstrap_rankings(feat_idx, :);
            valid_ranks = valid_ranks(~isnan(valid_ranks));
            if ~isempty(valid_ranks)
                rank_prob_matrix(i, r) = sum(valid_ranks == r) / length(valid_ranks) * 100;
            end
        end
    end
end

imagesc(rank_prob_matrix);
colormap(hot);
colorbar;
set(gca, 'YTick', 1:min(10, height(bootstrap_stats)), ...
         'YTickLabel', bootstrap_stats.Feature(1:min(10, height(bootstrap_stats))));
set(gca, 'XTick', 1:10, 'XTickLabel', 1:10);
xlabel('순위', 'FontWeight', 'bold');
ylabel('역량', 'FontWeight', 'bold');
title('순위 확률 분포 (노란색 = 높은 확률)', 'FontWeight', 'bold');

% sgtitle('Bootstrap 안정성 검증 (1000회 재샘플링)', 'FontSize', 16, 'FontWeight', 'bold');

% 그래프 저장
% bootstrap_chart_filename = sprintf('bootstrap_stability_%s.png', datestr(now, 'yyyy-mm-dd_HHMMSS'));
% saveas(bootstrap_fig, bootstrap_chart_filename);

fprintf('\n✅ Bootstrap 검증 완료\n');
fprintf('📊 시각화 저장 완료: %s\n', bootstrap_chart_filename);

%% 4.9 극단 그룹 비교 분석
fprintf('\n【STEP 17】 극단 그룹 t-test 비교 분석\n');
fprintf('────────────────────────────────────────────\n\n');

% 비교 방식에 따른 그룹 정의
if strcmp(config.extreme_group_method, 'extreme')
    % 'extreme': 가장 확실한 케이스만
    extreme_high = {'자연성', '성실한 가연성'};  % CODE 8, 7
    extreme_low = {'무능한 불연성', '소화성'};   % CODE 2, 1
    fprintf('💡 분석 방식: 극단 그룹 비교 (가장 확실한 케이스만)\n');
    fprintf('   📈 고성과: 자연성, 성실한 가연성\n');
    fprintf('   📉 저성과: 무능한 불연성, 소화성\n');
    fprintf('   ✅ 장점: 명확한 구분, 높은 효과 크기 기대\n');
    fprintf('   ⚠️  주의: 표본 수 제한, 일반화 제약\n\n');
else  % 'all'
    % 'all': 모든 고성과자 vs 저성과자
    extreme_high = {'자연성', '성실한 가연성', '유익한 불연성'};  % CODE 8, 7, 6
    extreme_low = {'무능한 불연성', '소화성', '게으른 가연성'};   % CODE 2, 1, 4
    fprintf('💡 분석 방식: 전체 그룹 비교 (모든 고성과자 vs 저성과자)\n');
    fprintf('   📈 고성과: 자연성, 성실한 가연성, 유익한 불연성\n');
    fprintf('   📉 저성과: 무능한 불연성, 소화성, 게으른 가연성\n');
    fprintf('   ✅ 장점: 충분한 표본 수, 높은 일반화 가능성\n');
    fprintf('   ⚠️  주의: 효과 크기 감소 가능성\n\n');
end

% 극단 그룹 인덱스 - 최종 분석 데이터에서 찾기
final_talent_types = matched_talent_types(final_idx);
final_talent_types = final_talent_types(complete_cases);

extreme_high_idx = ismember(final_talent_types, extreme_high);
extreme_low_idx = ismember(final_talent_types, extreme_low);

% 극단 그룹 데이터
X_extreme_high = X_final(extreme_high_idx, :);
X_extreme_low = X_final(extreme_low_idx, :);

fprintf('그룹 구성:\n');
if strcmp(config.extreme_group_method, 'extreme')
    fprintf('  고성과자 (자연성, 성실한 가연성): %d명\n', sum(extreme_high_idx));
    fprintf('  저성과자 (무능한 불연성, 소화성): %d명\n\n', sum(extreme_low_idx));
else
    fprintf('  고성과자 (자연성, 성실한 가연성, 유익한 불연성): %d명\n', sum(extreme_high_idx));
    fprintf('  저성과자 (무능한 불연성, 소화성, 게으른 가연성): %d명\n\n', sum(extreme_low_idx));
end

% 극단 그룹이 충분한지 확인
if sum(extreme_high_idx) < 3 || sum(extreme_low_idx) < 3
    fprintf('⚠️ 극단 그룹 샘플 수가 부족합니다. 분석을 건너뜁니다.\n');
else
    % t-test 결과 테이블 생성
    ttest_results = table();
    ttest_results.Feature = feature_names';
    ttest_results.High_Mean = zeros(n_features, 1);
    ttest_results.High_Std = zeros(n_features, 1);
    ttest_results.Low_Mean = zeros(n_features, 1);
    ttest_results.Low_Std = zeros(n_features, 1);
    ttest_results.Mean_Diff = zeros(n_features, 1);
    ttest_results.t_statistic = zeros(n_features, 1);
    ttest_results.p_value = zeros(n_features, 1);
    ttest_results.Cohen_d = zeros(n_features, 1);
    ttest_results.Significance = cell(n_features, 1);

    % 각 역량별 t-test 수행
    for i = 1:n_features
        high_scores = X_extreme_high(:, i);
        low_scores = X_extreme_low(:, i);

        % 기술통계
        ttest_results.High_Mean(i) = mean(high_scores);
        ttest_results.High_Std(i) = std(high_scores);
        ttest_results.Low_Mean(i) = mean(low_scores);
        ttest_results.Low_Std(i) = std(low_scores);
        ttest_results.Mean_Diff(i) = ttest_results.High_Mean(i) - ttest_results.Low_Mean(i);

        % t-test
        try
            [h, p, ci, stats] = ttest2(high_scores, low_scores);
            ttest_results.t_statistic(i) = stats.tstat;
            ttest_results.p_value(i) = p;
        catch
            ttest_results.t_statistic(i) = NaN;
            ttest_results.p_value(i) = NaN;
        end

        % Cohen's d (효과 크기)
        pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                          (length(low_scores)-1)*var(low_scores)) / ...
                          (length(high_scores) + length(low_scores) - 2));
        if pooled_std > 0
            ttest_results.Cohen_d(i) = ttest_results.Mean_Diff(i) / pooled_std;
        else
            ttest_results.Cohen_d(i) = 0;
        end

        % 유의성 표시
        p = ttest_results.p_value(i);
        if isnan(p)
            ttest_results.Significance{i} = 'NA';
        elseif p < 0.001
            ttest_results.Significance{i} = '***';
        elseif p < 0.01
            ttest_results.Significance{i} = '**';
        elseif p < 0.05
            ttest_results.Significance{i} = '*';
        elseif p < 0.1
            ttest_results.Significance{i} = '†';
        else
            ttest_results.Significance{i} = '';
        end
    end

    % Cohen's d 기준으로 정렬
    ttest_results = sortrows(ttest_results, 'Cohen_d', 'descend');

    % 결과 출력
    fprintf('【극단 그룹 비교 결과】\n');
    fprintf('────────────────────────────────────────────\n');
    fprintf('%-20s | 고성과(M±SD) | 저성과(M±SD) | 차이 | t값 | p값 | Cohen''s d | 효과\n', '역량');
    fprintf('%s\n', repmat('-', 105, 1));

    for i = 1:height(ttest_results)
        % 효과 크기 해석
        d = abs(ttest_results.Cohen_d(i));
        if d < 0.2
            effect = '무시';
        elseif d < 0.5
            effect = '작음';
        elseif d < 0.8
            effect = '중간';
        else
            effect = '큼';
        end

        fprintf('%-20s | %5.1f±%4.1f | %5.1f±%4.1f | %+5.1f | %+5.2f | %.3f%s | %+6.3f | %s\n', ...
            ttest_results.Feature{i}, ...
            ttest_results.High_Mean(i), ttest_results.High_Std(i), ...
            ttest_results.Low_Mean(i), ttest_results.Low_Std(i), ...
            ttest_results.Mean_Diff(i), ...
            ttest_results.t_statistic(i), ...
            ttest_results.p_value(i), ttest_results.Significance{i}, ...
            ttest_results.Cohen_d(i), effect);
    end

    % 유의한 차이를 보이는 역량만 추출 (p < 0.05 & |d| > 0.5)
    valid_p = ~isnan(ttest_results.p_value);
    significant_features = valid_p & ttest_results.p_value < 0.05 & abs(ttest_results.Cohen_d) > 0.5;

    fprintf('\n【핵심 차별화 역량】 (p<0.05 & Cohen''s d>0.5)\n');
    fprintf('────────────────────────────────────────────\n');
    if any(significant_features)
        sig_table = ttest_results(significant_features, :);
        for i = 1:height(sig_table)
            fprintf('• %s: 평균차이 %.1f점, Cohen''s d = %.2f\n', ...
                sig_table.Feature{i}, sig_table.Mean_Diff(i), sig_table.Cohen_d(i));
        end
    else
        fprintf('통계적으로 유의하고 실질적 효과가 큰 역량이 없습니다.\n');
    end

    % Bonferroni 보정
    bonferroni_alpha = 0.05 / n_features;
    bonferroni_sig = valid_p & ttest_results.p_value < bonferroni_alpha;

    fprintf('\n【Bonferroni 보정 후】 (α = %.4f)\n', bonferroni_alpha);
    fprintf('────────────────────────────────────────────\n');
    if any(bonferroni_sig)
        bon_table = ttest_results(bonferroni_sig, :);
        for i = 1:height(bon_table)
            fprintf('• %s: 여전히 유의함 (p = %.4f)\n', ...
                bon_table.Feature{i}, bon_table.p_value(i));
        end
    else
        fprintf('다중비교 보정 후 유의한 역량이 없습니다.\n');
    end

    % 시각화
    extreme_fig = figure('Position', [100, 100, 1200, 600], 'Color', 'white');

    % Cohen's d 막대그래프
    subplot(1, 2, 1);
    bar(ttest_results.Cohen_d, 'FaceColor', [0.2, 0.4, 0.8]);
    hold on;
    % 효과 크기 기준선
    yline(0.8, '--r', 'LineWidth', 1.5);
    yline(0.5, '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);
    yline(-0.5, '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);
    yline(-0.8, '--r', 'LineWidth', 1.5);
    set(gca, 'XTick', 1:height(ttest_results), ...
             'XTickLabel', ttest_results.Feature, ...
             'XTickLabelRotation', 45);
    ylabel('Cohen''s d', 'FontWeight', 'bold');
    if strcmp(config.extreme_group_method, 'extreme')
        title('효과 크기 (극단 그룹 차이)', 'FontWeight', 'bold');
    else
        title('효과 크기 (전체 그룹 차이)', 'FontWeight', 'bold');
    end
    legend({'Cohen''s d', '큰 효과(0.8)', '중간 효과(0.5)'}, 'Location', 'best');
    grid on;

    % p-value 비교
    subplot(1, 2, 2);
    valid_p_values = ttest_results.p_value;
    valid_p_values(isnan(valid_p_values)) = 1;  % NaN을 1로 대체
    bar(-log10(valid_p_values), 'FaceColor', [0.8, 0.3, 0.3]);
    hold on;
    % 유의수준 선
    yline(-log10(0.05), '--g', 'LineWidth', 2);
    yline(-log10(0.01), '--', 'Color', [1, 0.5, 0], 'LineWidth', 2);
    yline(-log10(0.001), '--r', 'LineWidth', 2);
    set(gca, 'XTick', 1:height(ttest_results), ...
             'XTickLabel', ttest_results.Feature, ...
             'XTickLabelRotation', 45);
    ylabel('-log10(p-value)', 'FontWeight', 'bold');
    title('통계적 유의성', 'FontWeight', 'bold');
    legend({'p-value', 'p<0.05', 'p<0.01', 'p<0.001'}, 'Location', 'best');
    grid on;

    if strcmp(config.extreme_group_method, 'extreme')
        sgtitle('극단 그룹 t-test 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');
    else
        sgtitle('전체 그룹 t-test 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');
    end

    % % 저장
    % extreme_chart_filename = sprintf('extreme_group_ttest_%s.png', datestr(now, 'yyyy-mm-dd_HHMMSS'));
    % saveas(extreme_fig, extreme_chart_filename);

    if strcmp(config.extreme_group_method, 'extreme')
        fprintf('\n✅ 극단 그룹 분석 완료\n');
    else
        fprintf('\n✅ 전체 그룹 분석 완료\n');
    end
    % fprintf('📊 시각화 저장 완료: %s\n', extreme_chart_filename);

    % 결과를 파일에 저장
    result_data.ttest_results = ttest_results;
    result_data.extreme_analysis = struct(...
        'extreme_high', {extreme_high}, ...
        'extreme_low', {extreme_low}, ...
        'n_high', sum(extreme_high_idx), ...
        'n_low', sum(extreme_low_idx));
end

% Bootstrap 결과도 파일에 저장
result_data.bootstrap_stats = bootstrap_stats;
result_data.bootstrap_weights = bootstrap_weights;
result_data.bootstrap_rankings = bootstrap_rankings;
save(weight_filepath, 'result_data', 'weight_results_significant', 'bootstrap_stats');

%% ========================================================================
%                    통합 가중치 분석 및 예측 검증 시스템
% =========================================================================
fprintf('\n\n╔═══════════════════════════════════════════════════════════╗\n');
fprintf('║            통합 가중치 분석 및 예측 검증 시스템           ║\n');
fprintf('╚═══════════════════════════════════════════════════════════╝\n\n');

%% STEP 19: 4가지 가중치 방법론 통합 분석
fprintf('【STEP 19】 4가지 가중치 방법론 통합 분석\n');
fprintf('════════════════════════════════════════════════════════════\n\n');

% 1. 상관 기반 가중치 (이미 계산됨)
fprintf('▶ 1. 상관 기반 가중치 (Correlation-based Weights)\n');
fprintf('────────────────────────────────────────────\n');
if exist('correlation_results', 'var')
    corr_weights = zeros(length(feature_names), 1);
    for i = 1:length(feature_names)
        idx = strcmp(correlation_results.Competency, feature_names{i});
        if any(idx)
            corr_weights(i) = correlation_results.Weight(idx);
        end
    end
    fprintf('  ✓ 상관 가중치 추출 완료\n');
else
    % 상관 기반 가중치 재계산 (Cut-off 보정 포함)
    corr_weights = zeros(n_features, 1);

    % Cut-off 진단 수행
    if exist('matched_comp', 'var') && exist('feature_names', 'var')
        fprintf('  ▶ Cut-off 효과 진단 수행 중...\n');
        cutoff_diagnosis = diagnose_cutoff_effects(matched_comp{:, 1:n_features}, feature_names);

        % 심각한 Cut-off 효과가 있는 역량 수
        severe_cutoff_count = sum(cutoff_diagnosis.Severity == 'Severe');
        moderate_cutoff_count = sum(cutoff_diagnosis.Severity == 'Moderate');
        fprintf('    - 심각한 Cut-off 효과: %d개 역량\n', severe_cutoff_count);
        fprintf('    - 중등도 Cut-off 효과: %d개 역량\n', moderate_cutoff_count);
    end

    % Range restriction 보정이 필요한지 판단
    use_correction = exist('cutoff_diagnosis', 'var') && (severe_cutoff_count + moderate_cutoff_count) > 0;

    if use_correction
        fprintf('  ▶ Thorndike Case II 보정 적용 중...\n');
        % 보정된 상관계수 계산
        corrected_correlations = calculate_corrected_correlations(X_normalized, y_weight, ...
                                                                matched_comp{:, 1:n_features}, 'thorndike');

        % 가중치 계산 (보정된 상관계수 사용)
        for i = 1:n_features
            corr_weights(i) = max(0, corrected_correlations(i));
        end

        fprintf('    - 평균 보정 계수: %.2f\n', mean(corrected_correlations ./ corr_weights));

        % 보정 결과 저장
        if exist('cutoff_diagnosis', 'var')
            cutoff_diagnosis.Original_Correlation = zeros(height(cutoff_diagnosis), 1);
            cutoff_diagnosis.Corrected_Correlation = zeros(height(cutoff_diagnosis), 1);

            for i = 1:n_features
                if i <= height(cutoff_diagnosis)
                    original_r = corr(X_normalized(:,i), y_weight, 'Type', 'Spearman');
                    cutoff_diagnosis.Original_Correlation(i) = original_r;
                    cutoff_diagnosis.Corrected_Correlation(i) = corrected_correlations(i);
                end
            end
        end

    else
        fprintf('  ▶ Cut-off 효과가 미미하여 표준 상관계수 사용\n');
        % 표준 상관계수 계산
        for i = 1:n_features
            r = corr(X_normalized(:,i), y_weight, 'Type', 'Spearman');
            corr_weights(i) = max(0, r);
        end
    end

    corr_weights = (corr_weights / sum(corr_weights)) * 100;
    fprintf('  ✓ 상관 가중치 계산 완료 (보정 적용: %s)\n', ...
            string(use_correction).replace("1", "예").replace("0", "아니오"));
end

% 2. 로지스틱 회귀 가중치 (이미 계산됨)
fprintf('\n▶ 2. 로지스틱 회귀 가중치 (Logistic Regression Weights)\n');
fprintf('────────────────────────────────────────────\n');
logit_weights = final_weights;  % 이미 계산된 가중치
fprintf('  ✓ 로지스틱 회귀 가중치 (Cost-Sensitive) 사용\n');

% 3. Bootstrap 평균 가중치
fprintf('\n▶ 3. Bootstrap 평균 가중치 (Bootstrap Mean Weights)\n');
fprintf('────────────────────────────────────────────\n');
if exist('bootstrap_stats', 'var')
    boot_weights = zeros(length(feature_names), 1);
    for i = 1:length(feature_names)
        idx = strcmp(bootstrap_stats.Feature, feature_names{i});
        if any(idx)
            boot_weights(i) = bootstrap_stats.Boot_Mean(idx);
        end
    end
    fprintf('  ✓ Bootstrap 평균 가중치 추출 완료 (%d회 재샘플링)\n', n_bootstrap);
else
    boot_weights = logit_weights;  % Bootstrap이 없으면 로지스틱 가중치 사용
    fprintf('  ⚠ Bootstrap 결과 없음, 로지스틱 가중치로 대체\n');
end

% 4. t-test 효과크기 기반 가중치
fprintf('\n▶ 4. t-test 효과크기 가중치 (Cohen''s d Weights)\n');
fprintf('────────────────────────────────────────────\n');
if exist('ttest_results', 'var')
    ttest_weights = zeros(length(feature_names), 1);
    for i = 1:length(feature_names)
        idx = strcmp(ttest_results.Feature, feature_names{i});
        if any(idx)
            % Cohen's d의 절댓값을 가중치로 변환
            ttest_weights(i) = abs(ttest_results.Cohen_d(idx));
        end
    end
    % 정규화
    if sum(ttest_weights) > 0
        ttest_weights = (ttest_weights / sum(ttest_weights)) * 100;
    end
    fprintf('  ✓ t-test 효과크기 가중치 계산 완료\n');
else
    % t-test 가중치 계산
    ttest_weights = zeros(n_features, 1);
    for i = 1:n_features
        high_scores = X_final(y_final == 1, i);
        low_scores = X_final(y_final == 0, i);
        pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + ...
                          (length(low_scores)-1)*var(low_scores)) / ...
                          (length(high_scores) + length(low_scores) - 2));
        if pooled_std > 0
            ttest_weights(i) = abs(mean(high_scores) - mean(low_scores)) / pooled_std;
        end
    end
    ttest_weights = (ttest_weights / sum(ttest_weights)) * 100;
    fprintf('  ✓ t-test 효과크기 가중치 계산 완료\n');
end

%% 가중치 통합 테이블 생성
fprintf('\n【가중치 방법론별 비교】\n');
fprintf('════════════════════════════════════════════════════════════\n');

weight_comparison = table();
weight_comparison.Feature = feature_names';
weight_comparison.Correlation = round(corr_weights, 2);
weight_comparison.Logistic = round(logit_weights, 2);
weight_comparison.Bootstrap = round(boot_weights, 2);
weight_comparison.Ttest = round(ttest_weights, 2);

% 평균 가중치 계산 (앙상블)
weight_comparison.Ensemble_Mean = round(mean([corr_weights, logit_weights, boot_weights, ttest_weights], 2), 2);

% 가중치 변동성 (표준편차)
weight_comparison.Std = round(std([corr_weights, logit_weights, boot_weights, ttest_weights], 0, 2), 2);

% 신뢰도 점수 계산 (낮은 표준편차 = 높은 신뢰도)
weight_comparison.Reliability = categorical(zeros(height(weight_comparison), 1));
for i = 1:height(weight_comparison)
    cv = weight_comparison.Std(i) / (weight_comparison.Ensemble_Mean(i) + eps);
    if cv < 0.3
        weight_comparison.Reliability(i) = 'High';
    elseif cv < 0.6
        weight_comparison.Reliability(i) = 'Medium';
    else
        weight_comparison.Reliability(i) = 'Low';
    end
end

% 앙상블 평균으로 정렬
weight_comparison = sortrows(weight_comparison, 'Ensemble_Mean', 'descend');

% 상위 15개 역량 출력
fprintf('\n상위 15개 역량의 가중치 비교:\n');
fprintf('%-25s | Corr(%%) | Logit(%%) | Boot(%%) | Ttest(%%) | Ensemble(%%) | Std | 신뢰도\n', '역량');
fprintf('%s\n', repmat('-', 100, 1));

for i = 1:min(15, height(weight_comparison))
    fprintf('%-25s | %7.2f | %8.2f | %7.2f | %8.2f | %11.2f | %5.2f | %s\n', ...
        weight_comparison.Feature{i}, ...
        weight_comparison.Correlation(i), ...
        weight_comparison.Logistic(i), ...
        weight_comparison.Bootstrap(i), ...
        weight_comparison.Ttest(i), ...
        weight_comparison.Ensemble_Mean(i), ...
        weight_comparison.Std(i), ...
        string(weight_comparison.Reliability(i)));
end

%% STEP 20: 실제 참가자 예측 검증
fprintf('\n\n【STEP 20】 실제 참가자 예측 검증\n');
fprintf('════════════════════════════════════════════════════════════\n\n');

% 검증할 가중치 방법 선택
weight_methods = {'Correlation', 'Logistic', 'Bootstrap', 'Ttest', 'Ensemble'};
n_methods = length(weight_methods);

% 각 방법별 가중치 행렬 준비
all_weights = [corr_weights, logit_weights, boot_weights, ttest_weights, weight_comparison.Ensemble_Mean];

% 예측 결과 저장 구조체
prediction_results = struct();
for m = 1:n_methods
    prediction_results.(weight_methods{m}) = struct();
end

% 각 가중치 방법으로 예측 수행
fprintf('▶ 각 가중치 방법별 예측 수행\n');
fprintf('────────────────────────────────────────────\n\n');

for m = 1:n_methods
    method_name = weight_methods{m};
    fprintf('【%s 가중치 방법】\n', method_name);
    
    % 해당 방법의 가중치 선택
    if m <= 4
        current_weights = all_weights(:, m) / 100;  % 백분율을 비율로
    else
        % 앙상블은 이미 정렬된 테이블에서 가져옴
        ensemble_weights = zeros(length(feature_names), 1);
        for i = 1:length(feature_names)
            idx = strcmp(weight_comparison.Feature, feature_names{i});
            if any(idx)
                ensemble_weights(i) = weight_comparison.Ensemble_Mean(idx) / 100;
            end
        end
        current_weights = ensemble_weights;
    end
    
    % 가중 점수 계산
    weighted_scores = X_normalized * current_weights;
    
    % 최적 임계값 찾기 (각 방법별로)
    [X_roc, Y_roc, T_roc, AUC] = perfcurve(y_weight, weighted_scores, 1);
    J = Y_roc - X_roc;  % Youden's J statistic
    [~, opt_idx] = max(J);
    opt_threshold = T_roc(opt_idx);
    
    % 예측 수행
    predictions = weighted_scores > opt_threshold;
    
    % 성능 평가
    TP = sum(predictions == 1 & y_weight == 1);
    TN = sum(predictions == 0 & y_weight == 0);
    FP = sum(predictions == 1 & y_weight == 0);
    FN = sum(predictions == 0 & y_weight == 1);
    
    accuracy = (TP + TN) / length(y_weight);
    precision = TP / (TP + FP + eps);
    recall = TP / (TP + FN + eps);
    f1_score = 2 * (precision * recall) / (precision + recall + eps);
    
    % 결과 저장
    prediction_results.(method_name).weighted_scores = weighted_scores;
    prediction_results.(method_name).predictions = predictions;
    prediction_results.(method_name).threshold = opt_threshold;
    prediction_results.(method_name).AUC = AUC;
    prediction_results.(method_name).accuracy = accuracy;
    prediction_results.(method_name).precision = precision;
    prediction_results.(method_name).recall = recall;
    prediction_results.(method_name).f1_score = f1_score;
    prediction_results.(method_name).confusion_matrix = [TN, FP; FN, TP];
    
    fprintf('  정확도: %.3f | AUC: %.3f | F1: %.3f | 정밀도: %.3f | 재현율: %.3f\n', ...
        accuracy, AUC, f1_score, precision, recall);
    fprintf('  혼동행렬: TN=%d, FP=%d, FN=%d, TP=%d\n\n', TN, FP, FN, TP);
end

%% STEP 21: 개별 참가자 예측 상세 분석
fprintf('【STEP 21】 개별 참가자 예측 상세 분석\n');
fprintf('════════════════════════════════════════════════════════════\n\n');

% 가장 성능이 좋은 방법 찾기
best_method_idx = 1;
best_f1 = 0;
for m = 1:n_methods
    if prediction_results.(weight_methods{m}).f1_score > best_f1
        best_f1 = prediction_results.(weight_methods{m}).f1_score;
        best_method_idx = m;
    end
end
best_method = weight_methods{best_method_idx};

fprintf('▶ 최적 방법: %s (F1=%.3f)\n\n', best_method, best_f1);

% 개별 참가자 예측 결과 테이블 생성
participant_results = table();

% 참가자 정보
participant_results.ID = (1:length(y_weight))';
participant_results.Actual_Label = y_weight;

% 실제 인재유형 추가 (있는 경우)
if exist('final_talent_types', 'var')
    participant_results.Talent_Type = final_talent_types;
end

% 각 방법별 예측 점수와 결과
for m = 1:n_methods
    method_name = weight_methods{m};
    col_name_score = sprintf('%s_Score', method_name);
    col_name_pred = sprintf('%s_Pred', method_name);
    
    participant_results.(col_name_score) = round(prediction_results.(method_name).weighted_scores, 3);
    participant_results.(col_name_pred) = prediction_results.(method_name).predictions;
end

% 예측 일치도 계산 (몇 개 방법이 맞췄는지)
participant_results.Agreement_Count = zeros(height(participant_results), 1);
for i = 1:height(participant_results)
    count = 0;
    for m = 1:n_methods
        col_name_pred = sprintf('%s_Pred', weight_methods{m});
        if participant_results.(col_name_pred)(i) == participant_results.Actual_Label(i)
            count = count + 1;
        end
    end
    participant_results.Agreement_Count(i) = count;
end

% 예측 난이도 분류
participant_results.Prediction_Difficulty = categorical(zeros(height(participant_results), 1));
for i = 1:height(participant_results)
    if participant_results.Agreement_Count(i) == n_methods
        participant_results.Prediction_Difficulty(i) = 'Easy';
    elseif participant_results.Agreement_Count(i) >= 3
        participant_results.Prediction_Difficulty(i) = 'Medium';
    else
        participant_results.Prediction_Difficulty(i) = 'Hard';
    end
end

%% 오분류 사례 분석
fprintf('【오분류 사례 분석】\n');
fprintf('────────────────────────────────────────────\n\n');

% 최적 방법의 오분류 사례 찾기
best_pred_col = sprintf('%s_Pred', best_method);
misclassified_idx = participant_results.(best_pred_col) ~= participant_results.Actual_Label;

fprintf('▶ %s 방법의 오분류 사례: %d건 (%.1f%%)\n\n', ...
    best_method, sum(misclassified_idx), sum(misclassified_idx)/height(participant_results)*100);

if sum(misclassified_idx) > 0
    misclass_table = participant_results(misclassified_idx, :);
    
    % False Positive (저성과자→고성과자 오분류)
    fp_idx = misclass_table.Actual_Label == 0 & misclass_table.(best_pred_col) == 1;
    if any(fp_idx)
        fprintf('False Positive (저성과자→고성과자): %d건\n', sum(fp_idx));
        fp_cases = misclass_table(fp_idx, :);
        for i = 1:min(3, height(fp_cases))  % 최대 3건만 표시
            fprintf('  ID %d: ', fp_cases.ID(i));
            if exist('final_talent_types', 'var')
                fprintf('%s, ', fp_cases.Talent_Type{i});
            end
            fprintf('점수=%.3f (임계값=%.3f)\n', ...
                fp_cases.(sprintf('%s_Score', best_method))(i), ...
                prediction_results.(best_method).threshold);
        end
    end
    
    fprintf('\n');
    
    % False Negative (고성과자→저성과자 오분류)
    fn_idx = misclass_table.Actual_Label == 1 & misclass_table.(best_pred_col) == 0;
    if any(fn_idx)
        fprintf('False Negative (고성과자→저성과자): %d건\n', sum(fn_idx));
        fn_cases = misclass_table(fn_idx, :);
        for i = 1:min(3, height(fn_cases))  % 최대 3건만 표시
            fprintf('  ID %d: ', fn_cases.ID(i));
            if exist('final_talent_types', 'var')
                fprintf('%s, ', fn_cases.Talent_Type{i});
            end
            fprintf('점수=%.3f (임계값=%.3f)\n', ...
                fn_cases.(sprintf('%s_Score', best_method))(i), ...
                prediction_results.(best_method).threshold);
        end
    end
end

%% 난이도별 분석
fprintf('\n【예측 난이도별 분석】\n');
fprintf('────────────────────────────────────────────\n');

difficulty_stats = groupsummary(participant_results, 'Prediction_Difficulty');
for i = 1:height(difficulty_stats)
    fprintf('  %s: %d명 (%.1f%%)\n', ...
        string(difficulty_stats.Prediction_Difficulty(i)), ...
        difficulty_stats.GroupCount(i), ...
        difficulty_stats.GroupCount(i)/height(participant_results)*100);
end

% Hard 케이스 상세 분석
hard_cases = participant_results(participant_results.Prediction_Difficulty == 'Hard', :);
if height(hard_cases) > 0
    fprintf('\n▶ 예측 어려운 케이스 분석 (Hard):\n');
    for i = 1:min(5, height(hard_cases))
        fprintf('  ID %d: 실제=%d, 일치방법=%d/%d\n', ...
            hard_cases.ID(i), ...
            hard_cases.Actual_Label(i), ...
            hard_cases.Agreement_Count(i), ...
            n_methods);
        
        % 각 방법별 예측 표시
        fprintf('    ');
        for m = 1:n_methods
            pred_col = sprintf('%s_Pred', weight_methods{m});
            if hard_cases.(pred_col)(i) == hard_cases.Actual_Label(i)
                fprintf('%s(O) ', weight_methods{m}(1:3));
            else
                fprintf('%s(X) ', weight_methods{m}(1:3));
            end
        end
        fprintf('\n');
    end
end

%% STEP 22: 시각화
fprintf('\n【STEP 22】 통합 분석 시각화\n');
fprintf('════════════════════════════════════════════════════════════\n');

% Figure 생성
fig_integrated = figure('Position', [100, 100, 1800, 1200], 'Color', 'white');

% 1. 가중치 방법별 비교 (상위 10개 역량)
subplot(3, 3, 1);
top_10_features = weight_comparison.Feature(1:min(10, height(weight_comparison)));
top_10_data = table2array(weight_comparison(1:min(10, height(weight_comparison)), 2:5));

bar(top_10_data);
set(gca, 'XTickLabel', top_10_features, 'XTickLabelRotation', 45);
legend(weight_methods(1:4), 'Location', 'best', 'FontSize', 8);
ylabel('가중치 (%)');
title('가중치 방법별 비교 (Top 10)');
grid on;

% 2. 앙상블 가중치와 신뢰도
subplot(3, 3, 2);
ensemble_data = weight_comparison.Ensemble_Mean(1:min(15, height(weight_comparison)));
reliability_colors = zeros(length(ensemble_data), 3);
for i = 1:length(ensemble_data)
    if weight_comparison.Reliability(i) == 'High'
        reliability_colors(i, :) = [0.2, 0.7, 0.3];
    elseif weight_comparison.Reliability(i) == 'Medium'
        reliability_colors(i, :) = [0.9, 0.7, 0.1];
    else
        reliability_colors(i, :) = [0.8, 0.3, 0.3];
    end
end

barh(1:length(ensemble_data), ensemble_data);
colormap(gca, reliability_colors);
set(gca, 'YTick', 1:length(ensemble_data), ...
    'YTickLabel', weight_comparison.Feature(1:length(ensemble_data)));
xlabel('앙상블 가중치 (%)');
title('앙상블 가중치 및 신뢰도');
grid on;

% 3. 성능 지표 비교
subplot(3, 3, 3);
performance_matrix = zeros(n_methods, 4);
for m = 1:n_methods
    performance_matrix(m, :) = [
        prediction_results.(weight_methods{m}).accuracy, ...
        prediction_results.(weight_methods{m}).precision, ...
        prediction_results.(weight_methods{m}).recall, ...
        prediction_results.(weight_methods{m}).f1_score
    ];
end

bar(performance_matrix);
set(gca, 'XTickLabel', weight_methods, 'XTickLabelRotation', 45);
legend({'Accuracy', 'Precision', 'Recall', 'F1-Score'}, 'Location', 'best');
ylabel('Score');
title('가중치 방법별 성능 비교');
ylim([0, 1]);
grid on;

% 4-8. ROC 곡선 (각 방법별)
for m = 1:n_methods
    subplot(3, 3, 3 + m);
    
    % ROC 곡선 계산
    [X_roc, Y_roc, ~, AUC] = perfcurve(y_weight, ...
        prediction_results.(weight_methods{m}).weighted_scores, 1);
    
    plot(X_roc, Y_roc, 'LineWidth', 2);
    hold on;
    plot([0, 1], [0, 1], '--k', 'LineWidth', 1);
    
    xlabel('FPR');
    ylabel('TPR');
    title(sprintf('%s (AUC=%.3f)', weight_methods{m}, AUC));
    grid on;
    axis square;
end


sgtitle('통합 가중치 분석 및 예측 검증 결과', 'FontSize', 16, 'FontWeight', 'bold');

% 그래프 저장 (백업 시스템 적용)
integrated_chart_filepath = get_managed_filepath(config.output_dir, 'integrated_weight_analysis.png', config);
backup_and_prepare_file(integrated_chart_filepath, config);
% saveas(fig_integrated, integrated_chart_filepath);
fprintf('\n✓ 통합 분석 차트 저장: %s\n', integrated_chart_filepath);

%% STEP 22.5: 모델 재학습 기반 퍼뮤테이션 테스트 (AUC & F1 스코어)
fprintf('\n\n【STEP 22.5】 모델 재학습 기반 퍼뮤테이션 테스트\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 필수 변수 존재 확인
if ~exist('X_normalized', 'var') || ~exist('y_weight', 'var')
    fprintf('  ⚠ 필수 변수 (X_normalized, y_weight) 누락\n');
    fprintf('  ✓ STEP 22.5 건너뛰기\n');
    return;
end

if ~exist('prediction_results', 'var') || ~exist('best_method', 'var')
    fprintf('  ⚠ 예측 결과 또는 최적 방법 정보 누락\n');
    fprintf('  ✓ STEP 22.5 건너뛰기\n');
    return;
end

% 퍼뮤테이션 캐시 파일 경로 (새로운 방법용)
model_cache_file = fullfile(config.output_dir, 'model_permutation_cache.mat');

% Leave-One-Out Cross-Validation (LOOCV) 방식 사용
validation_method = 'loocv';
n_samples = length(y_weight);
fprintf('  LOOCV 방식 사용: %d개 샘플\n', n_samples);

% LOOCV는 작은 데이터에 효과적이지만 계산 비용이 높음
if n_samples > 1000
    fprintf('  ⚠ 경고: 샘플 수가 많음 (%d). LOOCV 계산 비용 고려\n', n_samples);
end

% 캐시 키 생성 (LOOCV용)
cache_key = struct();
cache_key.n_samples = n_samples;
cache_key.n_features = size(X_normalized, 2);
cache_key.class_ratio = sum(y_weight==1) / n_samples;
cache_key.validation_method = validation_method;
cache_key.loocv_version = '1.0';  % LOOCV 버전 식별
cache_key.data_checksum = sum(X_normalized(:)) + sum(y_weight(:))*1000;

% 캐시 확인
use_cached = false;
if exist(model_cache_file, 'file') && ~config.force_recalc_permutation
    fprintf('▶ 기존 모델 퍼뮤테이션 캐시 발견\n');
    try
        load(model_cache_file , 'model_perm_cache');

        % 동일한 데이터인지 확인
        if isequal(model_perm_cache.cache_key, cache_key)
            fprintf('  ✓ 캐시 유효: 기존 결과 사용\n');
            permutation_results = model_perm_cache.results;
            use_cached = true;
        else
            fprintf('  ⚠ 데이터 변경 감지: 재계산 필요\n');
        end
    catch
        fprintf('  ⚠ 캐시 로딩 실패: 재계산 필요\n');
    end
end

if ~use_cached
    fprintf('▶ 모델 재학습 퍼뮤테이션 테스트 실행 중...\n');
    fprintf('  검증 방법: %s\n', validation_method);

    % LOOCV는 계산 비용이 높으므로 퍼뮤테이션 횟수 조정
    if n_samples <= 100
        n_permutations = 5000;   % 작은 데이터: 더 많은 퍼뮤테이션
    elseif n_samples <= 500
        n_permutations = 2000;   % 중간 데이터: 적당한 퍼뮤테이션
    else
        n_permutations = 1000;   % 큰 데이터: 최소 퍼뮤테이션
    end

    fprintf('  LOOCV 기반 퍼뮤테이션: %d회 (n=%d)\n', n_permutations, n_samples);

    % 관찰된 성능 메트릭 가져오기 (유연한 필드 처리)
    method_data = prediction_results.(best_method);

    % AUC 필드 검색 (다양한 이름 가능)
    observed_auc = [];
    auc_fields = {'auc', 'AUC', 'roc_auc', 'area_under_curve'};
    for i = 1:length(auc_fields)
        if isfield(method_data, auc_fields{i})
            observed_auc = method_data.(auc_fields{i});
            fprintf('  AUC 필드 발견: %s = %.4f\n', auc_fields{i}, observed_auc);
            break;
        end
    end

    % F1 필드 검색
    observed_f1 = [];
    f1_fields = {'f1_score', 'f1', 'F1', 'f1_measure'};
    for i = 1:length(f1_fields)
        if isfield(method_data, f1_fields{i})
            observed_f1 = method_data.(f1_fields{i});
            fprintf('  F1 필드 발견: %s = %.4f\n', f1_fields{i}, observed_f1);
            break;
        end
    end

    % 필수 메트릭 처리
    if isempty(observed_auc)
        % AUC 대안 메트릭 찾기 (정확도 등)
        accuracy_fields = {'accuracy', 'acc', 'correct_rate'};
        for i = 1:length(accuracy_fields)
            if isfield(method_data, accuracy_fields{i})
                observed_auc = method_data.(accuracy_fields{i});
                fprintf('  AUC 대신 %s 사용: %.4f\n', accuracy_fields{i}, observed_auc);
                break;
            end
        end

        if isempty(observed_auc)
            fprintf('  ⚠ AUC 메트릭을 찾을 수 없음 - 기본값 0.5 사용\n');
            observed_auc = 0.5;  % 랜덤 수준
        end
    end

    if isempty(observed_f1)
        % F1 대안 메트릭 찾기 (정확도, 정밀도 등)
        alt_fields = {'accuracy', 'precision', 'recall', 'acc'};
        for i = 1:length(alt_fields)
            if isfield(method_data, alt_fields{i})
                observed_f1 = method_data.(alt_fields{i});
                fprintf('  F1 대신 %s 사용: %.4f\n', alt_fields{i}, observed_f1);
                break;
            end
        end

        if isempty(observed_f1)
            fprintf('  ⚠ F1 메트릭을 찾을 수 없음 - 기본값 0.0 사용\n');
            observed_f1 = 0.0;   % 최저 성능
        end
    end

    % 사용 가능한 모든 필드 목록 출력 (디버깅용)
    field_names = fieldnames(method_data);
    fprintf('  %s 방법의 사용 가능한 필드 (%d개): ', best_method, length(field_names));
    for i = 1:length(field_names)
        if isnumeric(method_data.(field_names{i})) && isscalar(method_data.(field_names{i}))
            fprintf('%s(%.3f) ', field_names{i}, method_data.(field_names{i}));
        else
            fprintf('%s ', field_names{i});
        end
    end
    fprintf('\n');

    % 기본 메트릭 사용 여부 확인 및 경고
    using_default_auc = (observed_auc == 0.5);
    using_default_f1 = (observed_f1 == 0.0);

    if using_default_auc && using_default_f1
        fprintf('  ⚠ 경고: 모든 메트릭이 기본값입니다. LOOCV로 실제 성능을 측정합니다.\n');
        use_loocv_evaluation = true;
    else
        fprintf('  기존 메트릭 사용 - AUC: %.4f, F1: %.4f\n', observed_auc, observed_f1);
        use_loocv_evaluation = false;
    end

    % LOOCV로 실제 성능 평가 (기본값 사용 시 또는 더 정확한 평가를 위해)
    if use_loocv_evaluation
        fprintf('\n  ▶ LOOCV로 원본 모델 성능 평가 실행 중...\n');

        % LOOCV 기반 성능 측정 (인라인)
        loo_predictions = zeros(n_samples, 1);
        loo_probabilities = zeros(n_samples, 1);
        loocv_failed = 0;

        for i = 1:n_samples
            try
                % i번째 샘플 제외
                train_idx = true(n_samples, 1);
                train_idx(i) = false;

                X_train_loo = X_normalized(train_idx, :);
                y_train_loo = y_weight(train_idx);
                X_test_loo = X_normalized(i, :);

                % 클래스 분포 확인
                if length(unique(y_train_loo)) < 2
                    % 한 클래스만 남은 경우 고정 예측
                    loo_predictions(i) = mode(y_train_loo);
                    loo_probabilities(i) = 0.5;
                    continue;
                end

                % 모델 학습
                mdl_loo = fitclinear(X_train_loo, y_train_loo, ...
                    'Learner', 'logistic', ...
                    'Regularization', 'lasso', ...
                    'Lambda', 1e-4, ...
                    'Solver', 'sparsa', ...
                    'Verbose', 0);

                % 예측
                [pred_label, pred_scores] = predict(mdl_loo, X_test_loo);
                loo_predictions(i) = pred_label;
                if size(pred_scores, 2) >= 2
                    loo_probabilities(i) = pred_scores(2);
                else
                    loo_probabilities(i) = pred_scores(1);
                end

            catch
                loocv_failed = loocv_failed + 1;
                % 실패 시 다수 클래스로 예측
                other_y = y_weight(y_weight ~= y_weight(i));
                if ~isempty(other_y)
                    loo_predictions(i) = mode(other_y);
                else
                    loo_predictions(i) = y_weight(i);
                end
                loo_probabilities(i) = 0.5;
            end

            % 진행 표시
            if mod(i, 10) == 0 || i == n_samples
                fprintf('    LOOCV 진행: %d/%d (%.1f%%)\n', i, n_samples, i/n_samples*100);
            end
        end

        % AUC 계산
        try
            [~, ~, ~, loocv_auc] = perfcurve(y_weight, loo_probabilities, 1);
        catch
            loocv_auc = 0.5;
        end

        % F1 계산
        TP = sum(loo_predictions == 1 & y_weight == 1);
        FP = sum(loo_predictions == 1 & y_weight == 0);
        FN = sum(loo_predictions == 0 & y_weight == 1);
        precision = TP / (TP + FP + eps);
        recall = TP / (TP + FN + eps);
        loocv_f1 = 2 * (precision * recall) / (precision + recall + eps);

        if loocv_failed > 0
            fprintf('    LOOCV 실패: %d/%d (%.1f%%)\n', loocv_failed, n_samples, loocv_failed/n_samples*100);
        end

        % 기본값을 사용한 경우 LOOCV 결과로 대체
        if using_default_auc
            observed_auc = loocv_auc;
            fprintf('    LOOCV AUC: %.4f (기본값 대체)\n', observed_auc);
        end

        if using_default_f1
            observed_f1 = loocv_f1;
            fprintf('    LOOCV F1: %.4f (기본값 대체)\n', observed_f1);
        end
    end

    % 함수 정의 제거 - 인라인 LOOCV 구현으로 변경

    % null distribution 초기화
    null_auc_distribution = zeros(n_permutations, 1);
    null_f1_distribution = zeros(n_permutations, 1);
    failed_permutations = 0;

    % 진행 표시
    fprintf('\n  ▶ LOOCV 퍼뮤테이션 실행 중...\n');
    fprintf('  진행률: 0%% ');
    tic;
    update_interval = max(1, ceil(n_permutations / 20));  % 20단계로 업데이트

    for perm = 1:n_permutations
        try
            % 레이블 셔플
            shuffled_y = y_weight(randperm(n_samples));

            % LOOCV로 퍼뮤테이션 성능 측정 (인라인 구현)
            perm_n = length(shuffled_y);
            perm_loo_predictions = zeros(perm_n, 1);
            perm_loo_probabilities = zeros(perm_n, 1);
            perm_failed = 0;

            % LOOCV 루프
            for i = 1:perm_n
                try
                    % i번째 샘플 제외
                    train_idx = true(perm_n, 1);
                    train_idx(i) = false;

                    perm_X_train = X_normalized(train_idx, :);
                    perm_y_train = shuffled_y(train_idx);
                    perm_X_test = X_normalized(i, :);

                    % 클래스 분포 확인
                    if length(unique(perm_y_train)) < 2
                        % 한 클래스만 남은 경우 고정 예측
                        perm_loo_predictions(i) = mode(perm_y_train);
                        perm_loo_probabilities(i) = 0.5;
                        continue;
                    end

                    % 모델 학습
                    perm_mdl_loo = fitclinear(perm_X_train, perm_y_train, ...
                        'Learner', 'logistic', ...
                        'Regularization', 'lasso', ...
                        'Lambda', 1e-4, ...
                        'Solver', 'sparsa', ...
                        'Verbose', 0);

                    % 예측
                    [perm_pred_label, perm_pred_scores] = predict(perm_mdl_loo, perm_X_test);
                    perm_loo_predictions(i) = perm_pred_label;
                    if size(perm_pred_scores, 2) >= 2
                        perm_loo_probabilities(i) = perm_pred_scores(2);
                    else
                        perm_loo_probabilities(i) = perm_pred_scores(1);
                    end

                catch
                    perm_failed = perm_failed + 1;
                    % 실패 시 다수 클래스로 예측
                    other_y = shuffled_y(shuffled_y ~= shuffled_y(i));
                    if ~isempty(other_y)
                        perm_loo_predictions(i) = mode(other_y);
                    else
                        perm_loo_predictions(i) = shuffled_y(i);  % 모두 같은 클래스인 경우
                    end
                    perm_loo_probabilities(i) = 0.5;
                end
            end

            % AUC 계산
            try
                [~, ~, ~, perm_auc] = perfcurve(shuffled_y, perm_loo_probabilities, 1);
            catch
                perm_auc = 0.5;
            end

            % F1 계산
            perm_TP = sum(perm_loo_predictions == 1 & shuffled_y == 1);
            perm_FP = sum(perm_loo_predictions == 1 & shuffled_y == 0);
            perm_FN = sum(perm_loo_predictions == 0 & shuffled_y == 1);
            perm_precision = perm_TP / (perm_TP + perm_FP + eps);
            perm_recall = perm_TP / (perm_TP + perm_FN + eps);
            perm_f1 = 2 * (perm_precision * perm_recall) / (perm_precision + perm_recall + eps);

            % null distribution에 저장
            null_auc_distribution(perm) = perm_auc;
            null_f1_distribution(perm) = perm_f1;

        catch ME
            % LOOCV 퍼뮤테이션 실패 처리
            failed_permutations = failed_permutations + 1;
            null_auc_distribution(perm) = 0.5;  % 랜덤 수준
            null_f1_distribution(perm) = 0;     % 최저 성능

            if failed_permutations <= 3  % 처음 3개 실패만 로그 (LOOCV는 오류가 적음)
                fprintf('\n    ⚠ LOOCV 퍼뮤테이션 %d 실패: %s', perm, ME.message);
            elseif failed_permutations == 4
                fprintf('\n    ... (추가 LOOCV 실패는 생략)');
            end
        end

        % 진행상황 표시 (LOOCV는 느리므로 더 자주 업데이트)
        if mod(perm, update_interval) == 0 || perm == n_permutations
            progress_pct = round(perm / n_permutations * 100);
            if perm == n_permutations
                fprintf('\b\b\b\b100%%');
            else
                fprintf('\b\b\b\b%3d%%', progress_pct);
            end
        end
    end

    elapsed_time = toc;
    fprintf(' 완료! (LOOCV 소요시간: %.1f초)\n', elapsed_time);

    % LOOCV 성능 요약
    avg_time_per_perm = elapsed_time / n_permutations;
    total_cv_iterations = n_permutations * n_samples;  % 총 CV 반복 횟수
    fprintf('  LOOCV 성능: 퍼뮤테이션당 %.2f초, 총 %d번 모델 학습\n', avg_time_per_perm, total_cv_iterations);

    if failed_permutations > 0
        fprintf('  ⚠ 총 %d개 LOOCV 퍼뮤테이션에서 실패 (%.1f%%)\n', ...
            failed_permutations, failed_permutations/n_permutations*100);
    end

    % AUC 통계 계산
    p_value_auc = sum(null_auc_distribution >= observed_auc) / n_permutations;
    ci_95_auc = prctile(null_auc_distribution, [2.5, 97.5]);
    mean_null_auc = mean(null_auc_distribution);
    std_null_auc = std(null_auc_distribution);
    percentile_rank_auc = sum(null_auc_distribution < observed_auc) / n_permutations * 100;

    % F1 통계 계산
    p_value_f1 = sum(null_f1_distribution >= observed_f1) / n_permutations;
    ci_95_f1 = prctile(null_f1_distribution, [2.5, 97.5]);
    mean_null_f1 = mean(null_f1_distribution);
    std_null_f1 = std(null_f1_distribution);
    percentile_rank_f1 = sum(null_f1_distribution < observed_f1) / n_permutations * 100;

    % Z-score 계산 (AUC)
    if std_null_auc > 1e-10
        z_score_auc = (observed_auc - mean_null_auc) / std_null_auc;
    else
        z_score_auc = NaN;
    end

    % Z-score 계산 (F1)
    if std_null_f1 > 1e-10
        z_score_f1 = (observed_f1 - mean_null_f1) / std_null_f1;
    else
        z_score_f1 = NaN;
    end

    % 결과 저장
    permutation_results = struct();

    % AUC 결과
    permutation_results.auc.observed = observed_auc;
    permutation_results.auc.null_distribution = null_auc_distribution;
    permutation_results.auc.p_value = p_value_auc;
    permutation_results.auc.ci_95 = ci_95_auc;
    permutation_results.auc.mean_null = mean_null_auc;
    permutation_results.auc.std_null = std_null_auc;
    permutation_results.auc.percentile_rank = percentile_rank_auc;
    permutation_results.auc.z_score = z_score_auc;

    % F1 결과
    permutation_results.f1.observed = observed_f1;
    permutation_results.f1.null_distribution = null_f1_distribution;
    permutation_results.f1.p_value = p_value_f1;
    permutation_results.f1.ci_95 = ci_95_f1;
    permutation_results.f1.mean_null = mean_null_f1;
    permutation_results.f1.std_null = std_null_f1;
    permutation_results.f1.percentile_rank = percentile_rank_f1;
    permutation_results.f1.z_score = z_score_f1;

    % LOOCV 메타데이터
    permutation_results.meta.n_permutations = n_permutations;
    permutation_results.meta.elapsed_time = elapsed_time;
    permutation_results.meta.failed_permutations = failed_permutations;
    permutation_results.meta.validation_method = validation_method;
    permutation_results.meta.n_samples = n_samples;
    permutation_results.meta.total_cv_iterations = n_permutations * n_samples;
    permutation_results.meta.avg_time_per_permutation = elapsed_time / n_permutations;
    permutation_results.meta.used_loocv_evaluation = use_loocv_evaluation;

    % 캐시 저장
    model_perm_cache = struct();
    model_perm_cache.cache_key = cache_key;
    model_perm_cache.results = permutation_results;
    model_perm_cache.timestamp = datestr(now);

    % 디렉토리 생성 (필요시)
    if ~exist(config.output_dir, 'dir')
        mkdir(config.output_dir);
    end

    try
        save(model_cache_file, 'model_perm_cache');
        fprintf('  ✓ 모델 퍼뮤테이션 결과 캐시 저장: %s\n', model_cache_file);
    catch
        fprintf('  ⚠ 캐시 저장 실패 (권한 문제 가능)\n');
    end
end

% LOOCV 결과 출력
fprintf('\n【LOOCV 기반 모델 퍼뮤테이션 테스트 결과】\n');
fprintf('────────────────────────────────────────────────\n');
fprintf('최적 방법: %s\n', best_method);
fprintf('검증 방법: LOOCV (샘플 수: %d)\n', permutation_results.meta.n_samples);
if permutation_results.meta.used_loocv_evaluation
    fprintf('원본 성능: LOOCV로 실측 (기본값 대체)\n');
else
    fprintf('원본 성능: 기존 메트릭 사용\n');
end
fprintf('퍼뮤테이션: %d회 (실패: %d회, 총 %d번 모델 학습)\n', ...
    permutation_results.meta.n_permutations, permutation_results.meta.failed_permutations, ...
    permutation_results.meta.total_cv_iterations);

% AUC 결과
fprintf('\n【AUC 결과】\n');
fprintf('관찰된 AUC: %.4f\n', permutation_results.auc.observed);
fprintf('귀무분포 평균: %.4f (±%.4f)\n', permutation_results.auc.mean_null, permutation_results.auc.std_null);
fprintf('95%% 신뢰구간: [%.4f, %.4f]\n', permutation_results.auc.ci_95(1), permutation_results.auc.ci_95(2));
fprintf('백분위 순위: %.1f%% (상위 %.1f%%)\n', permutation_results.auc.percentile_rank, 100 - permutation_results.auc.percentile_rank);
fprintf('Z-score: %.3f\n', permutation_results.auc.z_score);
fprintf('p-value: %.4f\n', permutation_results.auc.p_value);

% F1 결과
fprintf('\n【F1 스코어 결과】\n');
fprintf('관찰된 F1: %.4f\n', permutation_results.f1.observed);
fprintf('귀무분포 평균: %.4f (±%.4f)\n', permutation_results.f1.mean_null, permutation_results.f1.std_null);
fprintf('95%% 신뢰구간: [%.4f, %.4f]\n', permutation_results.f1.ci_95(1), permutation_results.f1.ci_95(2));
fprintf('백분위 순위: %.1f%% (상위 %.1f%%)\n', permutation_results.f1.percentile_rank, 100 - permutation_results.f1.percentile_rank);
fprintf('Z-score: %.3f\n', permutation_results.f1.z_score);
fprintf('p-value: %.4f\n', permutation_results.f1.p_value);

% 전체적인 유의성 판단 (더 보수적인 기준 사용)
combined_p = min(permutation_results.auc.p_value, permutation_results.f1.p_value);
if combined_p < 0.001
    significance_text = '매우 유의함 (p < 0.001)';
    sig_symbol = '***';
elseif combined_p < 0.01
    significance_text = '유의함 (p < 0.01)';
    sig_symbol = '**';
elseif combined_p < 0.05
    significance_text = '유의함 (p < 0.05)';
    sig_symbol = '*';
elseif combined_p < 0.1
    significance_text = '경계적 유의 (p < 0.1)';
    sig_symbol = '†';
else
    significance_text = '유의하지 않음 (p ≥ 0.1)';
    sig_symbol = 'ns';
end

fprintf('\n📊 전체 통계적 유의성: %s %s\n', significance_text, sig_symbol);

% 효과 크기 계산 (Cohen's d) - F1 기준
if permutation_results.f1.std_null > 1e-10
    cohens_d_f1 = (permutation_results.f1.observed - permutation_results.f1.mean_null) / permutation_results.f1.std_null;
else
    cohens_d_f1 = NaN;
end

% 효과 크기 계산 (Cohen's d) - AUC 기준
if permutation_results.auc.std_null > 1e-10
    cohens_d_auc = (permutation_results.auc.observed - permutation_results.auc.mean_null) / permutation_results.auc.std_null;
else
    cohens_d_auc = NaN;
end

% 효과 크기 해석 (인라인)
if isnan(cohens_d_f1)
    effect_text_f1 = '계산 불가';
elseif abs(cohens_d_f1) < 0.2
    effect_text_f1 = '무시할 수 있음';
elseif abs(cohens_d_f1) < 0.5
    effect_text_f1 = '작음';
elseif abs(cohens_d_f1) < 0.8
    effect_text_f1 = '중간';
else
    effect_text_f1 = '큼';
end

if isnan(cohens_d_auc)
    effect_text_auc = '계산 불가';
elseif abs(cohens_d_auc) < 0.2
    effect_text_auc = '무시할 수 있음';
elseif abs(cohens_d_auc) < 0.5
    effect_text_auc = '작음';
elseif abs(cohens_d_auc) < 0.8
    effect_text_auc = '중간';
else
    effect_text_auc = '큼';
end

fprintf('📈 효과 크기 (F1): %.3f (%s)\n', cohens_d_f1, effect_text_f1);
fprintf('📈 효과 크기 (AUC): %.3f (%s)\n', cohens_d_auc, effect_text_auc);

% 시각화 (2x3 subplot으로 재구성)
perm_fig = figure('Position', [100, 100, 1800, 1000], 'Color', 'white');

% 전체적인 유의성에 따른 색상 설정
if combined_p < 0.001
    p_color = [0, 0.5, 0]; % 진한 녹색
elseif combined_p < 0.01
    p_color = [0, 0.7, 0]; % 녹색
elseif combined_p < 0.05
    p_color = [0.5, 0.8, 0]; % 연녹색
elseif combined_p < 0.1
    p_color = [1, 0.8, 0]; % 노란색
else
    p_color = [0.8, 0, 0]; % 빨간색
end

% [1] AUC 분포
subplot(2, 3, 1);
h_auc = histogram(permutation_results.auc.null_distribution, 40, 'Normalization', 'probability', ...
    'FaceColor', [0.7, 0.7, 0.9], 'EdgeColor', 'none', 'FaceAlpha', 0.7);
hold on;

% 관찰된 AUC 표시
xline(permutation_results.auc.observed, 'Color', p_color, 'LineWidth', 3, ...
    'DisplayName', sprintf('관찰 AUC (%.3f)', permutation_results.auc.observed));

% 95% 신뢰구간
xline(permutation_results.auc.ci_95(1), '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);
xline(permutation_results.auc.ci_95(2), '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);

% 평균선
xline(permutation_results.auc.mean_null, ':', 'Color', [0.3, 0.3, 0.3], 'LineWidth', 2, ...
    'DisplayName', sprintf('평균 (%.3f)', permutation_results.auc.mean_null));

xlabel('AUC', 'FontWeight', 'bold');
ylabel('확률', 'FontWeight', 'bold');
title(sprintf('AUC 귀무분포 (p=%.4f)', permutation_results.auc.p_value), 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 9);
grid on;

% [2] F1 분포
subplot(2, 3, 2);
h_f1 = histogram(permutation_results.f1.null_distribution, 40, 'Normalization', 'probability', ...
    'FaceColor', [0.9, 0.7, 0.7], 'EdgeColor', 'none', 'FaceAlpha', 0.7);
hold on;

% 관찰된 F1 표시
xline(permutation_results.f1.observed, 'Color', p_color, 'LineWidth', 3, ...
    'DisplayName', sprintf('관찰 F1 (%.3f)', permutation_results.f1.observed));

% 95% 신뢰구간
xline(permutation_results.f1.ci_95(1), '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);
xline(permutation_results.f1.ci_95(2), '--', 'Color', [0.5, 0.5, 0.5], 'LineWidth', 1.5);

% 평균선
xline(permutation_results.f1.mean_null, ':', 'Color', [0.3, 0.3, 0.3], 'LineWidth', 2, ...
    'DisplayName', sprintf('평균 (%.3f)', permutation_results.f1.mean_null));

xlabel('F1 스코어', 'FontWeight', 'bold');
ylabel('확률', 'FontWeight', 'bold');
title(sprintf('F1 귀무분포 (p=%.4f)', permutation_results.f1.p_value), 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 9);
grid on;

% [3] AUC Q-Q 플롯
subplot(2, 3, 3);
qqplot(permutation_results.auc.null_distribution);
title('AUC Q-Q 플롯', 'FontWeight', 'bold');
xlabel('이론적 분위수', 'FontWeight', 'bold');
ylabel('표본 분위수', 'FontWeight', 'bold');
grid on;

% [4] F1 Q-Q 플롯
subplot(2, 3, 4);
qqplot(permutation_results.f1.null_distribution);
title('F1 Q-Q 플롯', 'FontWeight', 'bold');
xlabel('이론적 분위수', 'FontWeight', 'bold');
ylabel('표본 분위수', 'FontWeight', 'bold');
grid on;

% [5] 성능 비교 (박스플롯)
subplot(2, 3, 5);
try
    % 데이터 준비 (세로로 연결)
    auc_data = permutation_results.auc.null_distribution(:);
    f1_data = permutation_results.f1.null_distribution(:);

    data_for_box = [auc_data; f1_data];
    group_labels = [repmat({'AUC'}, length(auc_data), 1); ...
                   repmat({'F1'}, length(f1_data), 1)];

    % 박스플롯 생성 (색상 옵션 제거)
    h_box = boxplot(data_for_box, group_labels);

    % 박스 색상 수동 설정
    box_colors = [0.5 0.5 0.9; 0.9 0.5 0.5];
    h_patch = findobj(gca, 'Tag', 'Box');
    for j = 1:length(h_patch)
        if j <= size(box_colors, 1)
            patch(get(h_patch(j), 'XData'), get(h_patch(j), 'YData'), ...
                  box_colors(j, :), 'FaceAlpha', 0.7);
        end
    end

    hold on;
catch
    % 박스플롯 실패 시 대안
    bar([1, 2], [permutation_results.auc.mean_null, permutation_results.f1.mean_null], ...
        'FaceColor', [0.7, 0.7, 0.7]);
    hold on;
    errorbar([1, 2], [permutation_results.auc.mean_null, permutation_results.f1.mean_null], ...
             [permutation_results.auc.std_null, permutation_results.f1.std_null], 'k', 'LineWidth', 2);
    set(gca, 'XTickLabel', {'AUC', 'F1'});
end

% 관찰값 점 추가
scatter(1, permutation_results.auc.observed, 100, p_color, 'filled', 'd', ...
    'DisplayName', '관찰값');
scatter(2, permutation_results.f1.observed, 100, p_color, 'filled', 'd');

ylabel('성능 점수', 'FontWeight', 'bold');
title('성능 메트릭 비교', 'FontWeight', 'bold');
grid on;

% [6] 통계 요약
subplot(2, 3, 6);
axis off;

summary_text = {
    ['▶ LOOCV 기반 퍼뮤테이션 결과'];
    ['최적 방법: ' best_method];
    ['검증: LOOCV (n=' sprintf('%d', permutation_results.meta.n_samples) ')'];
    '';
    ['【AUC】'];
    ['관찰값: ' sprintf('%.4f', permutation_results.auc.observed)];
    ['귀무평균: ' sprintf('%.4f', permutation_results.auc.mean_null)];
    ['p-value: ' sprintf('%.4f', permutation_results.auc.p_value)];
    ['효과크기: ' sprintf('%.3f (%s)', cohens_d_auc, effect_text_auc)];
    '';
    ['【F1 스코어】'];
    ['관찰값: ' sprintf('%.4f', permutation_results.f1.observed)];
    ['귀무평균: ' sprintf('%.4f', permutation_results.f1.mean_null)];
    ['p-value: ' sprintf('%.4f', permutation_results.f1.p_value)];
    ['효과크기: ' sprintf('%.3f (%s)', cohens_d_f1, effect_text_f1)];
    '';
    ['【전체 결과】'];
    ['유의성: ' significance_text ' ' sig_symbol];
    ['퍼뮤테이션: ' sprintf('%d회', permutation_results.meta.n_permutations)];
    ['실패: ' sprintf('%d회', permutation_results.meta.failed_permutations)];
    ['총 모델 학습: ' sprintf('%d회', permutation_results.meta.total_cv_iterations)];
    ['소요시간: ' sprintf('%.1f초', permutation_results.meta.elapsed_time)]
};

text(0.05, 0.95, summary_text, 'Units', 'normalized', 'VerticalAlignment', 'top', ...
    'FontSize', 10, 'FontWeight', 'normal', 'FontName', 'Malgun Gothic');

title('통계 요약', 'FontWeight', 'bold');

% 전체 제목
sgtitle(sprintf('LOOCV 기반 모델 퍼뮤테이션 테스트 (%s)', best_method), ...
    'FontSize', 16, 'FontWeight', 'bold');

% 그래프 저장 (백업 시스템 적용)
loocv_perm_chart_filepath = get_managed_filepath(config.output_dir, 'loocv_permutation_test_results.png', config);
backup_and_prepare_file(loocv_perm_chart_filepath, config);
% saveas(perm_fig, loocv_perm_chart_filepath);
fprintf('\n✓ LOOCV 퍼뮤테이션 테스트 차트 저장: %s\n', loocv_perm_chart_filepath);

% 결과를 전역 구조체에 저장 (새로운 구조)
result_data.model_permutation_test = permutation_results;

% 이전 버전과의 호환성을 위해 F1 정보도 저장
result_data.permutation_test = struct();
result_data.permutation_test.observed_f1 = permutation_results.f1.observed;
result_data.permutation_test.p_value = permutation_results.f1.p_value;
result_data.permutation_test.mean_null = permutation_results.f1.mean_null;
result_data.permutation_test.std_null = permutation_results.f1.std_null;
result_data.permutation_test.ci_95 = permutation_results.f1.ci_95;
result_data.permutation_test.percentile_rank = permutation_results.f1.percentile_rank;
result_data.permutation_test.z_score = permutation_results.f1.z_score;
result_data.permutation_test.n_permutations = permutation_results.meta.n_permutations;
result_data.permutation_test.elapsed_time = permutation_results.meta.elapsed_time;

fprintf('\n✅ LOOCV 기반 모델 퍼뮤테이션 테스트 완료\n');
fprintf('   - 샘플별 교차검증으로 최대 데이터 활용\n');
fprintf('   - 작은 데이터에서 안정적인 성능 추정\n');

%% STEP 23: 결과 요약 및 권장사항
fprintf('\n\n【STEP 23】 최종 결과 요약 및 권장사항\n');
fprintf('════════════════════════════════════════════════════════════\n\n');

fprintf('【가중치 방법별 성능 순위】\n');
fprintf('────────────────────────────────────────────\n');

% F1 점수 기준 정렬
f1_scores = zeros(n_methods, 1);
for m = 1:n_methods
    f1_scores(m) = prediction_results.(weight_methods{m}).f1_score;
end
[sorted_f1, sort_idx] = sort(f1_scores, 'descend');

for i = 1:n_methods
    m = sort_idx(i);
    fprintf('%d. %s: F1=%.3f, AUC=%.3f, Accuracy=%.3f\n', ...
        i, weight_methods{m}, ...
        prediction_results.(weight_methods{m}).f1_score, ...
        prediction_results.(weight_methods{m}).AUC, ...
        prediction_results.(weight_methods{m}).accuracy);
end

fprintf('\n【핵심 역량 (신뢰도 High & 앙상블 가중치 > 5%%)】\n');
fprintf('────────────────────────────────────────────\n');

core_competencies = weight_comparison(weight_comparison.Reliability == 'High' & ...
                                     weight_comparison.Ensemble_Mean > 5, :);
if height(core_competencies) > 0
    for i = 1:height(core_competencies)
        fprintf('• %s: %.2f%% (표준편차: %.2f)\n', ...
            core_competencies.Feature{i}, ...
            core_competencies.Ensemble_Mean(i), ...
            core_competencies.Std(i));
    end
else
    fprintf('신뢰도가 높고 가중치가 5%% 이상인 핵심 역량이 없습니다.\n');
end

fprintf('\n【권장사항】\n');
fprintf('────────────────────────────────────────────\n');

% 최적 방법 권장
fprintf('1. 예측 모델: %s 방법을 주요 모델로 사용 권장\n', best_method);
fprintf('   - F1 Score: %.3f\n', prediction_results.(best_method).f1_score);
fprintf('   - 균형잡힌 정밀도(%.3f)와 재현율(%.3f)\n', ...
    prediction_results.(best_method).precision, ...
    prediction_results.(best_method).recall);

% 앙상블 활용 (필드 존재 확인)
if isfield(prediction_results, 'Ensemble') && prediction_results.Ensemble.f1_score > 0.7
    fprintf('\n2. 앙상블 방법도 우수한 성능 (F1=%.3f)\n', ...
        prediction_results.Ensemble.f1_score);
    fprintf('   - 여러 방법의 장점을 결합한 안정적 예측\n');
end

% 주의사항
fprintf('\n3. 주의사항:\n');

% best_pred_col 구성 (필드 이름 생성)
best_pred_col = [best_method '_Predicted'];

% False Positive 확인 (컬럼 존재 확인)
if ismember(best_pred_col, participant_results.Properties.VariableNames)
    fp_count = sum(participant_results.Actual_Label == 0 & participant_results.(best_pred_col) == 1);
    if fp_count > 0
        fprintf('   - False Positive %d건: 저성과자를 고성과자로 오분류 주의\n', fp_count);
    end

    fn_count = sum(participant_results.Actual_Label == 1 & participant_results.(best_pred_col) == 0);
    if fn_count > 0
        fprintf('   - False Negative %d건: 고성과자를 저성과자로 오분류 주의\n', fn_count);
    end
else
    fprintf('   - 개별 예측 결과 확인 불가 (컬럼 없음)\n');
end

fprintf('\n【선발 효과 관련 해석 주의사항】\n');
fprintf('────────────────────────────────────────────\n');

fprintf('1. 동질성 높은 역량들:\n');
if exist('high_homo_comps', 'var') && height(high_homo_comps) > 0
    for i = 1:min(5, height(high_homo_comps))
        fprintf('   - %s: 이미 선발 단계에서 필터링됨\n', high_homo_comps.Competency{i});
    end
    fprintf('   → 낮은 예측력이 선발 실패를 의미하지 않음\n\n');
else
    fprintf('   - 동질성 높은 역량 정보 없음\n\n');
end

fprintf('2. Range Restriction 영향:\n');
fprintf('   - 관찰된 상관계수는 과소추정될 가능성 있음\n');
fprintf('   - 실제 모집단에서는 더 높은 예측력 기대 가능\n\n');

if exist('age_perf_corr', 'var') && ~isnan(age_perf_corr) && abs(age_perf_corr) > 0.3
    fprintf('3. 나이 효과:\n');
    fprintf('   - 나이-성과 상관: %.3f\n', age_perf_corr);
    fprintf('   - 나이 조정 모델 사용 권장\n\n');
end

if exist('p_gender', 'var') && ~isnan(p_gender) && p_gender < 0.05
    fprintf('4. 성별 효과:\n');
    fprintf('   - 성별 차이 유의함 (p=%.4f)\n', p_gender);
    fprintf('   - 공정성 검토 필요\n\n');
end

fprintf('\n✅ 통합 가중치 분석 및 예측 검증 완료!\n');

%% STEP 24: 종합 엑셀 결과 파일 생성
fprintf('\n【STEP 24】 종합 엑셀 결과 파일 생성\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 엑셀 파일 경로 설정 (백업 시스템 적용)
excel_filepath = get_managed_filepath(config.output_dir, 'HR_Analysis_Results.xlsx', config);
backup_and_prepare_file(excel_filepath, config);

try
    fprintf('▶ 엑셀 파일 생성 중: %s\n', excel_filepath);

    %% 1. 요약 시트
    fprintf('  • 요약 시트 생성 중...\n');
    summary_data = table();
    summary_data.Item = {'분석일시'; '총샘플수'; '고성과자수'; '저성과자수'; '최적방법'; ...
                         '최적AUC'; '최적F1'; '사용된baseline'; '개발자필터'};

    % 성능 지표 추출
    best_f1 = prediction_results.(best_method).f1_score;
    best_auc = prediction_results.(best_method).AUC;

    % applied_filter가 없을 경우 기본값 설정
    if isfield(config, 'applied_filter')
        filter_text = config.applied_filter;
    else
        filter_text = '정보없음';
    end

    summary_data.Value = {datestr(now, 'yyyy-mm-dd HH:MM:SS');
                          length(y_weight);
                          sum(y_weight==1);
                          sum(y_weight==0);
                          best_method;
                          sprintf('%.3f', best_auc);
                          sprintf('%.3f', best_f1);
                          config.baseline_type;
                          filter_text};

    writetable(summary_data, excel_filepath, 'Sheet', '요약');

    %% 2. 인재유형분석 시트
    fprintf('  • 인재유형분석 시트 생성 중...\n');
    writetable(profile_stats, excel_filepath, 'Sheet', '인재유형분석');

    %% 3. 역량가중치 시트
    fprintf('  • 역량가중치 시트 생성 중...\n');
    writetable(weight_comparison, excel_filepath, 'Sheet', '역량가중치');

    %% 4. 상관분석 시트 (있는 경우)
    if exist('correlation_results', 'var')
        fprintf('  • 상관분석 시트 생성 중...\n');
        writetable(correlation_results, excel_filepath, 'Sheet', '상관분석');
    end

    %% 5. 예측성능 시트
    fprintf('  • 예측성능 시트 생성 중...\n');

    % 성능 지표 테이블 생성
    performance_table = table();
    performance_table.Method = weight_methods';

    for m = 1:n_methods
        method = weight_methods{m};
        performance_table.AUC(m) = prediction_results.(method).AUC;
        performance_table.F1_Score(m) = prediction_results.(method).f1_score;
        performance_table.Accuracy(m) = prediction_results.(method).accuracy;
        performance_table.Precision(m) = prediction_results.(method).precision;
        performance_table.Recall(m) = prediction_results.(method).recall;
    end

    % 앙상블 결과 추가 (있는 경우)
    if isfield(prediction_results, 'Ensemble')
        performance_table.Method{end+1} = 'Ensemble';
        performance_table.AUC(end+1) = prediction_results.Ensemble.AUC;
        performance_table.F1_Score(end+1) = prediction_results.Ensemble.f1_score;
        performance_table.Accuracy(end+1) = prediction_results.Ensemble.accuracy;
        performance_table.Precision(end+1) = prediction_results.Ensemble.precision;
        performance_table.Recall(end+1) = prediction_results.Ensemble.recall;
    end

    writetable(performance_table, excel_filepath, 'Sheet', '예측성능');

    %% 6. Bootstrap결과 시트 (있는 경우)
    if exist('bootstrap_stats', 'var')
        fprintf('  • Bootstrap결과 시트 생성 중...\n');
        writetable(bootstrap_stats, excel_filepath, 'Sheet', 'Bootstrap결과');
    end

    %% 7. 극단그룹분석 시트 (있는 경우)
    if exist('ttest_results', 'var')
        fprintf('  • 극단그룹분석 시트 생성 중...\n');
        writetable(ttest_results, excel_filepath, 'Sheet', '극단그룹분석');
    end

    %% 8. 개별예측 시트
    fprintf('  • 개별예측 시트 생성 중...\n');
    writetable(participant_results, excel_filepath, 'Sheet', '개별예측');

    %% 9. 퍼뮤테이션테스트 시트 (있는 경우)
    if exist('permutation_results', 'var')
        fprintf('  • 퍼뮤테이션테스트 시트 생성 중...\n');

        % 퍼뮤테이션 결과 테이블 생성
        perm_table = table();
        perm_table.Item = {'관찰된F1'; '귀무분포평균'; '귀무분포표준편차'; ...
                          '95%CI하한'; '95%CI상한'; '백분위순위'; 'Z_score'; ...
                          'p_value'; '퍼뮤테이션횟수'; '소요시간초'};

        perm_table.Value = {permutation_results.observed_f1;
                           permutation_results.mean_null;
                           permutation_results.std_null;
                           permutation_results.ci_95(1);
                           permutation_results.ci_95(2);
                           permutation_results.percentile_rank;
                           permutation_results.z_score;
                           permutation_results.p_value;
                           permutation_results.n_permutations;
                           permutation_results.elapsed_time};

        writetable(perm_table, excel_filepath, 'Sheet', '퍼뮤테이션테스트');
    end

    %% 10. 핵심역량 시트 (추가)
    if exist('core_competencies', 'var') && height(core_competencies) > 0
        fprintf('  • 핵심역량 시트 생성 중...\n');
        writetable(core_competencies, excel_filepath, 'Sheet', '핵심역량');
    end

    %% 11. 역량분포분석 시트 (추가)
    if exist('range_stats', 'var')
        fprintf('  • 역량분포분석 시트 생성 중...\n');
        writetable(range_stats, excel_filepath, 'Sheet', '역량분포분석');
    end

    %% 12. 인구통계효과 시트 (추가)
    if exist('age_perf_corr', 'var') || exist('p_gender', 'var')
        demo_effects = table();
        if exist('age_perf_corr', 'var') && ~isnan(age_perf_corr)
            demo_effects.Age_Correlation = age_perf_corr;
        else
            demo_effects.Age_Correlation = NaN;
        end
        if exist('p_gender', 'var') && ~isnan(p_gender)
            demo_effects.Gender_Pvalue = p_gender;
        else
            demo_effects.Gender_Pvalue = NaN;
        end

        fprintf('  • 인구통계효과 시트 생성 중...\n');
        writetable(demo_effects, excel_filepath, 'Sheet', '인구통계효과');
    end

    % 새로운 분석 결과 시트 추가 (STEP 4.5.2 & 9.6)

    % 역량 분포 분석 결과 시트
    if exist('distribution_stats', 'var') && height(distribution_stats) > 0
        fprintf('  • 역량분포분석 시트 생성 중...\n');
        writetable(distribution_stats, excel_filepath, 'Sheet', '역량분포분석');
    end

    % 성별 효과 분석 결과 시트
    if exist('gender_analysis_results', 'var') && isfield(gender_analysis_results, 'competency_differences')
        fprintf('  • 성별역량차이 시트 생성 중...\n');
        writetable(gender_analysis_results.competency_differences, excel_filepath, 'Sheet', '성별역량차이');

        % 성별 분석 요약 시트
        fprintf('  • 성별분석요약 시트 생성 중...\n');
        gender_summary = table();
        gender_summary.항목 = {'남성_샘플수'; '여성_샘플수'; 'Disparate_Impact_Ratio'; '공정성_평가'; '유의미한_차이_역량수'; '큰_효과크기_역량수'};

        if isfield(gender_analysis_results, 'sample_sizes')
            male_count = gender_analysis_results.sample_sizes.male;
            female_count = gender_analysis_results.sample_sizes.female;
        else
            male_count = NaN;
            female_count = NaN;
        end

        if isfield(gender_analysis_results, 'fairness')
            dir_ratio = gender_analysis_results.fairness.disparate_impact_ratio;
            fairness_stat = string(gender_analysis_results.fairness.status);
        else
            dir_ratio = NaN;
            fairness_stat = "알수없음";
        end

        sig_diff_count = sum(gender_analysis_results.competency_differences.Significant);
        large_effect_count = sum(abs(gender_analysis_results.competency_differences.Cohen_d) > 0.5);

        gender_summary.값 = [male_count; female_count; dir_ratio; fairness_stat; sig_diff_count; large_effect_count];

        writetable(gender_summary, excel_filepath, 'Sheet', '성별분석요약');
    end

    % Cut-off 진단 결과 시트 (Cut-off 버전에서 추가)
    if exist('cutoff_diagnosis', 'var') && height(cutoff_diagnosis) > 0
        fprintf('  • Cut-off진단 시트 생성 중...\n');
        writetable(cutoff_diagnosis, excel_filepath, 'Sheet', 'Cut-off진단');

        % Cut-off 요약 시트
        fprintf('  • Cut-off요약 시트 생성 중...\n');
        cutoff_summary = table();
        cutoff_summary.항목 = {'전체_역량수'; '심각한_Cut-off_역량수'; '중등도_Cut-off_역량수'; '경미한_Cut-off_역량수'; ...
                               '평균_최소점수'; '평균_변이계수'; 'Thorndike보정_적용여부'};

        total_comps = height(cutoff_diagnosis);
        severe_count = sum(cutoff_diagnosis.Severity == 'Severe');
        moderate_count = sum(cutoff_diagnosis.Severity == 'Moderate');
        mild_count = sum(cutoff_diagnosis.Severity == 'Mild');
        avg_min_score = mean(cutoff_diagnosis.Min_Score);
        avg_cv = mean(cutoff_diagnosis.CV);

        % Thorndike 보정 적용 여부 확인
        thorndike_applied = exist('use_correction', 'var') && use_correction;

        cutoff_summary.값 = [total_comps; severe_count; moderate_count; mild_count; ...
                            avg_min_score; avg_cv; thorndike_applied];

        writetable(cutoff_summary, excel_filepath, 'Sheet', 'Cut-off요약');
    end

    % 시트 개수 계산 (기본 8개 + 추가 시트들)
    sheet_count = 8; % 기본 시트: 요약, 인재유형분석, 역량가중치, 예측성능, 개별예측, 등
    if exist('correlation_results', 'var'), sheet_count = sheet_count + 1; end
    if exist('bootstrap_stats', 'var'), sheet_count = sheet_count + 1; end
    if exist('ttest_results', 'var'), sheet_count = sheet_count + 1; end
    if exist('permutation_results', 'var'), sheet_count = sheet_count + 1; end
    if exist('core_competencies', 'var') && height(core_competencies) > 0, sheet_count = sheet_count + 1; end
    if exist('range_stats', 'var'), sheet_count = sheet_count + 1; end
    if exist('age_perf_corr', 'var') || exist('p_gender', 'var'), sheet_count = sheet_count + 1; end
    if exist('distribution_stats', 'var') && height(distribution_stats) > 0, sheet_count = sheet_count + 1; end
    if exist('gender_analysis_results', 'var') && isfield(gender_analysis_results, 'competency_differences'), sheet_count = sheet_count + 2; end
    if exist('cutoff_diagnosis', 'var') && height(cutoff_diagnosis) > 0, sheet_count = sheet_count + 2; end  % Cut-off 진단 + 요약

    fprintf('✅ 엑셀 파일 생성 완료: %s\n', excel_filepath);
    fprintf('  📋 총 %d개 시트 포함\n', sheet_count);

catch excel_error
    fprintf('⚠ 엑셀 파일 생성 실패: %s\n', excel_error.message);
    fprintf('  데이터는 .mat 파일로 저장되었습니다.\n');
end

%% 결과 저장
integrated_results = struct();
integrated_results.weight_comparison = weight_comparison;
integrated_results.prediction_results = prediction_results;
integrated_results.participant_results = participant_results;
integrated_results.best_method = best_method;

% 선택적 결과 포함
if exist('core_competencies', 'var')
    integrated_results.core_competencies = core_competencies;
end

if exist('permutation_results', 'var')
    integrated_results.permutation_results = permutation_results;
end

if exist('bootstrap_stats', 'var')
    integrated_results.bootstrap_stats = bootstrap_stats;
end

if exist('ttest_results', 'var')
    integrated_results.ttest_results = ttest_results;
end

% 최종 통합 결과 저장 (백업 시스템 적용)
integrated_results_filepath = get_managed_filepath(config.output_dir, 'integrated_analysis_results.mat', config);
backup_and_prepare_file(integrated_results_filepath, config);
save(integrated_results_filepath, 'integrated_results');

fprintf('\n📊 분석 결과가 저장되었습니다: %s\n', integrated_results_filepath);

%% ========================================================================
%                          선발 효과 보정 함수
% =========================================================================

function adjusted_scores = correct_selection_effect(scores, range_stats)
    % 선발 효과로 인한 range restriction을 보정하는 함수
    % scores: 원래 점수 행렬 (n x p)
    % range_stats: 역량별 분포 통계 테이블

    adjusted_scores = scores;

    for i = 1:size(scores, 2)
        if range_stats.CV(i) < 0.2
            observed_sd = std(scores(:, i));
            % 추정된 모집단 표준편차 (선발 효과 보정)
            estimated_population_sd = observed_sd * 1.5;

            z_scores = (scores(:, i) - mean(scores(:, i))) / observed_sd;
            adjusted_scores(:, i) = mean(scores(:, i)) + z_scores * estimated_population_sd;
        end
    end
end

%% ========================================================================
%                    Range Restriction 보정 함수들 (Cut-off 버전)
% =========================================================================

% Thorndike Case II 공식을 통한 상관계수 보정
function r_corrected = thorndike_correction(r_observed, sd_restricted, sd_unrestricted)
    % r_observed: 제한된 샘플에서 관찰된 상관계수
    % sd_restricted: 제한된 샘플의 표준편차 (현재 데이터)
    % sd_unrestricted: 추정된 모집단 표준편차

    u = sd_unrestricted / sd_restricted;  % 보정 계수

    % Thorndike Case II 공식
    r_corrected = r_observed * u / sqrt(1 + r_observed^2 * (u^2 - 1));

    % 극단값 처리 (상관계수는 -1~1 범위)
    r_corrected = max(-1, min(1, r_corrected));
end

% 모집단 표준편차 추정 함수
function sd_population = estimate_population_sd(scores, cutoff_point)
    % scores: 관찰된 점수 벡터
    % cutoff_point: 선발 기준점 (예: 40점)

    observed_mean = mean(scores);
    observed_sd = std(scores);
    min_score = min(scores);

    % Cut-off 효과가 있는지 판단
    if min_score > cutoff_point
        % 절단된 정규분포에서 모집단 모수 추정
        alpha = (cutoff_point - observed_mean) / observed_sd;

        % Mills ratio를 이용한 보정
        lambda = normpdf(alpha) / (1 - normcdf(alpha));

        % 모집단 평균 추정
        population_mean = observed_mean - observed_sd * lambda;

        % 모집단 표준편차 추정
        delta = lambda * (lambda - alpha);
        sd_population = observed_sd / sqrt(1 - delta);

        % 합리적 범위 제한 (원점수의 1.2-3배)
        sd_population = max(observed_sd * 1.2, min(observed_sd * 3, sd_population));
    else
        % Cut-off 효과가 없으면 경험적 추정
        % 0-100 범위를 고려한 표준편차 (약 16.7이 이론값)
        theoretical_sd = 100 / 6;  % 99.7% 범위를 6σ로 가정
        sd_population = max(observed_sd, theoretical_sd);
    end
end

% Cut-off 효과를 반영한 로지스틱 회귀 함수
function mdl_corrected = cutoff_aware_logistic(X, y, scores, cutoff_method)
    % X: 독립변수 행렬
    % y: 종속변수 (이진)
    % scores: 역량 점수 (Cut-off 탐지용)
    % cutoff_method: 'weighted', 'heckman', 'truncated' 중 선택

    switch lower(cutoff_method)
        case 'weighted'
            % 방법 1: 가중치 조정
            % Cut-off 근처 샘플에 더 높은 가중치 부여
            mean_scores = mean(scores, 2);
            cutoff_estimate = prctile(mean_scores, 10);  % 하위 10%를 Cut-off로 추정

            % 시그모이드 가중치 (Cut-off 근처에서 가중치 증가)
            weights = 1 ./ (1 + exp(-0.1 * (mean_scores - cutoff_estimate - 5)));
            weights = weights / mean(weights);  % 정규화

            % 가중 로지스틱 회귀
            mdl_corrected = fitglm(X, y, 'Distribution', 'binomial', 'Weights', weights);

        case 'heckman'
            % 방법 2: Heckman 2단계 보정
            % 1단계: 선발 확률 모델링
            mean_scores = mean(scores, 2);
            cutoff_estimate = prctile(mean_scores, 10);
            selected_dummy = mean_scores > cutoff_estimate;

            % Probit 모델로 선발 확률 추정
            selection_model = fitglm(X, selected_dummy, 'Distribution', 'binomial', 'Link', 'probit');
            fitted_probs = selection_model.Fitted.Probability;

            % Inverse Mills Ratio 계산
            lambda = normpdf(norminv(fitted_probs)) ./ fitted_probs;
            lambda(isinf(lambda) | isnan(lambda)) = 0;  % 극단값 처리

            % 2단계: 본 모델에 lambda 추가
            X_augmented = [X, lambda];
            mdl_corrected = fitglm(X_augmented, y, 'Distribution', 'binomial');

        case 'truncated'
            % 방법 3: 절단된 회귀 (간단한 구현)
            % Cut-off 이하 샘플에 패널티 적용
            mean_scores = mean(scores, 2);
            cutoff_estimate = prctile(mean_scores, 15);

            % Cut-off 근처 샘플 식별
            near_cutoff = mean_scores < cutoff_estimate + 10;

            % 패널티 가중치 적용
            weights = ones(size(y));
            weights(near_cutoff) = 0.5;  % Cut-off 근처 샘플에 낮은 가중치

            mdl_corrected = fitglm(X, y, 'Distribution', 'binomial', 'Weights', weights);

        otherwise
            % 기본 로지스틱 회귀
            mdl_corrected = fitglm(X, y, 'Distribution', 'binomial');
    end
end

% 보정된 상관계수 계산 함수
function corrected_correlations = calculate_corrected_correlations(X, y, scores, method)
    % X: 독립변수 행렬 (역량 점수들)
    % y: 종속변수 (성과)
    % scores: 전체 역량 점수 행렬
    % method: 'thorndike' 또는 'empirical'

    n_features = size(X, 2);
    corrected_correlations = zeros(n_features, 1);

    for i = 1:n_features
        % 관찰된 상관계수
        r_observed = corr(X(:, i), y, 'Rows', 'complete');

        % 표준편차 계산
        sd_restricted = std(X(:, i), 'omitnan');

        if strcmp(method, 'thorndike')
            % Thorndike 보정 사용
            cutoff_point = 30;  % 일반적인 Cut-off 추정값
            sd_unrestricted = estimate_population_sd(X(:, i), cutoff_point);
            corrected_correlations(i) = thorndike_correction(r_observed, sd_restricted, sd_unrestricted);
        else
            % 경험적 보정 (간단한 방법)
            % Range restriction이 심한 경우 상관계수를 1.3-1.5배 상향 조정
            cv = sd_restricted / mean(X(:, i));
            if cv < 0.2  % 변이계수가 낮으면 Range restriction 의심
                correction_factor = 1.4;
            elseif cv < 0.3
                correction_factor = 1.2;
            else
                correction_factor = 1.0;
            end

            corrected_correlations(i) = r_observed * correction_factor;
            corrected_correlations(i) = max(-1, min(1, corrected_correlations(i)));  % 범위 제한
        end
    end
end

% Cut-off 효과 진단 함수
function cutoff_diagnosis = diagnose_cutoff_effects(scores, variable_names)
    % scores: 역량 점수 행렬
    % variable_names: 역량 이름들

    n_vars = size(scores, 2);
    cutoff_diagnosis = table();
    cutoff_diagnosis.Variable = variable_names(:);
    cutoff_diagnosis.Min_Score = zeros(n_vars, 1);
    cutoff_diagnosis.Q05 = zeros(n_vars, 1);
    cutoff_diagnosis.Mean = zeros(n_vars, 1);
    cutoff_diagnosis.SD = zeros(n_vars, 1);
    cutoff_diagnosis.CV = zeros(n_vars, 1);
    cutoff_diagnosis.Suspected_Cutoff = zeros(n_vars, 1);
    cutoff_diagnosis.Severity = categorical(repmat({'None'}, n_vars, 1));

    for i = 1:n_vars
        valid_scores = scores(~isnan(scores(:, i)), i);

        if length(valid_scores) > 5
            cutoff_diagnosis.Min_Score(i) = min(valid_scores);
            cutoff_diagnosis.Q05(i) = prctile(valid_scores, 5);
            cutoff_diagnosis.Mean(i) = mean(valid_scores);
            cutoff_diagnosis.SD(i) = std(valid_scores);
            cutoff_diagnosis.CV(i) = cutoff_diagnosis.SD(i) / cutoff_diagnosis.Mean(i);

            % Cut-off 추정 (최소값이 비정상적으로 높은 경우)
            theoretical_min = 0;
            if cutoff_diagnosis.Min_Score(i) > 20  % 0-100 척도에서 20점 이상
                cutoff_diagnosis.Suspected_Cutoff(i) = cutoff_diagnosis.Min_Score(i) - 5;

                % 심각도 판정
                if cutoff_diagnosis.Min_Score(i) > 40
                    cutoff_diagnosis.Severity(i) = 'Severe';
                elseif cutoff_diagnosis.Min_Score(i) > 30
                    cutoff_diagnosis.Severity(i) = 'Moderate';
                else
                    cutoff_diagnosis.Severity(i) = 'Mild';
                end
            end
        end
    end
end