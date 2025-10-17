% =======================================================================
%            역량검사 가중치 적용 점수 추출 (NaN-safe, 100점 환산)
% =======================================================================
% 목적
%  - cost_sensitive_weights.mat 의 (feature_names, final_weights)를 사용하여
%    역량 데이터의 가중 합산 점수를 계산 (NaN 안전)
%  - 결과를 0~100 점수로 선형 환산(100점 단위)
%  - 참조 엑셀의 '역량검사_종합점수' 형식을 최대한 따름
%
% 주요 특징
%  - NaN-안전 정규화(열 평균/표준편차는 omitnan)
%  - 개인별로 관측된 피처만 가중합 → 행별 가중치 재정규화
%  - 최종 점수 0~100 스케일, 정수 반올림
%  - 가중치 상세/메타데이터 시트 동시 저장
%  - 기존 파일 자동 백업
%
% 작성일: 2025-09-23
% =======================================================================

clear; clc; close all;

% ---- 전역 폰트(선택) ---------------------------------------------------
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);

fprintf('=============================================\n');
fprintf('   역량검사 가중치 적용 점수 추출 (NaN-safe)\n');
fprintf('=============================================\n\n');

%% 1) 설정 ----------------------------------------------------------------
fprintf('【STEP 1】 설정\n');
fprintf('--------------------------------------------\n');

config = struct();
% 개발자 전용 가중치 모델 사용
config.weight_model_dir    = 'D:\project\HR데이터\결과\자가불소_developer';  % 개발자 전용 가중치 디렉토리
config.new_output_dir      = 'D:\project\HR데이터\결과\자가불소_developer';  % 개발자 전용 출력 디렉토리
config.reference_file      = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가.xlsx';
config.hr_file             = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file           = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가.xlsx';
config.comp_sheet          = '역량검사_상위항목';     % 역량 점수/지표가 들어있는 시트
config.id_col_hr           = 'ID';
config.id_col_comp         = 'ID';                   % 없으면 1열 사용
config.site_col_candidates = {'사이트','Site','site','태넌트','Var2'}; % 사이트/조직 컬럼 후보
config.timestamp           = datestr(now, 'yyyy-mm-dd_HHMMSS');
config.output_filename     = sprintf('역량검사_가중치적용점수_개발자_%s.xlsx', config.timestamp);

if ~exist(config.new_output_dir, 'dir')
    mkdir(config.new_output_dir);
    fprintf('  ✓ 출력 디렉토리 생성: %s\n', config.new_output_dir);
else
    fprintf('  ✓ 출력 디렉토리 확인: %s\n', config.new_output_dir);
end

%% 2) 가중치 파일 로드 -----------------------------------------------------
fprintf('\n【STEP 2】 가중치 파일 로드\n');
fprintf('--------------------------------------------\n');

weight_file = fullfile(config.weight_model_dir, 'cost_sensitive_weights.mat');
assert(exist(weight_file,'file')==2, '가중치 파일을 찾을 수 없습니다: %s', weight_file);
load(weight_file); % result_data 기대

fprintf('  ✓ 가중치 파일 로드: %s\n', weight_file);

% 유효성 검사
assert(exist('result_data','var')==1, '가중치 파일에 result_data 변수가 없습니다.');
fld_ok = isfield(result_data, {'final_weights','feature_names'});
assert(all(fld_ok), 'result_data에 final_weights / feature_names가 없습니다.');

final_weights = result_data.final_weights(:);    % px1
feature_names = string(result_data.feature_names(:)); % px1 string
assert(numel(final_weights)==numel(feature_names), '가중치와 피처명이 불일치합니다.');

fprintf('  ✓ 역량 수: %d개\n', numel(feature_names));
fprintf('    - 가중치 범위: %.2f%% ~ %.2f%%\n', min(final_weights), max(final_weights));
fprintf('    - 역량명 예시: %s\n', strjoin(feature_names(1:min(10,end))', ', '));

%% 3) 원본 데이터 로드 ------------------------------------------------------
fprintf('\n【STEP 3】 원본 데이터 로드\n');
fprintf('--------------------------------------------\n');

% HR
hr_data = readtable(config.hr_file, 'VariableNamingRule','preserve');
fprintf('  ✓ HR 데이터: %d행 x %d열\n', height(hr_data), width(hr_data));
assert(ismember(config.id_col_hr, hr_data.Properties.VariableNames), ...
       'HR 데이터에 ID 컬럼(%s)이 없습니다.', config.id_col_hr);

% 역량
comp_data = readtable(config.comp_file, 'Sheet', config.comp_sheet, ...
                      'VariableNamingRule','preserve');
fprintf('  ✓ 역량 데이터(%s): %d행 x %d열\n', config.comp_sheet, height(comp_data), width(comp_data));

if ~ismember(config.id_col_comp, comp_data.Properties.VariableNames)
    config.id_col_comp = comp_data.Properties.VariableNames{1};
    fprintf('  ⚠ 지정한 ID 컬럼 없음 → 1열(%s)을 ID로 사용\n', config.id_col_comp);
end

%% 4) ID 매칭 --------------------------------------------------------------
fprintf('\n【STEP 4】 ID 매칭\n');
fprintf('--------------------------------------------\n');

hr_ids   = hr_data.(config.id_col_hr);
comp_ids = comp_data.(config.id_col_comp);
[common_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_ids);

assert(~isempty(common_ids), '공통 ID가 없습니다. 입력 데이터를 확인하세요.');
hr_matched   = hr_data(hr_idx, :);
comp_matched = comp_data(comp_idx, :);

fprintf('  ✓ 공통 ID: %d개 (HR:%d, COMP:%d)\n', numel(common_ids), numel(hr_ids), numel(comp_ids));

%% 5) 가중치-컬럼 매칭 & 행렬 구성 -----------------------------------------
fprintf('\n【STEP 5】 역량 컬럼 매칭\n');
fprintf('--------------------------------------------\n');

p = numel(feature_names);
matched_mask = false(p,1);
X = [];                         % nxp
kept_features = strings(0);
kept_weights  = [];

for i = 1:p
    fn = char(feature_names(i));
    if ismember(fn, comp_matched.Properties.VariableNames)
        col = comp_matched.(fn);
        if ~isnumeric(col)
            % 숫자형으로 변환 시도
            try
                col = double(col);
            catch
                col = nan(height(comp_matched),1);
            end
        end
        X = [X, col(:)]; %#ok<AGROW>
        kept_features(end+1,1) = string(fn); %#ok<AGROW>
        kept_weights(end+1,1)  = final_weights(i); %#ok<AGROW>
        matched_mask(i) = true;
        fprintf('  ✓ %-10s : 가중치 %.2f%%, 유효값 %d\n', fn, final_weights(i), sum(~isnan(col)));
    else
        fprintf('  ✗ %-10s : 컬럼 없음\n', fn);
    end
end

assert(~isempty(X), '매칭된 역량 컬럼이 없습니다.');
n = size(X,1);
fprintf('  → 최종 매칭: %d/%d개 역량 사용\n', numel(kept_features), p);

%% 6) NaN-safe 가중 점수 계산 (Z-정규화 제거, 원점수 가중합) ---------------
fprintf('\n【STEP 6】 점수 계산 (NaN-safe, 원점수 가중합)\n');
fprintf('--------------------------------------------\n');

% (1) 가중치 (비율)
w = kept_weights(:) / 100;                 % px1

% (2) 행별로 관측된 피처만 가중합 → 가중치 재정규화
%     score = sum_j x_ij * w_j (관측치만) / sum_j w_j (관측치만)
Wrow  = repmat(w', n, 1);                  % nxp
mask  = ~isnan(X);                         % nxp
num   = nansum(X .* Wrow, 2);              % nx1, 원점수 * 가중치
den   = nansum(mask .* Wrow, 2);           % nx1, 관측된 가중치 합

score_weighted = num ./ den;               % nx1, 가중 평균 점수

% (3) 관측치가 없는 행은 NaN 처리
no_obs = (den==0 | isnan(den));
score_weighted(no_obs) = NaN;

% (4) 점수를 그대로 사용 (이미 원점수 기반이므로 0~100 범위일 것으로 예상)
%     만약 범위를 벗어나는 경우 클리핑 적용
score_100 = score_weighted;
score_100(score_100 < 0) = 0;              % 음수 방지
score_100(score_100 > 100) = 100;          % 100 초과 방지

% (5) 정수 반올림 (원하면 소수점 유지 가능)
score_100 = round(score_100);

fprintf('  ✓ 점수 계산 완료\n');
fprintf('    - NaN 점수(전부 결측): %d명\n', sum(no_obs));
fprintf('    - 점수 범위: [%s ~ %s]\n', num2str(min(score_100,[],'omitnan')), num2str(max(score_100,[],'omitnan')));
fprintf('    - 평균±표준편차: %.2f ± %.2f\n', mean(score_100,'omitnan'), std(score_100,'omitnan'));

%% 7) 결과 테이블 구성 (참조 구조 최대한 준수) -----------------------------
fprintf('\n【STEP 7】 결과 테이블 생성\n');
fprintf('--------------------------------------------\n');

% 참조 시트 컬럼(가능 시)
ref_columns = {};
try
    ref_tbl = readtable(config.reference_file, 'Sheet','역량검사_종합점수', ...
                        'VariableNamingRule','preserve');
    ref_columns = ref_tbl.Properties.VariableNames;
    fprintf('  ✓ 참조 구조 감지: %s\n', strjoin(ref_columns, ', '));
catch
    fprintf('  ⚠ 참조 시트 읽기 실패 → 기본 구조 사용( ID, 사이트, 총점 )\n');
end

% 결과 기본 컬럼
result = table();
result.ID = common_ids;

% 사이트 추출: 후보 컬럼 중 존재하는 첫 번째 사용
site_used = '';
for k = 1:numel(config.site_col_candidates)
    cand = config.site_col_candidates{k};
    if ismember(cand, hr_matched.Properties.VariableNames)
        result.('사이트') = hr_matched.(cand);
        site_used = cand;
        fprintf('  ✓ 사이트 정보: HR.%s 사용\n', cand);
        break;
    end
end
if isempty(site_used)
    result.('사이트') = repmat({''}, height(result), 1);
    fprintf('  ⚠ 사이트 컬럼 없음 → 빈값\n');
end

% 총점
result.('총점') = score_100;

% 인재유형 추가
talent_col_candidates = {'인재유형', 'TalentType', 'talent_type', '유형'};
talent_col_found = '';
for k = 1:numel(talent_col_candidates)
    cand = talent_col_candidates{k};
    if ismember(cand, hr_matched.Properties.VariableNames)
        result.('인재유형') = hr_matched.(cand);
        talent_col_found = cand;
        fprintf('  ✓ 인재유형 정보: HR.%s 사용\n', cand);
        break;
    end
end
if isempty(talent_col_found)
    result.('인재유형') = repmat({''}, height(result), 1);
    fprintf('  ⚠ 인재유형 컬럼 없음 → 빈값\n');
end

% 재직여부 추가
employment_col_candidates = {'재직여부', '근무상태', '현황', 'EmploymentStatus'};
employment_col_found = '';
for k = 1:numel(employment_col_candidates)
    cand = employment_col_candidates{k};
    if ismember(cand, hr_matched.Properties.VariableNames)
        result.('재직여부') = hr_matched.(cand);
        employment_col_found = cand;
        fprintf('  ✓ 재직여부 정보: HR.%s 사용\n', cand);
        break;
    end
end
if isempty(employment_col_found)
    result.('재직여부') = repmat({''}, height(result), 1);
    fprintf('  ⚠ 재직여부 컬럼 없음 → 빈값\n');
end

% 역진 누락 사유 추가
omission_col_candidates = {'역진 누락 사유', '역량진단 누락 사유', '누락사유', 'OmissionReason'};
omission_col_found = '';
for k = 1:numel(omission_col_candidates)
    cand = omission_col_candidates{k};
    if ismember(cand, hr_matched.Properties.VariableNames)
        result.('역진누락사유') = hr_matched.(cand);
        omission_col_found = cand;
        fprintf('  ✓ 역진 누락 사유 정보: HR.%s 사용\n', cand);
        break;
    end
end
if isempty(omission_col_found)
    result.('역진누락사유') = repmat({''}, height(result), 1);
    fprintf('  ⚠ 역진 누락 사유 컬럼 없음 → 빈값\n');
end

% 근속년수 추가
tenure_col_candidates = {'근속년수', '재직기간', '근무기간', '근속'};
tenure_col_found = '';
for k = 1:numel(tenure_col_candidates)
    cand = tenure_col_candidates{k};
    if ismember(cand, hr_matched.Properties.VariableNames)
        % 원본 문자열 그대로 추가
        result.('근속년수') = hr_matched.(cand);

        % 개월수로 변환한 값도 추가
        tenure_data = hr_matched.(cand);
        tenure_months = zeros(height(hr_matched), 1);
        for i = 1:length(tenure_data)
            tenure_str = char(tenure_data{i});
            tenure_months(i) = parse_tenure_string(tenure_str);
        end
        result.('근속개월수') = tenure_months;

        tenure_col_found = cand;
        fprintf('  ✓ 근속년수 정보: HR.%s 사용\n', cand);
        fprintf('    - 평균 근속: %.1f개월 (%.1f년)\n', mean(tenure_months(tenure_months>0)), mean(tenure_months(tenure_months>0))/12);
        break;
    end
end
if isempty(tenure_col_found)
    result.('근속년수') = repmat({''}, height(result), 1);
    result.('근속개월수') = zeros(height(result), 1);
    fprintf('  ⚠ 근속년수 컬럼 없음 → 빈값\n');
end

%% ★ 소화성 확률 계산 추가 ★
fprintf('\n【STEP 7.5】 소화성 확률 계산\n');
fprintf('────────────────────────────────────────────\n');

sohwa_model_file = 'D:\project\HR데이터\결과\소화성분석\sohwa_integrated_results_최신.mat';

% 최신 소화성 분석 결과 파일 찾기
sohwa_result_dir = 'D:\project\HR데이터\결과\소화성분석';
if exist(sohwa_result_dir, 'dir')
    files = dir(fullfile(sohwa_result_dir, 'sohwa_integrated_results_*.mat'));
    if ~isempty(files)
        [~, idx] = max([files.datenum]);
        sohwa_model_file = fullfile(files(idx).folder, files(idx).name);
        fprintf('  ✓ 소화성 모델 파일 발견: %s\n', files(idx).name);
    end
end

if exist(sohwa_model_file, 'file')
    try
        % MAT 파일 로드 (변수명 확인)
        loaded_data = load(sohwa_model_file);

        % 변수명이 'results' 또는 'integrated_results'일 수 있음
        if isfield(loaded_data, 'results')
            sohwa_results = loaded_data.results;
        elseif isfield(loaded_data, 'integrated_results')
            sohwa_results = loaded_data.integrated_results;
        else
            error('MAT 파일에 results 또는 integrated_results 변수가 없습니다.');
        end

        if isfield(sohwa_results, 'logit_model') && ...
           isfield(sohwa_results, 'zscore_params')

            logit_model = sohwa_results.logit_model;
            zscore_params = sohwa_results.zscore_params;

            % 역량명 매칭
            sohwa_probs = nan(n, 1);

            % Z-score 변환할 데이터 준비
            n_comps_sohwa = height(zscore_params);
            X_for_sohwa = nan(n, n_comps_sohwa);

            for i = 1:n_comps_sohwa
                comp_name = char(zscore_params.('역량명')(i));

                % X 행렬에서 해당 역량 찾기
                comp_idx = find(strcmp(kept_features, comp_name), 1);

                if ~isempty(comp_idx)
                    X_for_sohwa(:, i) = X(:, comp_idx);
                end
            end

            % Z-score 정규화 (소화성 모델 파라미터 사용)
            mu_sohwa = zscore_params.('평균')';
            sigma_sohwa = zscore_params.('표준편차')';
            X_sohwa_z = (X_for_sohwa - mu_sohwa) ./ sigma_sohwa;

            % NaN 제거 (결측치가 있는 행은 예측 불가)
            valid_rows = ~any(isnan(X_sohwa_z), 2);

            if sum(valid_rows) > 0
                % Logistic 확률 계산
                logits = X_sohwa_z(valid_rows, :) * logit_model.Beta + logit_model.Bias;
                sohwa_probs(valid_rows) = 1 ./ (1 + exp(-logits));

                fprintf('  ✓ 소화성 확률 계산 완료: %d명\n', sum(valid_rows));
                fprintf('    - 평균 확률: %.2f%% ± %.2f%%\n', ...
                    mean(sohwa_probs(valid_rows))*100, std(sohwa_probs(valid_rows))*100);
                fprintf('    - 범위: %.2f%% ~ %.2f%%\n', ...
                    min(sohwa_probs(valid_rows))*100, max(sohwa_probs(valid_rows))*100);
            else
                fprintf('  ⚠ 유효한 데이터 없음 (모든 행에 결측치 존재)\n');
            end

            % 결과에 소화성 확률 추가
            result.('소화성확률_퍼센트') = round(sohwa_probs * 100, 1);

        else
            fprintf('  ⚠ 소화성 모델 또는 파라미터 없음\n');
            result.('소화성확률_퍼센트') = nan(n, 1);
        end
    catch ME
        fprintf('  ⚠ 소화성 확률 계산 실패: %s\n', ME.message);
        result.('소화성확률_퍼센트') = nan(n, 1);
    end
else
    fprintf('  ⚠ 소화성 모델 파일 없음: %s\n', sohwa_model_file);
    result.('소화성확률_퍼센트') = nan(n, 1);
end

% 제외할 컬럼 목록 (참조 파일에는 있지만 결과에서는 제외)
exclude_columns = {'잡다DEV 개발자 역검', '잡다DEV 개발 구현 능력 검사'};

% 참조 컬럼에서 제외할 컬럼 필터링
if ~isempty(ref_columns)
    ref_columns = ref_columns(~ismember(ref_columns, exclude_columns));
    fprintf('  ℹ 제외 컬럼: %s\n', strjoin(exclude_columns, ', '));
end

% 참조 컬럼에 맞춰 재배치(가능하면)
if ~isempty(ref_columns)
    % 우선순위 매핑
    map = containers.Map;
    map('ID')    = 'ID';
    % 참조에서 사이트/태넌트 비슷한 것 찾아보기
    site_aliases = {'사이트','태넌트','Site','site'};
    site_key = '';
    for s = 1:numel(site_aliases)
        if ismember(site_aliases{s}, ref_columns)
            site_key = site_aliases{s}; break;
        end
    end
    if ~isempty(site_key)
        map(site_key) = '사이트';
    end
    % 점수 비슷한 것
    score_aliases = {'총점','점수','Score','score'};
    score_key = '';
    for s = 1:numel(score_aliases)
        if ismember(score_aliases{s}, ref_columns)
            score_key = score_aliases{s}; break;
        end
    end
    if ~isempty(score_key)
        map(score_key) = '총점';
    end

    % 최종 테이블(없는 참조 컬럼은 생성)
    out = table();
    for c = 1:numel(ref_columns)
        rc = ref_columns{c};
        if isKey(map, rc) && ismember(map(rc), result.Properties.VariableNames)
            out.(rc) = result.(map(rc));
        else
            % 참조에는 있는데 결과에 없는 경우: 빈값/NaN
            if height(result)>0
                if contains(rc, {'점','score','Score'})
                    out.(rc) = nan(height(result),1);
                else
                    out.(rc) = repmat({''}, height(result),1);
                end
            else
                out.(rc) = [];
            end
        end
    end

    % 결과에만 있는 나머지 컬럼 추가(정보 보존)
    map_values = values(map);
    if iscell(map_values)
        map_values_set = map_values;
    else
        map_values_set = cellstr(map_values);
    end
    extra = setdiff(result.Properties.VariableNames, map_values_set);
    for e = 1:numel(extra)
        out.(extra{e}) = result.(extra{e});
    end
    result = out;
else
    % 참조 컬럼이 없으면 result 그대로 사용 (소화성확률_퍼센트 포함)
    fprintf('  ⚠ 참조 구조 없음 → 기본 순서: ID, 사이트, 총점, 소화성확률_퍼센트\n');
end

fprintf('  ✓ 결과 테이블: %d행 x %d열\n', height(result), width(result));

%% 8) 저장(메인/가중치/메타데이터) ----------------------------------------
fprintf('\n【STEP 8】 엑셀 저장\n');
fprintf('--------------------------------------------\n');

% 백업 폴더 생성
backup_dir = fullfile(config.new_output_dir, 'backup');
if ~exist(backup_dir, 'dir')
    mkdir(backup_dir);
    fprintf('  ✓ 백업 디렉토리 생성: %s\n', backup_dir);
end

% 기존 파일 백업 (역량검사_가중치적용점수_개발자로 시작하는 모든 파일)
existing_files = dir(fullfile(config.new_output_dir, '역량검사_가중치적용점수_개발자*.xlsx'));
if ~isempty(existing_files)
    fprintf('  📦 기존 파일 백업 중...\n');
    for i = 1:length(existing_files)
        src_file = fullfile(existing_files(i).folder, existing_files(i).name);
        dst_file = fullfile(backup_dir, existing_files(i).name);
        try
            movefile(src_file, dst_file);
            fprintf('    • %s → backup/\n', existing_files(i).name);
        catch ME
            fprintf('    ⚠ 백업 실패: %s (%s)\n', existing_files(i).name, ME.message);
        end
    end
    fprintf('  ✓ %d개 파일 백업 완료\n', length(existing_files));
else
    fprintf('  ℹ 백업할 기존 파일 없음\n');
end

output_path = fullfile(config.new_output_dir, config.output_filename);

% 메인
writetable(result, output_path, 'Sheet','역량검사_종합점수', 'WriteMode','overwrite');
fprintf('  ✓ 시트 저장: 역량검사_종합점수\n');

% 가중치 상세(매칭된 항목만)
weight_detail = table();
weight_detail.('역량명') = kept_features;
weight_detail.('가중치_퍼센트') = kept_weights;
weight_detail = sortrows(weight_detail, '가중치_퍼센트','descend');
weight_detail.('순위') = (1:height(weight_detail)).';

writetable(weight_detail, output_path, 'Sheet','가중치_상세정보', 'WriteMode','append');
fprintf('  ✓ 시트 저장: 가중치_상세정보\n');

% 메타데이터
meta = table();
meta.('항목') = { ...
    '생성일시'; '원본_분석파일'; '사용된_시트'; '매칭된_샘플수'; ...
    '사용된_역량수'; '가중치_방법'; '정규화'; '행별_가중치재정규화'; ...
    '점수_최솟값'; '점수_최댓값'; ...
    '점수_평균'; '점수_표준편차'; '점수_중앙값'; ...
    '전부결측_NaN행_수' ...
    };
meta.('값') = { ...
    datestr(now,'yyyy-mm-dd HH:MM:SS'); ...
    'competency_weighted_score_export_final.m'; ...
    config.comp_sheet; ...
    sprintf('%d', numel(common_ids)); ...
    sprintf('%d', numel(kept_features)); ...
    '로지스틱 회귀 + 비용민감 (가중치 입력)'; ...
    '정규화 없음 (원점수 사용)'; ...
    '관측 피처만 가중합 후 행별 재정규화'; ...
    num2str(min(score_100,[],'omitnan')); ...
    num2str(max(score_100,[],'omitnan')); ...
    sprintf('%.2f', mean(score_100,'omitnan')); ...
    sprintf('%.2f', std(score_100,'omitnan')); ...
    sprintf('%.2f', median(score_100,'omitnan')); ...
    sprintf('%d', sum(no_obs)) ...
    };

writetable(meta, output_path, 'Sheet','메타데이터', 'WriteMode','append');
fprintf('  ✓ 시트 저장: 메타데이터\n');

fprintf('\n  ✅ 전체 파일 저장 완료: %s\n', output_path);

%% 9) 콘솔 요약 ------------------------------------------------------------
fprintf('\n【STEP 9】 요약\n');
fprintf('--------------------------------------------\n');
fprintf('📁 출력 디렉토리: %s\n', config.new_output_dir);
fprintf('📊 생성 파일: %s\n', config.output_filename);
fprintf('🎯 매칭 샘플: %d명\n', numel(common_ids));
fprintf('📈 사용 역량: %d개\n', numel(kept_features));
fprintf('📊 점수 통계 (원점수 가중합):\n');
fprintf('   • 범위: %s ~ %s\n', num2str(min(score_100,[],'omitnan')), num2str(max(score_100,[],'omitnan')));
fprintf('   • 평균: %.2f ± %.2f\n', mean(score_100,'omitnan'), std(score_100,'omitnan'));
fprintf('   • 중앙값: %.2f\n', median(score_100,'omitnan'));

% 상위/하위 5명 (NaN 제외)
valid = ~isnan(score_100);
[~, ord] = sort(score_100(valid), 'descend');
valid_ids = result.ID(valid);
topk = min(5, sum(valid));

%% ========================================================================
%  보조 함수: 근속년수 문자열 파싱
%  ========================================================================
function months = parse_tenure_string(tenure_str)
    % 근속년수 문자열을 개월 수로 변환
    % 입력 예시: "2년 3개월", "1년", "6개월", "2년3개월" 등
    % 출력: 개월 수 (숫자)
    
    months = 0;
    
    if isempty(tenure_str) || ~ischar(tenure_str)
        return;
    end
    
    % 공백 제거
    tenure_str = strtrim(tenure_str);
    
    % 년 추출
    year_pattern = '(\d+)\s*년';
    year_match = regexp(tenure_str, year_pattern, 'tokens');
    if ~isempty(year_match)
        years = str2double(year_match{1}{1});
        months = months + years * 12;
    end
    
    % 개월 추출
    month_pattern = '(\d+)\s*개월';
    month_match = regexp(tenure_str, month_pattern, 'tokens');
    if ~isempty(month_match)
        months_val = str2double(month_match{1}{1});
        months = months + months_val;
    end
    
    % 숫자만 있는 경우 (개월로 가정)
    if months == 0
        num_pattern = '^\s*(\d+)\s*$';
        num_match = regexp(tenure_str, num_pattern, 'tokens');
        if ~isempty(num_match)
            months = str2double(num_match{1}{1});
        end
    end
end