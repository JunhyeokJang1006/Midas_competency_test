% =======================================================================
%                 역량검사 가중치 적용 점수 추출 및 저장 시스템
% =======================================================================
%
% 목적: competency_statistical_analysis_order_logistic_revised.m에서
%       계산된 가중치를 적용한 역량검사 점수를 추출하여
%       '23-25년 역량검사.xlsx'의 '역량검사_종합점수' 시트와 동일한 형태로 저장
%
% 주요 기능:
% 1. 기존 분석 결과에서 가중치 적용된 점수 추출
% 2. 참조 엑셀 파일과 동일한 구조로 데이터 재구성
% 3. 결과를 '자가불소_revised' 디렉토리에 저장
% 4. 다양한 가중치 방법론별 점수 제공 (상관기반, 로지스틱, 앙상블 등)
%
% 작성자: Claude Code
% 작성일: 2025-09-23
% =======================================================================

clear; clc; close all;


% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 12);
set(0, 'DefaultTextFontSize', 12);

fprintf('=========================================\n');
fprintf('   역량검사 가중치 적용 점수 추출 시스템\n');
fprintf('=========================================\n\n');

%% 1. 설정 및 경로 정의
fprintf('【STEP 1】 설정 및 경로 정의\n');
fprintf('----------------------------------------\n');

% 기본 경로 설정
config = struct();
config.original_output_dir = 'D:\project\HR데이터\결과\자가불소';
config.new_output_dir = 'D:\project\HR데이터\결과\자가불소_revised';
config.reference_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 새 출력 디렉토리 생성
if ~exist(config.new_output_dir, 'dir')
    mkdir(config.new_output_dir);
    fprintf('  ✓ 새 출력 디렉토리 생성: %s\n', config.new_output_dir);
else
    fprintf('  ✓ 출력 디렉토리 확인: %s\n', config.new_output_dir);
end

%% 2. 기존 분석 결과 로드
fprintf('\n【STEP 2】 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

% 가중치 파일 로드
weight_file = fullfile(config.original_output_dir, 'cost_sensitive_weights.mat');
if exist(weight_file, 'file')
    fprintf('  ✓ 가중치 파일 로드: %s\n', weight_file);
    load(weight_file);

    % 로드된 변수 확인
    if exist('result_data', 'var')
        fprintf('  ✓ result_data 구조체 로드 완료\n');

        % 필수 필드 확인
        required_fields = {'final_weights', 'weighted_scores', 'feature_names', ...
                          'competency_data', 'hr_data', 'valid_indices'};
        missing_fields = {};

        for i = 1:length(required_fields)
            if ~isfield(result_data, required_fields{i})
                missing_fields{end+1} = required_fields{i};
            else
                fprintf('    - %s: 확인\n', required_fields{i});
            end
        end

        if ~isempty(missing_fields)
            fprintf('  ⚠ 누락된 필드: %s\n', strjoin(missing_fields, ', '));
        end
    else
        error('result_data 구조체를 찾을 수 없습니다.');
    end

    % 통합 가중치 결과 확인
    integrated_file = fullfile(config.original_output_dir, 'integrated_analysis_results.mat');
    if exist(integrated_file, 'file')
        fprintf('  ✓ 통합 분석 결과 로드: %s\n', integrated_file);
        load(integrated_file);

        if exist('integrated_results', 'var') && isfield(integrated_results, 'weight_comparison')
            fprintf('    - weight_comparison 테이블 확인\n');
            weight_comparison = integrated_results.weight_comparison;
        end

        if exist('integrated_results', 'var') && isfield(integrated_results, 'prediction_results')
            fprintf('    - prediction_results 확인\n');
            prediction_results = integrated_results.prediction_results;
        end
    else
        fprintf('  ⚠ 통합 분석 결과 파일 없음\n');
    end

else
    error('가중치 파일을 찾을 수 없습니다: %s', weight_file);
end

%% 3. 참조 파일 구조 분석
fprintf('\n【STEP 3】 참조 파일 구조 분석\n');
fprintf('----------------------------------------\n');

% 참조 엑셀 파일 읽기
try
    reference_data = readtable(config.reference_file, 'Sheet', '역량검사_종합점수', ...
                              'VariableNamingRule', 'preserve');
    fprintf('  ✓ 참조 파일 로드 완료\n');
    fprintf('    - 크기: %d행 x %d열\n', height(reference_data), width(reference_data));

    % 컬럼 구조 확인
    ref_columns = reference_data.Properties.VariableNames;
    fprintf('    - 컬럼: %s\n', strjoin(ref_columns, ', '));

catch ME
    error('참조 파일 읽기 실패: %s', ME.message);
end

%% 4. 데이터 로드 및 전처리
fprintf('\n【STEP 4】 데이터 로드 및 전처리\n');
fprintf('----------------------------------------\n');

% 기본 데이터 추출 시도
data_loaded = false;
if exist('result_data', 'var')
    if isfield(result_data, 'competency_data') && isfield(result_data, 'hr_data') && ...
       isfield(result_data, 'valid_indices') && isfield(result_data, 'feature_names')

        competency_data = result_data.competency_data;
        hr_data = result_data.hr_data;
        valid_indices = result_data.valid_indices;
        feature_names = result_data.feature_names;
        final_weights = result_data.final_weights;
        data_loaded = true;

        fprintf('  ✓ result_data에서 기본 데이터 추출 완료\n');
        fprintf('    - 유효 샘플 수: %d개\n', length(valid_indices));
        fprintf('    - 역량 개수: %d개\n', length(feature_names));
    end
end

% 데이터가 없으면 원본 파일에서 직접 로드
if ~data_loaded
    fprintf('  ⚠ result_data에 필요한 필드가 없음. 원본 데이터 직접 로드\n');

    % HR 데이터 로드
    try
        fprintf('    - HR 데이터 로드 중...\n');
        hr_data = readtable(config.hr_file, 'VariableNamingRule', 'preserve');
        fprintf('      ✓ HR 데이터 로드 완료: %d행\n', height(hr_data));
    catch ME
        error('HR 데이터 로드 실패: %s', ME.message);
    end

    % 역량 데이터 로드
    try
        fprintf('    - 역량 데이터 로드 중...\n');
        competency_data = readtable(config.comp_file, 'VariableNamingRule', 'preserve');
        fprintf('      ✓ 역량 데이터 로드 완료: %d행\n', height(competency_data));
    catch ME
        error('역량 데이터 로드 실패: %s', ME.message);
    end

    % 데이터 병합 (ID 기준)
    fprintf('    - 데이터 병합 중...\n');

    % 공통 ID 찾기
    if ismember('ID', hr_data.Properties.VariableNames) && ismember('Var1', competency_data.Properties.VariableNames)
        [merged_data, hr_idx, comp_idx] = innerjoin(hr_data, competency_data, 'Keys', 'ID');
        valid_indices = hr_idx;  % HR 데이터 기준 인덱스
        fprintf('      ✓ ID 기준 병합 완료: %d개 매칭\n', length(valid_indices));
    else
        error('HR 데이터와 역량 데이터에서 공통 ID 컬럼을 찾을 수 없습니다.');
    end

    % 역량 컬럼 추출
    comp_cols = competency_data.Properties.VariableNames;
    excluded_cols = {'Var1', 'ID', '이름', '부서명', '직책명', '입사일', '인재유형', '성과점수', ...
                    '순위', '등급', '사원번호', '사번', '직급명', '조직명', '인재유형_최종'};

    feature_names = {};
    for i = 1:length(comp_cols)
        if ~any(strcmpi(comp_cols{i}, excluded_cols))
            feature_names{end+1} = comp_cols{i};
        end
    end

    fprintf('      ✓ 유효 역량 컬럼 추출: %d개\n', length(feature_names));

    % 가중치 확인
    if exist('result_data', 'var') && isfield(result_data, 'final_weights')
        final_weights = result_data.final_weights;
        fprintf('      ✓ 기존 가중치 사용: %d개\n', length(final_weights));
    else
        % 가중치가 없으면 균등 가중치 사용
        final_weights = ones(length(feature_names), 1) * (100 / length(feature_names));
        fprintf('      ⚠ 가중치 없음. 균등 가중치 사용\n');
    end

    data_loaded = true;
end

%% 5. 다양한 가중치 방법론별 점수 계산
fprintf('\n【STEP 5】 다양한 가중치 방법론별 점수 계산\n');
fprintf('----------------------------------------\n');

% X_normalized 재구성 (기존 코드와 동일한 방식)
if exist('result_data', 'var') && isfield(result_data, 'X_normalized')
    X_normalized = result_data.X_normalized;
    fprintf('  ✓ 정규화된 특성 행렬 로드\n');
else
    % X_normalized가 없으면 재계산
    fprintf('  ⚠ X_normalized 재계산 중...\n');

    % 특성 행렬 구성
    X_raw = [];
    valid_feature_names = {};

    for i = 1:length(feature_names)
        try
            % 한글 컬럼명 처리
            if ismember(feature_names{i}, competency_data.Properties.VariableNames)
                col_data = competency_data.(feature_names{i})(valid_indices);
            else
                % 직접 접근이 안되면 table2array 사용
                col_idx = strcmp(competency_data.Properties.VariableNames, feature_names{i});
                if any(col_idx)
                    temp_data = table2array(competency_data(:, col_idx));
                    col_data = temp_data(valid_indices);
                else
                    continue;  % 컬럼을 찾을 수 없으면 건너뛰기
                end
            end

            % 숫자 데이터만 사용
            if isnumeric(col_data) && ~all(isnan(col_data))
                X_raw = [X_raw, col_data];
                valid_feature_names{end+1} = feature_names{i};
            end
        catch
            fprintf('      ⚠ 컬럼 "%s" 처리 실패, 건너뛰기\n', feature_names{i});
            continue;
        end
    end

    if isempty(X_raw)
        error('유효한 역량 데이터를 찾을 수 없습니다.');
    end

    % 정규화
    X_normalized = (X_raw - mean(X_raw, 1)) ./ (std(X_raw, 0, 1) + eps);
    feature_names = valid_feature_names;  % 실제 사용된 feature names로 업데이트

    fprintf('    - 정규화 완료: %d행 x %d열\n', size(X_normalized, 1), size(X_normalized, 2));
    fprintf('    - 유효 역량 수: %d개\n', length(feature_names));

    % 가중치 크기 조정 (feature 수가 다를 경우)
    if length(final_weights) ~= length(feature_names)
        fprintf('    ⚠ 가중치 개수와 역량 개수 불일치. 균등 가중치로 대체\n');
        final_weights = ones(length(feature_names), 1) * (100 / length(feature_names));
    end
end

% 가중치 방법론별 점수 계산
weighted_scores_methods = struct();

% 1. 로지스틱 회귀 가중치 (메인)
logistic_weights = final_weights / 100;  % 백분율을 비율로 변환
weighted_scores_methods.Logistic = X_normalized * logistic_weights;
fprintf('  ✓ 로지스틱 회귀 가중치 점수 계산 완료\n');

% 2. 상관기반 가중치 (통합 결과에서)
if exist('weight_comparison', 'var')
    % 상관기반
    corr_weights = zeros(length(feature_names), 1);
    for i = 1:length(feature_names)
        idx = strcmp(weight_comparison.Feature, feature_names{i});
        if any(idx)
            corr_weights(i) = weight_comparison.Correlation(idx) / 100;
        end
    end
    weighted_scores_methods.Correlation = X_normalized * corr_weights;

    % Bootstrap 평균
    boot_weights = zeros(length(feature_names), 1);
    for i = 1:length(feature_names)
        idx = strcmp(weight_comparison.Feature, feature_names{i});
        if any(idx)
            boot_weights(i) = weight_comparison.Bootstrap(idx) / 100;
        end
    end
    weighted_scores_methods.Bootstrap = X_normalized * boot_weights;

    % 앙상블 평균
    ensemble_weights = zeros(length(feature_names), 1);
    for i = 1:length(feature_names)
        idx = strcmp(weight_comparison.Feature, feature_names{i});
        if any(idx)
            ensemble_weights(i) = weight_comparison.Ensemble_Mean(idx) / 100;
        end
    end
    weighted_scores_methods.Ensemble = X_normalized * ensemble_weights;

    fprintf('  ✓ 추가 가중치 방법론 점수 계산 완료\n');
    fprintf('    - 상관기반, Bootstrap, 앙상블 방법\n');
end

%% 6. 결과 데이터 테이블 생성
fprintf('\n【STEP 6】 결과 데이터 테이블 생성\n');
fprintf('----------------------------------------\n');

% 기본 정보 추출
result_table = table();

% ID 정보 (HR 데이터에서)
if ismember('ID', hr_data.Properties.VariableNames)
    result_table.ID = hr_data.ID(valid_indices);
elseif ismember('사원번호', hr_data.Properties.VariableNames)
    result_table.ID = hr_data.('사원번호')(valid_indices);
else
    % ID가 없으면 순번으로 생성
    result_table.ID = (1:length(valid_indices))';
    fprintf('  ⚠ ID 컬럼을 찾을 수 없어 순번으로 생성\n');
end

% 사이트 정보 (있으면 추가, 없으면 빈 값)
if ismember('사이트', hr_data.Properties.VariableNames)
    result_table.('사이트') = hr_data.('사이트')(valid_indices);
else
    result_table.('사이트') = repmat({''}, length(valid_indices), 1);
end

% 빈 컬럼들 (참조 파일 구조에 맞춤)
result_table.('기타정보1') = repmat({''}, length(valid_indices), 1);
result_table.('기타정보2') = repmat({''}, length(valid_indices), 1);

% 메인 가중치 적용 점수 (로지스틱 회귀 방법)
result_table.('총점') = round(weighted_scores_methods.Logistic, 2);

% 추가 방법론별 점수 (있는 경우)
if isfield(weighted_scores_methods, 'Correlation')
    result_table.('총점_상관기반') = round(weighted_scores_methods.Correlation, 2);
end
if isfield(weighted_scores_methods, 'Bootstrap')
    result_table.('총점_Bootstrap') = round(weighted_scores_methods.Bootstrap, 2);
end
if isfield(weighted_scores_methods, 'Ensemble')
    result_table.('총점_앙상블') = round(weighted_scores_methods.Ensemble, 2);
end

fprintf('  ✓ 결과 테이블 생성 완료\n');
fprintf('    - 총 %d행 생성\n', height(result_table));
fprintf('    - 컬럼: %s\n', strjoin(result_table.Properties.VariableNames, ', '));

% 점수 통계
main_scores = result_table.('총점');
valid_scores = main_scores(~isnan(main_scores));

if ~isempty(valid_scores)
    fprintf('\n  【메인 가중치 점수 통계】\n');
    fprintf('    - 유효 점수 개수: %d개\n', length(valid_scores));
    fprintf('    - 점수 범위: %.2f ~ %.2f\n', min(valid_scores), max(valid_scores));
    fprintf('    - 평균: %.2f ± %.2f\n', mean(valid_scores), std(valid_scores));
end

%% 7. 엑셀 파일로 저장
fprintf('\n【STEP 7】 엑셀 파일로 저장\n');
fprintf('----------------------------------------\n');

% 저장 파일명 생성
output_filename = sprintf('역량검사_가중치적용점수_%s.xlsx', config.timestamp);
output_filepath = fullfile(config.new_output_dir, output_filename);

try
    % 메인 시트 저장 (참조 파일과 동일한 구조)
    writetable(result_table, output_filepath, 'Sheet', '역량검사_종합점수', ...
               'WriteMode', 'overwrite');
    fprintf('  ✓ 메인 시트 저장 완료: 역량검사_종합점수\n');

    % 가중치 정보 시트 추가
    if exist('weight_comparison', 'var')
        writetable(weight_comparison, output_filepath, 'Sheet', '가중치_상세정보', ...
                   'WriteMode', 'append');
        fprintf('  ✓ 가중치 정보 시트 저장 완료\n');
    end

    % 메타데이터 시트 추가
    metadata = table();
    metadata.('항목') = {'생성일시'; '원본_분석파일'; '총_샘플수'; '유효_샘플수'; ...
                    '역량_개수'; '가중치_방법'; '점수_범위_최소'; '점수_범위_최대'; '평균_점수'};

    metadata.('값') = {datestr(now, 'yyyy-mm-dd HH:MM:SS'); ...
                  'competency_statistical_analysis_order_logistic_revised.m'; ...
                  height(competency_data); ...
                  length(valid_indices); ...
                  length(feature_names); ...
                  '로지스틱 회귀 + 비용민감 학습'; ...
                  sprintf('%.2f', min(valid_scores)); ...
                  sprintf('%.2f', max(valid_scores)); ...
                  sprintf('%.2f', mean(valid_scores))};

    writetable(metadata, output_filepath, 'Sheet', '메타데이터', 'WriteMode', 'append');
    fprintf('  ✓ 메타데이터 시트 저장 완료\n');

    fprintf('\n  ✅ 전체 파일 저장 완료: %s\n', output_filepath);

catch ME
    error('파일 저장 실패: %s', ME.message);
end

%% 8. 원본 코드 수정 (출력 디렉토리 변경)
fprintf('\n【STEP 8】 원본 코드 출력 디렉토리 수정\n');
fprintf('----------------------------------------\n');

% 원본 파일 경로
original_file = 'D:\project\HR데이터\matlab\자가불소_현주CP님 피드백\competency_statistical_analysis_order_logistic_revised.m';
modified_file = 'D:\project\HR데이터\matlab\자가불소_현주CP님 피드백\competency_statistical_analysis_order_logistic_revised_updated.m';

try
    % 원본 파일 읽기
    fid = fopen(original_file, 'r', 'n', 'UTF-8');
    if fid == -1
        error('원본 파일을 열 수 없습니다: %s', original_file);
    end

    file_content = fread(fid, '*char')';
    fclose(fid);

    % 출력 디렉토리 경로 수정
    old_path = 'D:\project\HR데이터\결과\자가불소';
    new_path = 'D:\project\HR데이터\결과\자가불소_revised';

    % 텍스트 치환
    modified_content = strrep(file_content, old_path, new_path);

    % 수정된 파일 저장
    fid = fopen(modified_file, 'w', 'n', 'UTF-8');
    if fid == -1
        error('수정된 파일을 생성할 수 없습니다: %s', modified_file);
    end

    fwrite(fid, modified_content, 'char');
    fclose(fid);

    fprintf('  ✓ 원본 코드 수정 완료\n');
    fprintf('    - 원본: %s\n', original_file);
    fprintf('    - 수정본: %s\n', modified_file);
    fprintf('    - 변경사항: 출력 디렉토리를 "%s"로 변경\n', new_path);

catch ME
    fprintf('  ⚠ 원본 코드 수정 실패: %s\n', ME.message);
    fprintf('    수동으로 다음 경로를 변경해주세요:\n');
    fprintf('    "%s" → "%s"\n', old_path, new_path);
end

%% 9. 요약 보고서 생성
fprintf('\n【STEP 9】 요약 보고서 생성\n');
fprintf('----------------------------------------\n');

% 요약 정보 출력
fprintf('\n');
fprintf('=========================================\n');
fprintf('           작업 완료 요약\n');
fprintf('=========================================\n');
fprintf('📁 출력 디렉토리: %s\n', config.new_output_dir);
fprintf('📊 생성된 파일: %s\n', output_filename);
fprintf('📈 총 샘플 수: %d개\n', length(valid_indices));
fprintf('🎯 메인 점수 범위: %.2f ~ %.2f (평균: %.2f)\n', ...
        min(valid_scores), max(valid_scores), mean(valid_scores));

if exist('weight_comparison', 'var')
    fprintf('⚖️  가중치 방법론: %d가지 (로지스틱, 상관기반, Bootstrap, 앙상블)\n', ...
            width(result_table) - 4);  % ID, 사이트, 기타정보 제외
end

fprintf('\n📋 생성된 시트:\n');
fprintf('   • 역량검사_종합점수: 메인 결과 (%d행)\n', height(result_table));
if exist('weight_comparison', 'var')
    fprintf('   • 가중치_상세정보: 방법론별 가중치 비교\n');
end
fprintf('   • 메타데이터: 분석 정보 및 통계\n');

fprintf('\n💡 사용법:\n');
fprintf('   1. 생성된 엑셀 파일을 열어서 "역량검사_종합점수" 시트 확인\n');
fprintf('   2. "총점" 컬럼이 가중치가 적용된 최종 점수\n');
fprintf('   3. 추가 방법론별 점수도 함께 제공됨 (있는 경우)\n');
fprintf('   4. 메타데이터 시트에서 자세한 분석 정보 확인 가능\n');

fprintf('\n✅ 역량검사 가중치 적용 점수 추출 완료!\n');
fprintf('=========================================\n');

%% 10. 파일 정리 및 백업
fprintf('\n【STEP 10】 파일 정리\n');
fprintf('----------------------------------------\n');

% 생성된 파일들 목록
generated_files = {output_filepath};

if exist(modified_file, 'file')
    generated_files{end+1} = modified_file;
end

fprintf('생성된 파일 목록:\n');
for i = 1:length(generated_files)
    [~, fname, ext] = fileparts(generated_files{i});
    file_info = dir(generated_files{i});
    fprintf('  %d. %s%s (%.1f KB)\n', i, fname, ext, file_info.bytes/1024);
end

fprintf('\n작업이 완전히 완료되었습니다! 🎉\n');