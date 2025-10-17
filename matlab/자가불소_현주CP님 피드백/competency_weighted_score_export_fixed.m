% =======================================================================
%                 역량검사 가중치 적용 점수 추출 및 저장 시스템 (수정됨)
% =======================================================================
%
% 목적: competency_statistical_analysis_order_logistic_revised.m에서
%       계산된 가중치를 적용한 역량검사 점수를 추출하여
%       '23-25년 역량검사.xlsx'의 '역량검사_종합점수' 시트와 동일한 형태로 저장
%
% 주요 기능:
% 1. 기존 분석 결과에서 가중치 추출
% 2. 원본 데이터를 직접 로드하여 가중치 적용
% 3. 참조 엑셀 파일과 동일한 구조로 데이터 재구성
% 4. 결과를 '자가불소_revised' 디렉토리에 저장
%
% 작성자: Claude Code
% 작성일: 2025-09-23 (수정됨)
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

%% 2. 가중치 파일 로드
fprintf('\n【STEP 2】 가중치 파일 로드\n');
fprintf('----------------------------------------\n');

% 가중치 파일 로드
weight_file = fullfile(config.original_output_dir, 'cost_sensitive_weights.mat');
if exist(weight_file, 'file')
    fprintf('  ✓ 가중치 파일 로드: %s\n', weight_file);
    load(weight_file);

    % 가중치 정보 추출
    if exist('result_data', 'var') && isfield(result_data, 'final_weights') && isfield(result_data, 'feature_names')
        final_weights = result_data.final_weights;
        feature_names = result_data.feature_names;
        fprintf('  ✓ 가중치 정보 추출 완료\n');
        fprintf('    - 역량 개수: %d개\n', length(feature_names));
        fprintf('    - 가중치 범위: %.2f%% ~ %.2f%%\n', min(final_weights), max(final_weights));
    else
        error('가중치 파일에서 필요한 정보를 찾을 수 없습니다.');
    end
else
    error('가중치 파일을 찾을 수 없습니다: %s', weight_file);
end

%% 3. 원본 데이터 로드
fprintf('\n【STEP 3】 원본 데이터 로드\n');
fprintf('----------------------------------------\n');

% HR 데이터 로드
try
    fprintf('  • HR 데이터 로드 중...\n');
    hr_data = readtable(config.hr_file, 'VariableNamingRule', 'preserve');
    fprintf('    ✓ HR 데이터 로드 완료: %d행 x %d열\n', height(hr_data), width(hr_data));
    fprintf('    - 컬럼: %s\n', strjoin(hr_data.Properties.VariableNames(1:min(5,end)), ', '));
catch ME
    error('HR 데이터 로드 실패: %s', ME.message);
end

% 역량 데이터 로드
try
    fprintf('\n  • 역량 데이터 로드 중...\n');
    competency_data = readtable(config.comp_file, 'VariableNamingRule', 'preserve');
    fprintf('    ✓ 역량 데이터 로드 완료: %d행 x %d열\n', height(competency_data), width(competency_data));
    fprintf('    - 컬럼: %s\n', strjoin(competency_data.Properties.VariableNames(1:min(5,end)), ', '));
catch ME
    error('역량 데이터 로드 실패: %s', ME.message);
end

%% 4. 데이터 병합 및 매칭
fprintf('\n【STEP 4】 데이터 병합 및 매칭\n');
fprintf('----------------------------------------\n');

% ID 컬럼 확인 및 병합
hr_id_col = '';
comp_id_col = '';

% HR 데이터에서 ID 컬럼 찾기
if ismember('ID', hr_data.Properties.VariableNames)
    hr_id_col = 'ID';
elseif ismember('사원번호', hr_data.Properties.VariableNames)
    hr_id_col = '사원번호';
else
    error('HR 데이터에서 ID 컬럼을 찾을 수 없습니다.');
end

% 역량 데이터에서 ID 컬럼 찾기 (Var1이 실제 ID일 가능성)
if ismember('ID', competency_data.Properties.VariableNames)
    comp_id_col = 'ID';
elseif ismember('Var1', competency_data.Properties.VariableNames)
    comp_id_col = 'Var1';
else
    error('역량 데이터에서 ID 컬럼을 찾을 수 없습니다.');
end

fprintf('  • ID 매칭: HR.%s ↔ Competency.%s\n', hr_id_col, comp_id_col);

% 데이터 병합 수행
try
    % HR 데이터의 ID
    hr_ids = hr_data.(hr_id_col);

    % 역량 데이터의 ID
    comp_ids = competency_data.(comp_id_col);

    % 공통 ID 찾기
    [common_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_ids);

    fprintf('  ✓ 공통 ID 매칭 완료: %d개\n', length(common_ids));
    fprintf('    - HR 전체: %d개, 역량 전체: %d개\n', length(hr_ids), length(comp_ids));

    if length(common_ids) == 0
        error('공통 ID가 없습니다. 데이터 확인이 필요합니다.');
    end

    % 매칭된 데이터만 선택
    hr_matched = hr_data(hr_idx, :);
    comp_matched = competency_data(comp_idx, :);

catch ME
    error('데이터 병합 실패: %s', ME.message);
end

%% 5. 역량 데이터 전처리 및 특성 행렬 구성
fprintf('\n【STEP 5】 역량 데이터 전처리\n');
fprintf('----------------------------------------\n');

% 역량 컬럼 추출 (숫자 데이터만)
fprintf('  • 유효한 역량 컬럼 추출 중...\n');

% 제외할 컬럼들
excluded_cols = {'Var1', 'Var2', 'Var3', 'Var4', 'Var5', 'ID', '이름', '부서명', '직책명', ...
                '입사일', '인재유형', '성과점수', '순위', '등급', '사원번호', '사번', '직급명', ...
                '조직명', '인재유형_최종'};

all_comp_cols = comp_matched.Properties.VariableNames;
valid_comp_cols = {};
X_raw = [];

for i = 1:length(all_comp_cols)
    col_name = all_comp_cols{i};

    % 제외 컬럼 체크
    if any(strcmpi(col_name, excluded_cols))
        continue;
    end

    % 데이터 추출
    try
        col_data = comp_matched.(col_name);

        % 숫자 데이터이고 유효한 값이 있는지 확인
        if isnumeric(col_data) && sum(~isnan(col_data)) > 0
            valid_comp_cols{end+1} = col_name;
            X_raw = [X_raw, col_data];
        end
    catch
        % 한글 컬럼명 문제 시 건너뛰기
        continue;
    end
end

fprintf('    ✓ 유효 역량 컬럼: %d개\n', length(valid_comp_cols));
fprintf('    ✓ 특성 행렬 크기: %d행 x %d열\n', size(X_raw, 1), size(X_raw, 2));

if isempty(X_raw)
    error('유효한 역량 데이터를 찾을 수 없습니다.');
end

%% 6. 가중치와 역량 컬럼 매칭
fprintf('\n【STEP 6】 가중치와 역량 컬럼 매칭\n');
fprintf('----------------------------------------\n');

% 가중치의 feature_names와 실제 컬럼 매칭
matched_weights = [];
matched_features = {};
matched_data = [];

fprintf('  • 가중치 매칭 중...\n');

for i = 1:length(feature_names)
    feature_name = feature_names{i};

    % 실제 컬럼에서 찾기
    col_idx = find(strcmp(valid_comp_cols, feature_name));

    if ~isempty(col_idx)
        matched_weights = [matched_weights; final_weights(i)];
        matched_features{end+1} = feature_name;
        matched_data = [matched_data, X_raw(:, col_idx(1))];

        if i <= 10  % 처음 10개만 출력
            fprintf('    ✓ %s: %.2f%%\n', feature_name, final_weights(i));
        end
    end
end

fprintf('  ✓ 매칭된 역량: %d개 (전체 %d개 중)\n', length(matched_features), length(feature_names));

if isempty(matched_data)
    error('가중치와 매칭되는 역량 데이터가 없습니다.');
end

%% 7. 정규화 및 가중치 적용 점수 계산
fprintf('\n【STEP 7】 점수 계산\n');
fprintf('----------------------------------------\n');

% 데이터 정규화
fprintf('  • 데이터 정규화 중...\n');
X_normalized = (matched_data - mean(matched_data, 1)) ./ (std(matched_data, 0, 1) + eps);

% 가중치 정규화 (백분율을 비율로)
weights_normalized = matched_weights / 100;

% 가중치 적용 점수 계산
weighted_scores = X_normalized * weights_normalized;

fprintf('  ✓ 가중치 적용 점수 계산 완료\n');
fprintf('    - 점수 범위: %.2f ~ %.2f\n', min(weighted_scores), max(weighted_scores));
fprintf('    - 평균: %.2f ± %.2f\n', mean(weighted_scores), std(weighted_scores));

%% 8. 결과 테이블 생성
fprintf('\n【STEP 8】 결과 테이블 생성\n');
fprintf('----------------------------------------\n');

% 기본 정보로 테이블 구성
result_table = table();

% ID 정보
result_table.ID = common_ids;

% 사이트 정보 (HR 데이터에서)
if ismember('사이트', hr_matched.Properties.VariableNames)
    result_table.('사이트') = hr_matched.('사이트');
elseif ismember('Var2', hr_matched.Properties.VariableNames)
    % Var2가 사이트 정보일 가능성
    result_table.('사이트') = hr_matched.Var2;
else
    result_table.('사이트') = repmat({''}, length(common_ids), 1);
end

% 빈 컬럼들 (참조 파일 구조에 맞춤)
result_table.('기타정보1') = repmat({''}, length(common_ids), 1);
result_table.('기타정보2') = repmat({''}, length(common_ids), 1);

% 가중치 적용 총점
result_table.('총점') = round(weighted_scores, 2);

fprintf('  ✓ 결과 테이블 생성 완료\n');
fprintf('    - 총 %d행 생성\n', height(result_table));
fprintf('    - 컬럼: %s\n', strjoin(result_table.Properties.VariableNames, ', '));

%% 9. 엑셀 파일로 저장
fprintf('\n【STEP 9】 엑셀 파일로 저장\n');
fprintf('----------------------------------------\n');

% 저장 파일명 생성
output_filename = sprintf('역량검사_가중치적용점수_%s.xlsx', config.timestamp);
output_filepath = fullfile(config.new_output_dir, output_filename);

try
    % 메인 시트 저장
    writetable(result_table, output_filepath, 'Sheet', '역량검사_종합점수', ...
               'WriteMode', 'overwrite');
    fprintf('  ✓ 메인 시트 저장 완료: 역량검사_종합점수\n');

    % 가중치 정보 시트 추가
    weight_info = table();
    weight_info.('역량명') = matched_features';
    weight_info.('가중치_퍼센트') = matched_weights;
    weight_info = sortrows(weight_info, '가중치_퍼센트', 'descend');

    writetable(weight_info, output_filepath, 'Sheet', '가중치_정보', ...
               'WriteMode', 'append');
    fprintf('  ✓ 가중치 정보 시트 저장 완료\n');

    % 메타데이터 시트 추가
    metadata = table();
    metadata.('항목') = {'생성일시'; '매칭된_샘플수'; '사용된_역량수'; '가중치_방법'; ...
                        '점수_최솟값'; '점수_최댓값'; '점수_평균'; '점수_표준편차'};

    metadata.('값') = {datestr(now, 'yyyy-mm-dd HH:MM:SS'); ...
                      sprintf('%d', length(common_ids)); ...
                      sprintf('%d', length(matched_features)); ...
                      '로지스틱 회귀 + 비용민감 학습'; ...
                      sprintf('%.2f', min(weighted_scores)); ...
                      sprintf('%.2f', max(weighted_scores)); ...
                      sprintf('%.2f', mean(weighted_scores)); ...
                      sprintf('%.2f', std(weighted_scores))};

    writetable(metadata, output_filepath, 'Sheet', '메타데이터', 'WriteMode', 'append');
    fprintf('  ✓ 메타데이터 시트 저장 완료\n');

    fprintf('\n  ✅ 전체 파일 저장 완료: %s\n', output_filepath);

catch ME
    error('파일 저장 실패: %s', ME.message);
end

%% 10. 요약 보고서
fprintf('\n');
fprintf('=========================================\n');
fprintf('           작업 완료 요약\n');
fprintf('=========================================\n');
fprintf('📁 출력 디렉토리: %s\n', config.new_output_dir);
fprintf('📊 생성된 파일: %s\n', output_filename);
fprintf('🎯 매칭된 샘플: %d개\n', length(common_ids));
fprintf('📈 사용된 역량: %d개 (전체 %d개 중)\n', length(matched_features), length(feature_names));
fprintf('📊 점수 범위: %.2f ~ %.2f (평균: %.2f)\n', ...
        min(weighted_scores), max(weighted_scores), mean(weighted_scores));

fprintf('\n📋 생성된 시트:\n');
fprintf('   • 역량검사_종합점수: 메인 결과 (%d행)\n', height(result_table));
fprintf('   • 가중치_정보: 사용된 가중치 상세\n');
fprintf('   • 메타데이터: 분석 정보 및 통계\n');

fprintf('\n💡 사용법:\n');
fprintf('   1. 생성된 엑셀 파일의 "역량검사_종합점수" 시트 확인\n');
fprintf('   2. "총점" 컬럼이 가중치가 적용된 최종 점수\n');
fprintf('   3. "가중치_정보" 시트에서 각 역량별 가중치 확인\n');

% 상위/하위 점수 샘플 출력
fprintf('\n📊 점수 샘플:\n');
[~, top_idx] = maxk(weighted_scores, 3);
[~, bottom_idx] = mink(weighted_scores, 3);

fprintf('   상위 3명: ');
for i = 1:3
    fprintf('ID_%s(%.2f) ', num2str(common_ids(top_idx(i))), weighted_scores(top_idx(i)));
end
fprintf('\n');

fprintf('   하위 3명: ');
for i = 1:3
    fprintf('ID_%s(%.2f) ', num2str(common_ids(bottom_idx(i))), weighted_scores(bottom_idx(i)));
end
fprintf('\n');

fprintf('\n✅ 역량검사 가중치 적용 점수 추출 완료!\n');
fprintf('=========================================\n');