%% MIDAS 성장단계 평가 점수 산출 프로그램 (최종 개선 버전)
% 작성일: 2024
% 목적: 최근 3년 입사자의 발현 역량 데이터를 기반으로 성장 점수 계산
% 개선사항: 
%   - 세 번째 시트 직접 지정
%   - ActiveX 오류 완전 해결
%   - 메모리 효율성 개선
%   - 에러 처리 강화
%   - 성능 최적화
%   - 사용자 인터페이스 개선

clear; clc; close all;

%% 1. 초기 설정 및 파일 경로 설정
fprintf('=== MIDAS 성장단계 평가 분석 시작 ===\n');
fprintf('프로그램 버전: 2.0 (최종 개선 버전)\n');
fprintf('시작 시간: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

% 파일 경로 설정 (유연한 경로 처리)
base_path = 'D:\project\HR데이터\데이터\역량검사 요청 정보';
filename = fullfile(base_path, '최근 3년 입사자_인적정보.xlsx');

% 파일 존재 여부 확인
if ~exist(filename, 'file')
    fprintf('오류: 파일을 찾을 수 없습니다: %s\n', filename);
    fprintf('대안 경로를 확인하세요:\n');
    fprintf('  - %s\n', filename);
    error('파일을 찾을 수 없습니다.');
end

fprintf('파일 확인 완료: %s\n', filename);

%% 2. 데이터 로드 및 검증
fprintf('\n=== 데이터 로드 시작 ===\n');

try
    % 엑셀 파일의 시트 정보 확인
    [status, sheets] = xlsfinfo(filename);
    
    if isempty(sheets)
        error('엑셀 파일에서 시트를 찾을 수 없습니다.');
    end
    
    fprintf('발견된 시트 (%d개):\n', length(sheets));
    for i = 1:length(sheets)
        fprintf('  %d. %s\n', i, sheets{i});
    end
    
    % 세 번째 시트를 성장단계 시트로 직접 지정
    if length(sheets) >= 3
        sheet_name = sheets{3};
        fprintf('\n✓ 세 번째 시트 선택: %s\n', sheet_name);
    else
        error('세 번째 시트가 존재하지 않습니다. 총 %d개의 시트만 발견됨.', length(sheets));
    end
    
    % 데이터 읽기 (다중 방법 시도)
    fprintf('데이터 읽는 중...\n');
    raw_data = [];
    data_source = '';
    
    % 방법 1: readtable 사용 (권장)
    try
        data_table = readtable(filename, 'Sheet', sheet_name, ...
                              'VariableNamingRule', 'preserve', ...
                              'TextType', 'string');
        raw_data = table2cell(data_table);
        data_source = 'readtable';
        fprintf('✓ readtable로 데이터 읽기 성공\n');
    catch ME1
        fprintf('⚠ readtable 실패: %s\n', ME1.message);
        
        % 방법 2: xlsread 사용 (대안)
        try
            [~, ~, raw_data] = xlsread(filename, sheet_name);
            data_source = 'xlsread';
            fprintf('✓ xlsread로 데이터 읽기 성공\n');
        catch ME2
            fprintf('⚠ xlsread도 실패: %s\n', ME2.message);
            
            % 방법 3: readmatrix 사용 (최후 수단)
            try
                raw_data = readmatrix(filename, 'Sheet', sheet_name);
                data_source = 'readmatrix';
                fprintf('✓ readmatrix로 데이터 읽기 성공\n');
            catch ME3
                error('모든 데이터 읽기 방법 실패:\n- readtable: %s\n- xlsread: %s\n- readmatrix: %s', ...
                      ME1.message, ME2.message, ME3.message);
            end
        end
    end
    
    % 데이터 구조 파악
    [n_rows, n_cols] = size(raw_data);
    fprintf('✓ 데이터 크기: %d행 x %d열 (읽기 방법: %s)\n', n_rows, n_cols, data_source);
    
    % 데이터 유효성 검사
    if n_rows < 2
        error('데이터가 너무 적습니다. 최소 2행(헤더+데이터)이 필요합니다.');
    end
    
    if n_cols < 2
        error('데이터가 너무 적습니다. 최소 2열(ID+평가데이터)이 필요합니다.');
    end
    
catch ME
    fprintf('\n❌ 데이터 로드 실패!\n');
    fprintf('에러 메시지: %s\n', ME.message);
    if ~isempty(ME.stack)
        fprintf('에러 위치: %s (라인 %d)\n', ME.stack(1).name, ME.stack(1).line);
    end
    error('데이터 로드 오류: %s', ME.message);
end

%% 3. 데이터 구조 분석 및 헤더 찾기
fprintf('\n=== 데이터 구조 분석 ===\n');

% 안전한 데이터 미리보기
fprintf('데이터 미리보기 (첫 3행, 처음 8열):\n');
for i = 1:min(3, n_rows)
    fprintf('행 %d: ', i);
    for j = 1:min(8, n_cols)
        try
            cell_value = raw_data{i,j};
            if ischar(cell_value) || isstring(cell_value)
                fprintf('[%s] ', char(cell_value));
            elseif isnumeric(cell_value)
                if isnan(cell_value)
                    fprintf('[NaN] ');
                else
                    fprintf('[%.0f] ', cell_value);
                end
            elseif isa(cell_value, 'missing')
                fprintf('[Missing] ');
            elseif contains(class(cell_value), 'ActiveX')
                fprintf('[ActiveX_ERROR] ');
            else
                fprintf('[%s] ', class(cell_value));
            end
        catch
            fprintf('[ERROR] ');
        end
    end
    fprintf('\n');
end

% 헤더 행 찾기 (개선된 로직)
header_row = 0;
header_keywords = {'발현', 'H1', 'H2', 'ID', '사번', 'id', 'ID'};

for i = 1:min(10, n_rows)
    row_data = raw_data(i, :);
    keyword_found = false;
    
    for j = 1:length(row_data)
        if ischar(row_data{j}) || isstring(row_data{j})
            cell_str = char(row_data{j});
            for k = 1:length(header_keywords)
                if contains(cell_str, header_keywords{k}, 'IgnoreCase', true)
                    header_row = i;
                    keyword_found = true;
                    break;
                end
            end
            if keyword_found, break; end
        end
    end
    if keyword_found, break; end
end

if header_row == 0
    header_row = 1;
    fprintf('⚠ 헤더를 자동으로 찾을 수 없어 1행을 헤더로 가정합니다.\n');
else
    fprintf('✓ 헤더를 %d행에서 발견했습니다.\n', header_row);
end

% 헤더와 데이터 분리
headers = raw_data(header_row, :);
data_body = raw_data(header_row+1:end, :);

% 헤더 정보 출력
fprintf('\n헤더 정보:\n');
for i = 1:length(headers)
    if ischar(headers{i}) || isstring(headers{i})
        fprintf('  열 %d: %s\n', i, char(headers{i}));
    elseif isnumeric(headers{i}) && ~isnan(headers{i})
        fprintf('  열 %d: %.0f\n', i, headers{i});
    else
        fprintf('  열 %d: [빈값]\n', i);
    end
end

%% 4. ID 열과 발현역량 열 찾기 (개선된 로직)
fprintf('\n=== 열 매핑 ===\n');

id_col = 0;
eval_cols = [];

% ID 열 찾기
id_keywords = {'ID', '사번', 'id', 'ID', '직원번호', '사원번호'};
for i = 1:length(headers)
    if ischar(headers{i}) || isstring(headers{i})
        header_str = char(headers{i});
        for j = 1:length(id_keywords)
            if contains(header_str, id_keywords{j}, 'IgnoreCase', true)
                id_col = i;
                fprintf('✓ ID 열을 %d열에서 발견: %s\n', i, header_str);
                break;
            end
        end
        if id_col > 0, break; end
    end
end

if id_col == 0
    id_col = 1;
    fprintf('⚠ ID 열을 찾을 수 없어 1열을 ID로 가정합니다.\n');
end

% 발현역량 열 찾기
eval_keywords = {'발현', 'H1', 'H2', 'h1', 'h2', '역량', '평가'};
for i = 1:length(headers)
    if ischar(headers{i}) || isstring(headers{i})
        header_str = char(headers{i});
        for j = 1:length(eval_keywords)
            if contains(header_str, eval_keywords{j}, 'IgnoreCase', true)
                eval_cols = [eval_cols, i];
                fprintf('✓ 발현역량 열을 %d열에서 발견: %s\n', i, header_str);
                break;
            end
        end
    end
end

if isempty(eval_cols)
    % 발현역량 열을 찾을 수 없으면 ID 다음 5개 열을 가정
    eval_cols = (id_col+1):min(id_col+5, n_cols);
    fprintf('⚠ 발현역량 열을 자동으로 찾을 수 없어 %d~%d열을 사용합니다.\n', ...
            eval_cols(1), eval_cols(end));
end

fprintf('\n사용할 열 정보:\n');
fprintf('  ID 열: %d\n', id_col);
fprintf('  발현역량 열: %s\n', mat2str(eval_cols));

%% 5. 데이터 추출 및 정제
fprintf('\n=== 데이터 추출 및 정제 ===\n');

% 데이터 추출
employee_ids = data_body(:, id_col);
evaluation_data = data_body(:, eval_cols);

% ID 데이터 정제
if all(cellfun(@isnumeric, employee_ids))
    employee_ids = cell2mat(employee_ids);
else
    % ID가 문자열인 경우 숫자로 변환 시도
    try
        employee_ids = str2double(employee_ids);
        employee_ids = employee_ids(~isnan(employee_ids));
    catch
        fprintf('⚠ ID를 숫자로 변환할 수 없습니다. 문자열로 유지합니다.\n');
    end
end

% 유효한 데이터 필터링
valid_rows = true(size(data_body, 1), 1);
for i = 1:length(valid_rows)
    id_value = data_body{i, id_col};
    if isempty(id_value) || (isnumeric(id_value) && isnan(id_value)) || ...
       (ischar(id_value) && strcmp(strtrim(id_value), ''))
        valid_rows(i) = false;
    end
end

fprintf('데이터 필터링:\n');
fprintf('  전체 행: %d\n', size(data_body, 1));
fprintf('  유효한 행: %d\n', sum(valid_rows));

employee_ids = employee_ids(valid_rows);
evaluation_data = evaluation_data(valid_rows, :);

n_employees = length(employee_ids);
n_periods = size(evaluation_data, 2);

fprintf('\n✓ 데이터 로드 완료:\n');
fprintf('  - 직원 수: %d명\n', n_employees);
fprintf('  - 평가 기간: %d개\n', n_periods);

% 샘플 데이터 출력 (안전한 방식)
fprintf('\n샘플 데이터 (첫 3명):\n');
for i = 1:min(3, n_employees)
    fprintf('ID %s: ', num2str(employee_ids(i)));
    for j = 1:n_periods
        try
            cell_value = evaluation_data{i,j};
            if ischar(cell_value) || isstring(cell_value)
                fprintf('[%s] ', char(cell_value));
            elseif isnumeric(cell_value)
                if isnan(cell_value)
                    fprintf('[NaN] ');
                else
                    fprintf('[%.0f] ', cell_value);
                end
            elseif isa(cell_value, 'missing')
                fprintf('[Missing] ');
            elseif contains(class(cell_value), 'ActiveX')
                fprintf('[ActiveX_ERROR] ');
            else
                fprintf('[%s] ', class(cell_value));
            end
        catch
            fprintf('[ERROR] ');
        end
    end
    fprintf('\n');
end

% 평가 기간 레이블 생성
period_labels = {};
if n_periods == 5
    period_labels = {'23H1', '23H2', '24H1', '24H2', '25H1'};
else
    for i = 1:n_periods
        if i <= length(eval_cols) && (ischar(headers{eval_cols(i)}) || isstring(headers{eval_cols(i)}))
            period_labels{i} = char(headers{eval_cols(i)});
        else
            period_labels{i} = sprintf('Period%d', i);
        end
    end
end

fprintf('\n평가 기간 레이블: %s\n', strjoin(period_labels, ', '));

%% 6. 성장 점수 계산 (최적화된 버전)
fprintf('\n=== 성장 점수 계산 시작 ===\n');

% 결과 저장용 변수 초기화
growth_scores = zeros(n_employees, 1);
growth_patterns = cell(n_employees, 1);
score_details = cell(n_employees, 1);

% 프로그레스 바 설정
progress_interval = max(1, floor(n_employees / 20));
start_time = tic;

for i = 1:n_employees
    % 진행률 표시
    if mod(i, progress_interval) == 0 || i == 1 || i == n_employees
        elapsed_time = toc(start_time);
        if i > 1
            estimated_total = elapsed_time * n_employees / i;
            remaining_time = estimated_total - elapsed_time;
            fprintf('진행: %d/%d (%.1f%%) - 경과: %.1fs, 예상 남은 시간: %.1fs\n', ...
                    i, n_employees, i/n_employees*100, elapsed_time, remaining_time);
        else
            fprintf('진행: %d/%d (%.1f%%)\n', i, n_employees, i/n_employees*100);
        end
    end
    
    try
        employee_evals = evaluation_data(i, :);
        
        % 성장 점수 계산 (최적화된 로직)
        score = 0;
        base_level = '열린(Lv2)';
        consecutive_higher = 0;
        score_history = zeros(1, length(employee_evals));
        base_history = cell(1, length(employee_evals));
        
        for period = 1:length(employee_evals)
            current_eval = employee_evals{period};
            period_score = 0;
            
            % 데이터 정제 (강화된 버전)
            try
                if isnumeric(current_eval)
                    if isnan(current_eval)
                        current_eval = '#N/A';
                    else
                        current_eval = num2str(current_eval);
                    end
                elseif isa(current_eval, 'missing')
                    current_eval = '#N/A';
                elseif contains(class(current_eval), 'ActiveX')
                    current_eval = '#N/A';
                elseif isstring(current_eval)
                    current_eval = char(current_eval);
                elseif ~ischar(current_eval)
                    current_eval = char(current_eval);
                end
                
                % 공백 제거 및 대소문자 정규화
                current_eval = strtrim(current_eval);
                current_eval = upper(current_eval);  % 대소문자 통일
                
            catch
                current_eval = '#N/A';
            end
            
            % 평가 처리 (개선된 로직)
            if strcmp(current_eval, '#N/A') || isempty(current_eval) || ...
               strcmp(current_eval, 'NAN') || strcmp(current_eval, 'NA')
                period_score = 0;
                
            elseif contains(current_eval, '책임')
                if ~strcmp(base_level, '책임(Lv1)')
                    period_score = 30;
                    base_level = '책임(Lv1)';
                    consecutive_higher = 0;
                else
                    period_score = 0;
                end
                
            elseif contains(current_eval, '성취')
                if strcmp(base_level, '열린(Lv2)')
                    consecutive_higher = consecutive_higher + 1;
                    period_score = 10;
                    
                    if consecutive_higher >= 3
                        base_level = '성취(Lv1)';
                        consecutive_higher = 0;
                    end
                elseif strcmp(base_level, '성취(Lv1)') || strcmp(base_level, '성취(Lv2)')
                    period_score = 0;
                end
                
            elseif contains(current_eval, '열린(Lv2)')
                if strcmp(base_level, '열린(Lv2)')
                    period_score = 0;
                    if consecutive_higher > 0
                        consecutive_higher = 0;
                    end
                elseif contains(base_level, '성취')
                    period_score = -10;
                    consecutive_higher = 0;
                end
                
            elseif contains(current_eval, '열린(Lv1)')
                period_score = -5;
                consecutive_higher = 0;
            end
            
            score = score + period_score;
            score_history(period) = period_score;
            base_history{period} = base_level;
        end
        
        % 패턴 분류 (개선된 로직)
        if score >= 30
            pattern = '고성장';
        elseif score >= 20
            pattern = '중상성장';
        elseif score >= 10
            pattern = '중성장';
        elseif score > 0
            pattern = '저성장';
        elseif score == 0
            pattern = '정체';
        else
            pattern = '퇴보';
        end
        
        details.score_history = score_history;
        details.base_history = base_history;
        details.final_base_level = base_level;
        
        growth_scores(i) = score;
        growth_patterns{i} = pattern;
        score_details{i} = details;
        
    catch ME
        fprintf('⚠ 직원 %d 처리 중 에러: %s\n', i, ME.message);
        growth_scores(i) = 0;
        growth_patterns{i} = '에러';
        score_details{i} = struct('score_history', zeros(1, n_periods), ...
                                 'base_history', cell(1, n_periods), ...
                                 'final_base_level', '에러');
    end
end

total_time = toc(start_time);
fprintf('✓ 계산 완료! (총 소요시간: %.2f초)\n', total_time);

%% 7. 통계 분석 (강화된 버전)
fprintf('\n=== 통계 분석 ===\n');

% 유효한 데이터 필터링
valid_indices = growth_scores ~= 0 | ~cellfun(@(x) all(strcmp(x, '#N/A')), evaluation_data);
valid_scores = growth_scores(valid_indices);

% 데이터 유효성 검사
if isempty(valid_scores)
    fprintf('⚠ 경고: 유효한 데이터가 없습니다. 모든 직원의 점수가 0입니다.\n');
    valid_scores = growth_scores;
    valid_indices = true(size(growth_scores));
end

fprintf('유효한 데이터: %d명 (전체 %d명 중)\n', length(valid_scores), n_employees);

% 기본 통계 계산
stats = struct();
if ~isempty(valid_scores) && length(valid_scores) > 1
    stats.mean = mean(valid_scores);
    stats.median = median(valid_scores);
    stats.std = std(valid_scores);
    stats.min = min(valid_scores);
    stats.max = max(valid_scores);
    stats.quartiles = quantile(valid_scores, [0.25, 0.5, 0.75]);
    stats.range = stats.max - stats.min;
    stats.cv = stats.std / abs(stats.mean) * 100;  % 변동계수
else
    stats.mean = 0;
    stats.median = 0;
    stats.std = 0;
    stats.min = 0;
    stats.max = 0;
    stats.quartiles = [0, 0, 0];
    stats.range = 0;
    stats.cv = 0;
end

% 패턴별 집계
if ~isempty(growth_patterns(valid_indices))
    [unique_patterns, ~, pattern_indices] = unique(growth_patterns(valid_indices));
    pattern_counts = accumarray(pattern_indices, 1);
    pattern_percentages = pattern_counts / sum(valid_indices) * 100;
else
    unique_patterns = {'데이터없음'};
    pattern_counts = 0;
    pattern_percentages = 0;
end

%% 8. 결과 출력 (개선된 버전)
fprintf('\n========== MIDAS 성장단계 평가 분석 결과 ==========\n');
fprintf('분석 완료 시간: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('분석 대상: %d명 (유효 데이터: %d명)\n', n_employees, sum(valid_indices));

fprintf('\n[기초 통계]\n');
fprintf('  평균 점수: %.2f\n', stats.mean);
fprintf('  중앙값: %.2f\n', stats.median);
fprintf('  표준편차: %.2f\n', stats.std);
fprintf('  최소값: %.0f / 최대값: %.0f\n', stats.min, stats.max);
fprintf('  범위: %.0f\n', stats.range);
fprintf('  변동계수: %.1f%%\n', stats.cv);
fprintf('  사분위수: Q1=%.0f, Q2=%.0f, Q3=%.0f\n', ...
    stats.quartiles(1), stats.quartiles(2), stats.quartiles(3));

fprintf('\n[성장 패턴 분포]\n');
for i = 1:length(unique_patterns)
    fprintf('  %s: %d명 (%.1f%%)\n', ...
        unique_patterns{i}, pattern_counts(i), pattern_percentages(i));
end

%% 9. 시각화 (강화된 버전)
fprintf('\n=== 시각화 생성 ===\n');

try
    figure('Name', 'MIDAS 성장단계 평가 분석', 'Position', [100, 100, 1600, 900]);
    
    % 1) 점수 분포 히스토그램
    subplot(2, 4, 1);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        histogram(valid_scores, 15, 'FaceColor', [0.2 0.4 0.6], 'EdgeColor', 'white');
        xlabel('성장 점수');
        ylabel('인원수');
        title('성장 점수 분포');
        grid on;
        hold on;
        xline(stats.mean, 'r-', 'LineWidth', 2, 'DisplayName', '평균');
        xline(stats.median, 'g--', 'LineWidth', 1.5, 'DisplayName', '중앙값');
        legend('Location', 'best');
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('성장 점수 분포 (데이터 없음)');
    end

    % 2) 패턴별 분포 (파이 차트)
    subplot(2, 4, 2);
    if ~isempty(pattern_counts) && sum(pattern_counts) > 0
        pie(pattern_counts, unique_patterns);
        title('성장 패턴 분포');
        colormap(lines);
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('성장 패턴 분포 (데이터 없음)');
    end

    % 3) Box Plot
    subplot(2, 4, 3);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        boxplot(valid_scores, 'Labels', {'전체'});
        ylabel('성장 점수');
        title('성장 점수 Box Plot');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('성장 점수 Box Plot (데이터 없음)');
    end

    % 4) 누적 분포 함수
    subplot(2, 4, 4);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        [f, x] = ecdf(valid_scores);
        plot(x, f, 'LineWidth', 2, 'Color', [0.8 0.2 0.2]);
        xlabel('성장 점수');
        ylabel('누적 확률');
        title('누적 분포 함수 (CDF)');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('누적 분포 함수 (데이터 없음)');
    end

    % 5) 점수 구간별 분포
    subplot(2, 4, 5);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        score_ranges = [-inf, 0, 10, 20, 30, inf];
        score_labels = {'퇴보(<0)', '정체(0)', '저성장(1-10)', ...
                    '중성장(11-20)', '고성장(>20)'};
        score_groups = discretize(valid_scores, score_ranges);
        bar_counts = accumarray(score_groups, 1);
        bar(bar_counts, 'FaceColor', [0.3 0.6 0.3]);
        set(gca, 'XTickLabel', score_labels);
        xtickangle(45);
        ylabel('인원수');
        title('점수 구간별 분포');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('점수 구간별 분포 (데이터 없음)');
    end

    % 6) 상위 20명 성장 점수
    subplot(2, 4, 6);
    if ~isempty(growth_scores) && sum(growth_scores > 0) > 0
        [sorted_scores, sorted_idx] = sort(growth_scores, 'descend');
        top_n = min(20, sum(sorted_scores > 0));
        bar(1:top_n, sorted_scores(1:top_n), 'FaceColor', [0.9 0.6 0.1]);
        xlabel('순위');
        ylabel('성장 점수');
        title('상위 20명 성장 점수');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('상위 20명 성장 점수 (데이터 없음)');
    end

    % 7) 시간별 성장 추이 (새로운 차트)
    subplot(2, 4, 7);
    if n_periods > 1
        period_scores = zeros(1, n_periods);
        for p = 1:n_periods
            period_scores(p) = mean(cellfun(@(x) x(p), score_details));
        end
        plot(1:n_periods, period_scores, 'o-', 'LineWidth', 2, 'MarkerSize', 6);
        xlabel('평가 기간');
        ylabel('평균 점수');
        title('기간별 평균 성장 점수');
        set(gca, 'XTick', 1:n_periods, 'XTickLabel', period_labels);
        xtickangle(45);
        grid on;
    else
        text(0.5, 0.5, '기간 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('기간별 추이 (데이터 없음)');
    end

    % 8) 성장 패턴별 점수 분포 (새로운 차트)
    subplot(2, 4, 8);
    if length(unique_patterns) > 1
        pattern_scores = cell(length(unique_patterns), 1);
        for i = 1:length(unique_patterns)
            pattern_mask = strcmp(growth_patterns, unique_patterns{i});
            pattern_scores{i} = growth_scores(pattern_mask);
        end
        boxplot(cell2mat(pattern_scores), unique_patterns);
        ylabel('성장 점수');
        title('패턴별 점수 분포');
        xtickangle(45);
        grid on;
    else
        text(0.5, 0.5, '패턴 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
        title('패턴별 분포 (데이터 없음)');
    end
    
    fprintf('✓ 그래프 생성 완료!\n');
    
catch ME
    fprintf('⚠ 그래프 생성 중 에러: %s\n', ME.message);
    if ~isempty(ME.stack)
        fprintf('에러 위치: %s\n', ME.stack(1).name);
    end
end

%% 10. 결과 저장 (강화된 버전)
fprintf('\n=== 결과 저장 ===\n');

try
    % 기본 결과 테이블 생성
    result_table = table(employee_ids, growth_scores, growth_patterns, ...
        'VariableNames', {'ID', 'GrowthScore', 'Pattern'});

    % 엑셀 파일로 저장
    output_filename = sprintf('MIDAS_성장평가_결과_%s.xlsx', datestr(now, 'yyyymmdd_HHMMSS'));
    writetable(result_table, output_filename, 'WriteVariableNames', true);
    fprintf('✓ 기본 결과 저장: %s\n', output_filename);

    % 상세 결과 테이블 생성
    detailed_results = table();
    detailed_results.ID = employee_ids;
    
    % 각 기간별 발현역량 추가
    for p = 1:n_periods
        col_name = sprintf('Period%d_%s', p, period_labels{p});
        detailed_results.(col_name) = evaluation_data(:, p);
    end
    
    % 점수와 패턴 추가
    detailed_results.TotalScore = growth_scores;
    detailed_results.Pattern = growth_patterns;
    
    % 상세 결과 엑셀 파일로 저장
    detailed_output_filename = sprintf('MIDAS_성장평가_상세결과_%s.xlsx', datestr(now, 'yyyymmdd_HHMMSS'));
    writetable(detailed_results, detailed_output_filename, 'Sheet', '분석결과', 'WriteVariableNames', true);
    fprintf('✓ 상세 결과 저장: %s\n', detailed_output_filename);

    % MAT 파일로 저장
    mat_filename = sprintf('MIDAS_growth_analysis_%s.mat', datestr(now, 'yyyymmdd_HHMMSS'));
    save(mat_filename, 'result_table', 'detailed_results', 'stats', ...
         'score_details', 'evaluation_data', 'period_labels', 'n_periods');
    fprintf('✓ 데이터 파일 저장: %s\n', mat_filename);

    fprintf('\n✓ 모든 결과 저장 완료!\n');
    
catch ME
    fprintf('⚠ 결과 저장 중 에러: %s\n', ME.message);
    if ~isempty(ME.stack)
        fprintf('에러 위치: %s\n', ME.stack(1).name);
    end
end

%% 11. 개별 직원 상세 분석 (개선된 버전)
fprintf('\n=== 개별 직원 상세 분석 ===\n');

try
    analyze_individual = input('\n특정 직원 상세 분석을 원하시면 ID를 입력하세요 (종료: 0): ');
    
    if analyze_individual > 0
        idx = find(employee_ids == analyze_individual);
        if ~isempty(idx)
            fprintf('\n[ID %d 상세 분석]\n', analyze_individual);
            evals = evaluation_data(idx, :);
            
            fprintf('기간별 평가 결과:\n');
            for p = 1:length(period_labels)
                try
                    eval_str = char(evals{p});
                    fprintf('  %s: %s (점수: %+d)\n', period_labels{p}, ...
                            eval_str, score_details{idx}.score_history(p));
                catch
                    fprintf('  %s: [오류] (점수: %+d)\n', period_labels{p}, ...
                            score_details{idx}.score_history(p));
                end
            end
            fprintf('  총점: %d점 (패턴: %s)\n', ...
                    growth_scores(idx), growth_patterns{idx});
        else
            fprintf('해당 ID를 찾을 수 없습니다.\n');
        end
    end
catch ME
    fprintf('⚠ 개별 분석 중 에러: %s\n', ME.message);
end

%% 12. 프로그램 종료
fprintf('\n=== 프로그램 종료 ===\n');
fprintf('완료 시간: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('총 소요 시간: %.2f초\n', total_time);
fprintf('프로그램이 성공적으로 완료되었습니다.\n');
