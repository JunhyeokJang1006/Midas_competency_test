%% MIDAS 성장단계 평가 점수 산출 프로그램 (디버깅 버전)
% 작성일: 2024
% 목적: 최근 3년 입사자의 발현 역량 데이터를 기반으로 성장 점수 계산
% 개선사항: 세 번째 시트 직접 지정, 에러 처리 강화, 디버깅 정보 추가

clear; clc; close all;

%% 1. 데이터 불러오기 (디버깅 강화)
fprintf('=== MIDAS 성장단계 평가 분석 시작 ===\n\n');
fprintf('엑셀 파일 구조 확인 중...\n');

filename = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';

try
    % 파일 존재 여부 확인
    if ~exist(filename, 'file')
        error('파일을 찾을 수 없습니다: %s', filename);
    end
    
    % 엑셀 파일의 시트 정보 확인
    [status, sheets] = xlsfinfo(filename);
    
    if isempty(sheets)
        error('엑셀 파일에서 시트를 찾을 수 없습니다.');
    end
    
    fprintf('발견된 시트:\n');
    for i = 1:length(sheets)
        fprintf('  %d. %s\n', i, sheets{i});
    end
    
    % 세 번째 시트를 성장단계 시트로 직접 지정
    if length(sheets) >= 3
        sheet_name = sheets{3};
        fprintf('\n세 번째 시트를 성장단계 시트로 선택: %s\n', sheet_name);
    else
        error('세 번째 시트가 존재하지 않습니다. 총 %d개의 시트만 발견됨.', length(sheets));
    end
    
    fprintf('데이터 읽는 중...\n');
    
    % 선택된 시트에서 데이터 읽기 (readtable 사용으로 ActiveX 오류 방지)
    try
        % readtable을 사용하여 데이터 읽기 (원본 열 이름 보존)
        data_table = readtable(filename, 'Sheet', sheet_name, 'VariableNamingRule', 'preserve');
        fprintf('readtable로 데이터 읽기 성공\n');
        
        % 테이블을 셀 배열로 변환
        raw_data = table2cell(data_table);
        
    catch ME1
        fprintf('readtable 실패, xlsread 시도: %s\n', ME1.message);
        try
            % xlsread 사용 (기존 방법)
            [num_data, txt_data, raw_data] = xlsread(filename, sheet_name);
        catch ME2
            fprintf('xlsread도 실패: %s\n', ME2.message);
            error('데이터 읽기 실패: readtable과 xlsread 모두 실패');
        end
    end
    
    % 데이터 구조 파악
    [n_rows, n_cols] = size(raw_data);
    fprintf('데이터 크기: %d행 x %d열\n', n_rows, n_cols);
    
    % 디버깅: 첫 몇 행 출력
    fprintf('\n데이터 미리보기 (첫 5행):\n');
    for i = 1:min(5, n_rows)
        fprintf('행 %d: ', i);
        for j = 1:min(10, n_cols)  % 처음 10열만 출력
            try
                if ischar(raw_data{i,j})
                    fprintf('[%s] ', raw_data{i,j});
                elseif isnumeric(raw_data{i,j})
                    if isnan(raw_data{i,j})
                        fprintf('[NaN] ');
                    else
                        fprintf('[%.0f] ', raw_data{i,j});
                    end
                elseif isa(raw_data{i,j}, 'missing')
                    fprintf('[Missing] ');
                elseif contains(class(raw_data{i,j}), 'ActiveX')
                    fprintf('[ActiveX_ERROR] ');
                else
                    fprintf('[%s] ', class(raw_data{i,j}));
                end
            catch
                fprintf('[ERROR] ');
            end
        end
        fprintf('\n');
    end
    
    % 헤더 찾기 (발현역량 관련 열 확인)
    header_row = 0;
    for i = 1:min(10, n_rows)  % 처음 10행 내에서 헤더 찾기
        row_data = raw_data(i, :);
        if any(cellfun(@(x) ischar(x) && contains(x, '발현'), row_data)) || ...
           any(cellfun(@(x) ischar(x) && contains(x, 'H1'), row_data)) || ...
           any(cellfun(@(x) ischar(x) && contains(x, 'H2'), row_data)) || ...
           any(cellfun(@(x) ischar(x) && contains(x, 'ID'), row_data)) || ...
           any(cellfun(@(x) ischar(x) && contains(x, '사번'), row_data))
            header_row = i;
            fprintf('헤더를 %d행에서 발견했습니다.\n', i);
            break;
        end
    end
    
    if header_row == 0
        header_row = 1;  % 기본값
        fprintf('헤더를 자동으로 찾을 수 없어 1행을 헤더로 가정합니다.\n');
    end
    
    headers = raw_data(header_row, :);
    data_body = raw_data(header_row+1:end, :);
    
    % 디버깅: 헤더 정보 출력
    fprintf('\n헤더 정보:\n');
    for i = 1:length(headers)
        if ischar(headers{i})
            fprintf('  열 %d: %s\n', i, headers{i});
        elseif isnumeric(headers{i}) && ~isnan(headers{i})
            fprintf('  열 %d: %.0f\n', i, headers{i});
        else
            fprintf('  열 %d: [빈값]\n', i);
        end
    end
    
    % ID 열과 발현역량 열 찾기
    id_col = 0;
    eval_cols = [];
    
    for i = 1:length(headers)
        if ischar(headers{i})
            if contains(headers{i}, 'ID') || contains(headers{i}, '사번') || ...
               contains(headers{i}, 'id') || contains(headers{i}, 'ID')
                id_col = i;
                fprintf('ID 열을 %d열에서 발견: %s\n', i, headers{i});
            elseif contains(headers{i}, '발현') || contains(headers{i}, 'H1') || ...
                   contains(headers{i}, 'H2') || contains(headers{i}, 'h1') || ...
                   contains(headers{i}, 'h2')
                eval_cols = [eval_cols, i];
                fprintf('발현역량 열을 %d열에서 발견: %s\n', i, headers{i});
            end
        end
    end
    
    if id_col == 0
        id_col = 1;  % 첫 번째 열을 ID로 가정
        fprintf('ID 열을 찾을 수 없어 1열을 ID로 가정합니다.\n');
    end
    
    if isempty(eval_cols)
        % 발현역량 열을 찾을 수 없으면 ID 다음 5개 열을 가정
        eval_cols = (id_col+1):min(id_col+5, n_cols);
        fprintf('발현역량 열을 자동으로 찾을 수 없어 %d~%d열을 사용합니다.\n', ...
                eval_cols(1), eval_cols(end));
    end
    
    fprintf('\n사용할 열 정보:\n');
    fprintf('  ID 열: %d\n', id_col);
    fprintf('  발현역량 열: %s\n', mat2str(eval_cols));
    
    % 데이터 추출
    employee_ids = data_body(:, id_col);
    evaluation_data = data_body(:, eval_cols);
    
    % ID가 숫자인 경우 처리
    if all(cellfun(@isnumeric, employee_ids))
        employee_ids = cell2mat(employee_ids);
    end
    
    % 유효한 데이터 필터링 (ID가 있는 행만)
    valid_rows = ~cellfun(@(x) isempty(x) || (isnumeric(x) && isnan(x)), ...
                          data_body(:, id_col));
    
    fprintf('\n데이터 필터링:\n');
    fprintf('  전체 행: %d\n', size(data_body, 1));
    fprintf('  유효한 행: %d\n', sum(valid_rows));
    
    employee_ids = employee_ids(valid_rows);
    evaluation_data = evaluation_data(valid_rows, :);
    
    n_employees = length(employee_ids);
    n_periods = size(evaluation_data, 2);
    
    fprintf('\n데이터 로드 완료:\n');
    fprintf('  - 직원 수: %d명\n', n_employees);
    fprintf('  - 평가 기간: %d개\n', n_periods);
    
    % 디버깅: 샘플 데이터 출력
    fprintf('\n샘플 데이터 (첫 3명):\n');
    for i = 1:min(3, n_employees)
        fprintf('ID %s: ', num2str(employee_ids(i)));
        for j = 1:n_periods
            try
                if ischar(evaluation_data{i,j})
                    fprintf('[%s] ', evaluation_data{i,j});
                elseif isnumeric(evaluation_data{i,j})
                    if isnan(evaluation_data{i,j})
                        fprintf('[NaN] ');
                    else
                        fprintf('[%.0f] ', evaluation_data{i,j});
                    end
                elseif isa(evaluation_data{i,j}, 'missing')
                    fprintf('[Missing] ');
                elseif contains(class(evaluation_data{i,j}), 'ActiveX')
                    fprintf('[ActiveX_ERROR] ');
                else
                    fprintf('[%s] ', class(evaluation_data{i,j}));
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
            if i <= length(eval_cols) && ischar(headers{eval_cols(i)})
                period_labels{i} = headers{eval_cols(i)};
            else
                period_labels{i} = sprintf('Period%d', i);
            end
        end
    end
    
    fprintf('\n평가 기간 레이블: %s\n', strjoin(period_labels, ', '));
    
catch ME
    fprintf('\n에러 발생!\n');
    fprintf('에러 메시지: %s\n', ME.message);
    fprintf('에러 위치: %s (라인 %d)\n', ME.stack(1).name, ME.stack(1).line);
    
    % 상세 에러 정보 출력
    if length(ME.stack) > 1
        fprintf('\n스택 트레이스:\n');
        for i = 1:length(ME.stack)
            fprintf('  %d. %s (라인 %d)\n', i, ME.stack(i).name, ME.stack(i).line);
        end
    end
    
    error('파일 읽기 오류: %s', ME.message);
end

%% 2. 성장 점수 계산 함수 정의
% 함수를 별도로 정의하지 않고 인라인으로 처리

%% 3. 전체 직원 분석
fprintf('\n성장 점수 계산 시작...\n');

growth_scores = zeros(n_employees, 1);
growth_patterns = cell(n_employees, 1);
score_details = cell(n_employees, 1);

% 프로그레스 바
progress_interval = max(1, floor(n_employees / 20));

for i = 1:n_employees
    if mod(i, progress_interval) == 0 || i == 1 || i == n_employees
        fprintf('진행: %d/%d (%.1f%%)\n', i, n_employees, i/n_employees*100);
    end
    
    try
        employee_evals = evaluation_data(i, :);
        
        % 성장 점수 계산 (인라인)
        score = 0;
        base_level = '열린(Lv2)';
        consecutive_higher = 0;
        score_history = zeros(1, length(employee_evals));
        base_history = cell(1, length(employee_evals));
        
        for period = 1:length(employee_evals)
            current_eval = employee_evals{period};
            period_score = 0;
            
            % 데이터 정제 (ActiveX 오류 처리 포함)
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
                elseif ~ischar(current_eval)
                    current_eval = char(current_eval);
                end
                
                % 공백 제거
                current_eval = strtrim(current_eval);
            catch
                current_eval = '#N/A';
            end
            
            % 평가 처리
            if strcmp(current_eval, '#N/A') || isempty(current_eval) || ...
               strcmp(current_eval, 'NaN') || strcmp(current_eval, 'NA')
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
        
        % 패턴 분류
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
        fprintf('직원 %d 처리 중 에러: %s\n', i, ME.message);
        growth_scores(i) = 0;
        growth_patterns{i} = '에러';
        score_details{i} = struct('score_history', zeros(1, n_periods), ...
                                 'base_history', cell(1, n_periods), ...
                                 'final_base_level', '에러');
    end
end

fprintf('계산 완료!\n');

%% 4. 통계 분석
fprintf('\n통계 분석 중...\n');

% 유효한 데이터만 필터링 (입사 전 기간 제외)
valid_indices = growth_scores ~= 0 | ~cellfun(@(x) all(strcmp(x, '#N/A')), evaluation_data);
valid_scores = growth_scores(valid_indices);

% 유효한 데이터가 있는지 확인
if isempty(valid_scores)
    fprintf('경고: 유효한 데이터가 없습니다. 모든 직원의 점수가 0입니다.\n');
    valid_scores = growth_scores;  % 모든 데이터 사용
    valid_indices = true(size(growth_scores));  % 모든 인덱스 유효로 설정
end

fprintf('유효한 데이터: %d명\n', length(valid_scores));

% 기본 통계
stats = struct();
if ~isempty(valid_scores)
    stats.mean = mean(valid_scores);
    stats.median = median(valid_scores);
    stats.std = std(valid_scores);
    stats.min = min(valid_scores);
    stats.max = max(valid_scores);
    stats.quartiles = quantile(valid_scores, [0.25, 0.5, 0.75]);
else
    stats.mean = 0;
    stats.median = 0;
    stats.std = 0;
    stats.min = 0;
    stats.max = 0;
    stats.quartiles = [0, 0, 0];
end

% 패턴별 집계
if ~isempty(growth_patterns(valid_indices))
    [unique_patterns, ~, pattern_indices] = unique(growth_patterns(valid_indices));
    pattern_counts = accumarray(pattern_indices, 1);
else
    unique_patterns = {'데이터없음'};
    pattern_counts = 0;
end

%% 5. 결과 출력
fprintf('\n========== MIDAS 성장단계 평가 분석 결과 ==========\n');
fprintf('분석 대상: %d명 (유효 데이터: %d명)\n', n_employees, sum(valid_indices));
fprintf('\n[기초 통계]\n');
fprintf('  평균 점수: %.2f\n', stats.mean);
fprintf('  중앙값: %.2f\n', stats.median);
fprintf('  표준편차: %.2f\n', stats.std);
fprintf('  최소값: %.0f / 최대값: %.0f\n', stats.min, stats.max);
fprintf('  사분위수: Q1=%.0f, Q2=%.0f, Q3=%.0f\n', ...
    stats.quartiles(1), stats.quartiles(2), stats.quartiles(3));

fprintf('\n[성장 패턴 분포]\n');
for i = 1:length(unique_patterns)
    fprintf('  %s: %d명 (%.1f%%)\n', ...
        unique_patterns{i}, pattern_counts(i), ...
        pattern_counts(i)/sum(valid_indices)*100);
end

%% 6. 시각화
fprintf('\n그래프 생성 중...\n');

try
    figure('Name', 'MIDAS 성장단계 평가 분석', 'Position', [100, 100, 1400, 800]);

    % 1) 점수 분포 히스토그램
    subplot(2, 3, 1);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        histogram(valid_scores, 15, 'FaceColor', [0.2 0.4 0.6]);
        xlabel('성장 점수');
        ylabel('인원수');
        title('성장 점수 분포');
        grid on;
        hold on;
        xline(stats.mean, 'r-', 'LineWidth', 2);
        xline(stats.median, 'g--', 'LineWidth', 1.5);
        legend('분포', '평균', '중앙값', 'Location', 'best');
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('성장 점수 분포 (데이터 없음)');
    end

    % 2) 패턴별 분포 (파이 차트)
    subplot(2, 3, 2);
    if ~isempty(pattern_counts) && sum(pattern_counts) > 0
        pie(pattern_counts, unique_patterns);
        title('성장 패턴 분포');
        colormap(autumn);
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('성장 패턴 분포 (데이터 없음)');
    end

    % 3) Box Plot
    subplot(2, 3, 3);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        boxplot(valid_scores, 'Labels', {'전체'});
        ylabel('성장 점수');
        title('성장 점수 Box Plot');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('성장 점수 Box Plot (데이터 없음)');
    end

    % 4) 누적 분포 함수
    subplot(2, 3, 4);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        [f, x] = ecdf(valid_scores);
        plot(x, f, 'LineWidth', 2);
        xlabel('성장 점수');
        ylabel('누적 확률');
        title('누적 분포 함수 (CDF)');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('누적 분포 함수 (데이터 없음)');
    end

    % 5) 점수 구간별 분포
    subplot(2, 3, 5);
    if ~isempty(valid_scores) && length(valid_scores) > 1
        score_ranges = [-inf, 0, 10, 20, 30, inf];
        score_labels = {'퇴보(<0)', '정체(0)', '저성장(1-10)', ...
                    '중성장(11-20)', '고성장(>20)'};
        score_groups = discretize(valid_scores, score_ranges);
        bar_counts = accumarray(score_groups, 1);
        bar(bar_counts);
        set(gca, 'XTickLabel', score_labels);
        xtickangle(45);
        ylabel('인원수');
        title('점수 구간별 분포');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('점수 구간별 분포 (데이터 없음)');
    end

    % 6) 상위 20명 성장 점수
    subplot(2, 3, 6);
    if ~isempty(growth_scores) && sum(growth_scores > 0) > 0
        [sorted_scores, sorted_idx] = sort(growth_scores, 'descend');
        top_n = min(20, sum(sorted_scores > 0));
        bar(1:top_n, sorted_scores(1:top_n));
        xlabel('순위');
        ylabel('성장 점수');
        title('상위 20명 성장 점수');
        grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('상위 20명 성장 점수 (데이터 없음)');
    end
    
    fprintf('그래프 생성 완료!\n');
    
catch ME
    fprintf('그래프 생성 중 에러: %s\n', ME.message);
    fprintf('에러 위치: %s\n', ME.stack(1).name);
end

%% 7. 결과 저장
fprintf('\n결과 저장 중...\n');

try
    % 결과 테이블 생성
    result_table = table(employee_ids, growth_scores, growth_patterns, ...
        'VariableNames', {'ID', 'GrowthScore', 'Pattern'});

    % 엑셀 파일로 저장 (원본 열 이름 보존)
    output_filename = 'MIDAS_성장평가_결과.xlsx';
    writetable(result_table, output_filename, 'WriteVariableNames', true);

    % MAT 파일로 저장
    save('MIDAS_growth_analysis.mat', 'result_table', 'stats', ...
         'score_details', 'evaluation_data');

    fprintf('\n분석 완료!\n');
    fprintf('결과 파일: %s\n', output_filename);
    fprintf('데이터 파일: MIDAS_growth_analysis.mat\n');
    
catch ME
    fprintf('결과 저장 중 에러: %s\n', ME.message);
end

%% 8. 선택적: 개별 직원 상세 분석
% 특정 직원의 성장 궤적을 자세히 보고 싶을 때
try
    analyze_individual = input('\n특정 직원 상세 분석을 원하시면 ID를 입력하세요 (종료: 0): ');
    
    if analyze_individual > 0
        idx = find(employee_ids == analyze_individual);
        if ~isempty(idx)
            fprintf('\n[ID %d 상세 분석]\n', analyze_individual);
            evals = evaluation_data(idx, :);
            
            for p = 1:length(period_labels)
                fprintf('  %s: %s (점수: %+d)\n', period_labels{p}, ...
                        char(evals{p}), score_details{idx}.score_history(p));
            end
            fprintf('  총점: %d점 (패턴: %s)\n', ...
                    growth_scores(idx), growth_patterns{idx});
        else
            fprintf('해당 ID를 찾을 수 없습니다.\n');
        end
    end
catch ME
    fprintf('개별 분석 중 에러: %s\n', ME.message);
end

fprintf('\n프로그램 종료.\n');