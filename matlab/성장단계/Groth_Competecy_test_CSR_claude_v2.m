%% MIDAS 성장단계 평가 점수 산출 프로그램 (Claude v2.0)
% 작성일: 2024
% 작성자: Claude AI Assistant
% 목적: 최근 3년 입사자의 발현 역량 데이터를 기반으로 성장 점수 계산
% 특징: 안정성과 호환성을 최우선으로 하는 단일 스크립트

clear; clc; close all;

%% 1. 프로그램 시작
fprintf('================================================\n');
fprintf('    MIDAS 성장단계 평가 분석 시스템 v2.0\n');
fprintf('================================================\n');
fprintf('시작 시간: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

start_time = tic;

%% 2. 데이터 파일 설정 및 확인
data_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';

fprintf('데이터 파일 확인 중...\n');
fprintf('파일 경로: %s\n', data_file);

if ~exist(data_file, 'file')
    error('❌ 데이터 파일을 찾을 수 없습니다: %s', data_file);
end
fprintf('✅ 데이터 파일 확인 완료\n\n');

%% 3. 엑셀 파일 시트 정보 확인
try
    fprintf('엑셀 파일 구조 분석 중...\n');
    [~, sheet_names] = xlsfinfo(data_file);

    if isempty(sheet_names)
        error('엑셀 파일에서 시트를 찾을 수 없습니다.');
    end

    fprintf('발견된 시트 (%d개):\n', length(sheet_names));
    for i = 1:length(sheet_names)
        fprintf('  %d. %s\n', i, sheet_names{i});
    end

    % 세 번째 시트 선택
    if length(sheet_names) >= 3
        target_sheet = sheet_names{3};
        fprintf('\n✅ 분석 대상 시트: %s\n', target_sheet);
    else
        error('세 번째 시트가 존재하지 않습니다.');
    end

catch ME
    error('엑셀 파일 분석 실패: %s', ME.message);
end

%% 4. 데이터 읽기
fprintf('\n데이터 읽기 중...\n');

try
    % readtable 시도
    try
        opts = detectImportOptions(data_file, 'Sheet', target_sheet);
        opts.VariableNamingRule = 'preserve';
        data_table = readtable(data_file, opts);
        raw_data = table2cell(data_table);
        fprintf('✅ readtable로 데이터 읽기 성공\n');
    catch
        % xlsread 대안
        [~, ~, raw_data] = xlsread(data_file, target_sheet);
        fprintf('✅ xlsread로 데이터 읽기 성공\n');
    end

    [n_rows, n_cols] = size(raw_data);
    fprintf('데이터 크기: %d행 × %d열\n', n_rows, n_cols);

    if n_rows < 2 || n_cols < 2
        error('데이터가 부족합니다.');
    end

catch ME
    error('데이터 읽기 실패: %s', ME.message);
end

%% 5. 헤더 탐지 및 데이터 구조 파악
fprintf('\n데이터 구조 분석 중...\n');

% 헤더 행 찾기
header_row = 1;
header_keywords = {'발현', 'H1', 'H2', 'ID', '사번'};

for i = 1:min(5, n_rows)
    keyword_count = 0;
    for j = 1:n_cols
        if iscell(raw_data) && i <= size(raw_data,1) && j <= size(raw_data,2)
            cell_val = raw_data{i,j};
            if ischar(cell_val) || isstring(cell_val)
                for k = 1:length(header_keywords)
                    if contains(char(cell_val), header_keywords{k}, 'IgnoreCase', true)
                        keyword_count = keyword_count + 1;
                        break;
                    end
                end
            end
        end
    end

    if keyword_count >= 2
        header_row = i;
        fprintf('✅ 헤더를 %d행에서 발견\n', i);
        break;
    end
end

% 헤더와 데이터 분리
headers = raw_data(header_row, :);
data_rows = raw_data(header_row+1:end, :);

%% 6. 열 매핑
fprintf('\n데이터 열 매핑 중...\n');

% ID 열 찾기
id_col = 1;  % 기본값
for i = 1:length(headers)
    if iscell(headers) && i <= length(headers)
        header_val = headers{i};
        if (ischar(header_val) || isstring(header_val))
            if contains(char(header_val), {'ID', '사번', 'id'}, 'IgnoreCase', true)
                id_col = i;
                fprintf('✅ ID 열 발견: %d열 (%s)\n', i, char(header_val));
                break;
            end
        end
    end
end

% 발현역량 열 찾기
eval_cols = [];
for i = 1:length(headers)
    if iscell(headers) && i <= length(headers)
        header_val = headers{i};
        if (ischar(header_val) || isstring(header_val))
            if contains(char(header_val), {'발현', 'H1', 'H2', 'h1', 'h2'}, 'IgnoreCase', true)
                eval_cols = [eval_cols, i];
                fprintf('✅ 발현역량 열 발견: %d열 (%s)\n', i, char(header_val));
            end
        end
    end
end

% 발현역량 열이 없으면 기본 범위 설정
if isempty(eval_cols)
    eval_cols = (id_col+1):min(id_col+5, n_cols);
    fprintf('⚠ 기본 범위 사용: %d~%d열\n', eval_cols(1), eval_cols(end));
end

fprintf('사용할 열: ID=%d, 발현역량=%s\n', id_col, mat2str(eval_cols));

%% 7. 데이터 추출 및 정제
fprintf('\n데이터 추출 및 정제 중...\n');

% 데이터 추출
employee_ids = data_rows(:, id_col);
evaluation_data = data_rows(:, eval_cols);

% 유효한 직원만 필터링
valid_mask = false(size(employee_ids));
for i = 1:length(employee_ids)
    if iscell(employee_ids) && i <= length(employee_ids)
        id_val = employee_ids{i};
        if isnumeric(id_val) && ~isnan(id_val)
            valid_mask(i) = true;
        elseif (ischar(id_val) || isstring(id_val)) && ~isempty(strtrim(char(id_val)))
            valid_mask(i) = true;
        end
    end
end

employee_ids = employee_ids(valid_mask);
evaluation_data = evaluation_data(valid_mask, :);

% ID를 숫자로 변환
numeric_ids = zeros(length(employee_ids), 1);
for i = 1:length(employee_ids)
    if iscell(employee_ids) && i <= length(employee_ids)
        id_val = employee_ids{i};
        if isnumeric(id_val)
            numeric_ids(i) = id_val;
        else
            try
                numeric_ids(i) = str2double(char(id_val));
            catch
                numeric_ids(i) = i;  % 변환 실패시 인덱스 사용
            end
        end
    end
end
employee_ids = numeric_ids;

n_employees = length(employee_ids);
n_periods = size(evaluation_data, 2);

fprintf('✅ 최종 분석 대상: %d명, %d개 기간\n', n_employees, n_periods);

% 기간 레이블 생성
if n_periods == 5
    period_labels = {'23H1', '23H2', '24H1', '24H2', '25H1'};
else
    period_labels = cell(1, n_periods);
    for i = 1:n_periods
        period_labels{i} = sprintf('Period%d', i);
    end
end

%% 8. 성장 점수 계산
fprintf('\n성장 점수 계산 시작...\n');

growth_scores = zeros(n_employees, 1);
growth_patterns = cell(n_employees, 1);
calculation_details = cell(n_employees, 1);

% 진행률 표시 간격
progress_step = max(1, floor(n_employees / 10));

for emp_idx = 1:n_employees
    % 진행률 표시
    if mod(emp_idx, progress_step) == 0 || emp_idx == 1 || emp_idx == n_employees
        fprintf('  진행: %d/%d (%.1f%%)\n', emp_idx, n_employees, emp_idx/n_employees*100);
    end

    try
        % 개별 직원 평가 데이터
        emp_evaluations = evaluation_data(emp_idx, :);

        % 성장 점수 계산
        total_score = 0;
        current_level = '열린(Lv2)';
        consecutive_achievements = 0;
        score_per_period = zeros(1, n_periods);
        level_per_period = cell(1, n_periods);

        for period_idx = 1:n_periods
            period_score = 0;

            % 평가 데이터 정제
            if iscell(emp_evaluations) && period_idx <= size(emp_evaluations, 2)
                eval_raw = emp_evaluations{period_idx};
            else
                eval_raw = [];
            end

            % 평가 텍스트 변환
            if isnumeric(eval_raw) && isnan(eval_raw)
                eval_text = '';
            elseif isnumeric(eval_raw)
                eval_text = num2str(eval_raw);
            elseif ischar(eval_raw) || isstring(eval_raw)
                eval_text = char(eval_raw);
            else
                eval_text = '';
            end

            eval_text = strtrim(eval_text);

            % 점수 계산 로직
            if isempty(eval_text) || strcmp(eval_text, 'NaN') || strcmp(eval_text, '#N/A')
                period_score = 0;  % 평가 없음

            elseif contains(eval_text, '책임', 'IgnoreCase', true)
                if ~strcmp(current_level, '책임(Lv1)')
                    period_score = 30;
                    current_level = '책임(Lv1)';
                    consecutive_achievements = 0;
                end

            elseif contains(eval_text, '성취', 'IgnoreCase', true)
                if strcmp(current_level, '열린(Lv2)')
                    consecutive_achievements = consecutive_achievements + 1;
                    period_score = 10;
                    if consecutive_achievements >= 3
                        current_level = '성취(Lv1)';
                        consecutive_achievements = 0;
                    end
                end

            elseif contains(eval_text, '열린', 'IgnoreCase', true)
                if contains(eval_text, 'Lv2', 'IgnoreCase', true)
                    if strcmp(current_level, '열린(Lv2)')
                        period_score = 0;
                        consecutive_achievements = 0;
                    elseif contains(current_level, '성취')
                        period_score = -10;
                        consecutive_achievements = 0;
                    end
                elseif contains(eval_text, 'Lv1', 'IgnoreCase', true)
                    period_score = -5;
                    consecutive_achievements = 0;
                end
            end

            total_score = total_score + period_score;
            score_per_period(period_idx) = period_score;
            level_per_period{period_idx} = current_level;
        end

        % 성장 패턴 분류
        if total_score >= 30
            pattern = '고성장';
        elseif total_score >= 20
            pattern = '중상성장';
        elseif total_score >= 10
            pattern = '중성장';
        elseif total_score > 0
            pattern = '저성장';
        elseif total_score == 0
            pattern = '정체';
        else
            pattern = '퇴보';
        end

        % 결과 저장
        growth_scores(emp_idx) = total_score;
        growth_patterns{emp_idx} = pattern;
        calculation_details{emp_idx} = struct('score_history', score_per_period, ...
                                            'level_history', {level_per_period}, ...
                                            'final_level', current_level);

    catch ME
        fprintf('⚠ 직원 %d 처리 중 오류: %s\n', emp_idx, ME.message);
        growth_scores(emp_idx) = 0;
        growth_patterns{emp_idx} = '오류';
    end
end

fprintf('✅ 성장 점수 계산 완료\n');

%% 9. 통계 분석
fprintf('\n통계 분석 중...\n');

% 유효한 데이터 필터링
valid_scores = growth_scores(~strcmp(growth_patterns, '오류'));
valid_patterns = growth_patterns(~strcmp(growth_patterns, '오류'));

% 기본 통계
if ~isempty(valid_scores)
    stats_mean = mean(valid_scores);
    stats_median = median(valid_scores);
    stats_std = std(valid_scores);
    stats_min = min(valid_scores);
    stats_max = max(valid_scores);

    if length(valid_scores) > 1
        stats_q = quantile(valid_scores, [0.25, 0.5, 0.75]);
    else
        stats_q = [stats_mean, stats_mean, stats_mean];
    end
else
    stats_mean = 0; stats_median = 0; stats_std = 0;
    stats_min = 0; stats_max = 0; stats_q = [0, 0, 0];
end

% 패턴별 분포
[unique_patterns, ~, pattern_idx] = unique(valid_patterns);
pattern_counts = zeros(length(unique_patterns), 1);
for i = 1:length(unique_patterns)
    pattern_counts(i) = sum(strcmp(valid_patterns, unique_patterns{i}));
end

fprintf('✅ 통계 분석 완료\n');

%% 10. 결과 출력
fprintf('\n');
fprintf('================================================\n');
fprintf('                분석 결과\n');
fprintf('================================================\n');
fprintf('완료 시간: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('분석 대상: %d명\n', n_employees);

fprintf('\n[기본 통계]\n');
fprintf('  평균 점수: %.2f\n', stats_mean);
fprintf('  중앙값: %.2f\n', stats_median);
fprintf('  표준편차: %.2f\n', stats_std);
fprintf('  최솟값: %.0f / 최댓값: %.0f\n', stats_min, stats_max);
fprintf('  사분위수: Q1=%.0f, Q2=%.0f, Q3=%.0f\n', stats_q(1), stats_q(2), stats_q(3));

fprintf('\n[성장 패턴 분포]\n');
for i = 1:length(unique_patterns)
    percentage = (pattern_counts(i) / length(valid_patterns)) * 100;
    fprintf('  %s: %d명 (%.1f%%)\n', unique_patterns{i}, pattern_counts(i), percentage);
end

%% 11. 시각화
fprintf('\n시각화 생성 중...\n');

try
    figure('Name', 'MIDAS 성장단계 평가 분석', 'Position', [100, 100, 1200, 800]);

    % 1) 점수 분포 히스토그램
    subplot(2, 3, 1);
    if length(valid_scores) > 1
        histogram(valid_scores, 10, 'FaceColor', [0.3, 0.5, 0.8]);
        hold on;
        xline(stats_mean, 'r-', 'LineWidth', 2);
        xline(stats_median, 'g--', 'LineWidth', 2);
        xlabel('성장 점수'); ylabel('인원수'); title('성장 점수 분포');
        legend('분포', '평균', '중앙값'); grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('성장 점수 분포');
    end

    % 2) 패턴별 분포
    subplot(2, 3, 2);
    if ~isempty(pattern_counts) && sum(pattern_counts) > 0
        pie(pattern_counts, unique_patterns);
        title('성장 패턴 분포');
    else
        text(0.5, 0.5, '데이터 없음', 'HorizontalAlignment', 'center');
        title('성장 패턴 분포');
    end

    % 3) 박스 플롯
    subplot(2, 3, 3);
    if length(valid_scores) > 1
        boxplot(valid_scores);
        ylabel('성장 점수'); title('성장 점수 Box Plot');
        set(gca, 'XTickLabel', {'전체'}); grid on;
    else
        text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
        title('성장 점수 Box Plot');
    end

    % 4) 구간별 분포
    subplot(2, 3, 4);
    if length(valid_scores) > 0
        score_ranges = [-inf, 0, 10, 20, 30, inf];
        score_labels = {'퇴보', '정체', '저성장', '중성장', '고성장'};
        [counts, ~] = histcounts(valid_scores, score_ranges);
        bar(counts, 'FaceColor', [0.6, 0.8, 0.4]);
        set(gca, 'XTickLabel', score_labels);
        ylabel('인원수'); title('구간별 분포');
        xtickangle(45); grid on;
    else
        text(0.5, 0.5, '데이터 없음', 'HorizontalAlignment', 'center');
        title('구간별 분포');
    end

    % 5) 상위 성과자
    subplot(2, 3, 5);
    [sorted_scores, ~] = sort(growth_scores, 'descend');
    top_n = min(15, sum(sorted_scores > 0));
    if top_n > 0
        bar(1:top_n, sorted_scores(1:top_n), 'FaceColor', [0.9, 0.6, 0.2]);
        xlabel('순위'); ylabel('성장 점수'); title('상위 15명');
        grid on;
    else
        text(0.5, 0.5, '데이터 없음', 'HorizontalAlignment', 'center');
        title('상위 15명');
    end

    % 6) 기간별 평균
    subplot(2, 3, 6);
    if n_periods > 1
        period_avg = zeros(1, n_periods);
        for p = 1:n_periods
            period_scores = zeros(n_employees, 1);
            for e = 1:n_employees
                if ~isempty(calculation_details{e})
                    period_scores(e) = calculation_details{e}.score_history(p);
                end
            end
            period_avg(p) = mean(period_scores);
        end
        plot(1:n_periods, period_avg, 'o-', 'LineWidth', 2, 'MarkerSize', 6);
        xlabel('기간'); ylabel('평균 점수'); title('기간별 평균');
        set(gca, 'XTick', 1:n_periods, 'XTickLabel', period_labels);
        xtickangle(45); grid on;
    else
        text(0.5, 0.5, '기간 부족', 'HorizontalAlignment', 'center');
        title('기간별 평균');
    end

    fprintf('✅ 시각화 완료\n');

catch ME
    fprintf('⚠ 시각화 생성 오류: %s\n', ME.message);
end

%% 12. 결과 저장
fprintf('\n결과 저장 중...\n');

try
    timestamp = datestr(now, 'yyyymmdd_HHMMSS');

    % 기본 결과 테이블
    result_table = array2table([employee_ids, growth_scores], ...
                              'VariableNames', {'ID', 'GrowthScore'});
    result_table.Pattern = growth_patterns;

    % 상세 결과 테이블
    detailed_table = array2table(employee_ids, 'VariableNames', {'ID'});
    for p = 1:n_periods
        detailed_table.(period_labels{p}) = evaluation_data(:, p);
    end
    detailed_table.TotalScore = growth_scores;
    detailed_table.Pattern = growth_patterns;

    % 파일 저장
    basic_file = sprintf('MIDAS_성장평가_결과_v2_%s.xlsx', timestamp);
    detailed_file = sprintf('MIDAS_성장평가_상세_v2_%s.xlsx', timestamp);

    writetable(result_table, basic_file);
    writetable(detailed_table, detailed_file);

    fprintf('✅ 결과 저장 완료:\n');
    fprintf('  기본 결과: %s\n', basic_file);
    fprintf('  상세 결과: %s\n', detailed_file);

catch ME
    fprintf('⚠ 결과 저장 오류: %s\n', ME.message);
end

%% 13. 개별 분석 (선택사항)
fprintf('\n개별 직원 분석 (선택사항)\n');
try
    user_id = input('특정 직원 ID를 입력하세요 (종료: 0): ');

    if user_id > 0
        emp_idx = find(employee_ids == user_id);
        if ~isempty(emp_idx)
            fprintf('\n=== ID %d 상세 분석 ===\n', user_id);
            emp_evals = evaluation_data(emp_idx, :);
            emp_detail = calculation_details{emp_idx};

            fprintf('기간별 평가:\n');
            for p = 1:n_periods
                if iscell(emp_evals) && p <= size(emp_evals, 2)
                    eval_str = char(emp_evals{p});
                else
                    eval_str = 'N/A';
                end

                if ~isempty(emp_detail) && p <= length(emp_detail.score_history)
                    score = emp_detail.score_history(p);
                else
                    score = 0;
                end

                fprintf('  %s: %s → %+d점\n', period_labels{p}, eval_str, score);
            end

            fprintf('\n총합:\n');
            fprintf('  총 점수: %d점\n', growth_scores(emp_idx));
            fprintf('  성장 패턴: %s\n', growth_patterns{emp_idx});

        else
            fprintf('ID %d를 찾을 수 없습니다.\n', user_id);
        end
    end

catch ME
    fprintf('⚠ 개별 분석 오류: %s\n', ME.message);
end

%% 14. 프로그램 완료
total_time = toc(start_time);
fprintf('\n================================================\n');
fprintf('            프로그램 완료\n');
fprintf('================================================\n');
fprintf('완료 시간: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('총 실행 시간: %.2f초\n', total_time);
fprintf('처리 직원 수: %d명\n', n_employees);
fprintf('성공적으로 완료되었습니다.\n');
fprintf('================================================\n');