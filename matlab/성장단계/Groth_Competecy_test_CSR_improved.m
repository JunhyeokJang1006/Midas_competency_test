%% MIDAS 성장단계 평가 점수 산출 프로그램 (개선 버전)
% 작성일: 2024
% 목적: 최근 3년 입사자의 발현 역량 데이터를 기반으로 성장 점수 계산
% 개선사항: 중복 코드 제거, 에러 처리 강화, 로직 최적화, 코드 구조 개선

clear; clc; close all;

%% 1. 초기 설정 및 데이터 불러오기
fprintf('=== MIDAS 성장단계 평가 분석 시작 ===\n\n');

% 파일 경로 설정
filename = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';

% 데이터 로드 및 검증
try
    fprintf('엑셀 파일 구조 확인 중...\n');
    
    % 엑셀 파일의 시트 정보 확인
    [status, sheets] = xlsfinfo(filename);
    if isempty(sheets)
        error('엑셀 파일에서 시트를 찾을 수 없습니다.');
    end
    
    % 시트 목록 출력
    fprintf('발견된 시트:\n');
    for i = 1:length(sheets)
        fprintf('  %d. %s\n', i, sheets{i});
    end
    
    % 발현역량 데이터가 있는 시트 자동 선택
    sheet_name = '';
    for i = 1:length(sheets)
        if contains(sheets{i}, '발현') || contains(sheets{i}, '역량') || ...
           contains(sheets{i}, '평가') || contains(sheets{i}, 'Sheet1')
            sheet_name = sheets{i};
            break;
        end
    end
    
    % 시트를 찾지 못한 경우 사용자 선택
    if isempty(sheet_name)
        fprintf('\n발현역량 데이터 시트를 자동으로 찾지 못했습니다.\n');
        sheet_idx = input('시트 번호를 선택하세요: ');
        if sheet_idx < 1 || sheet_idx > length(sheets)
            error('유효하지 않은 시트 번호입니다.');
        end
        sheet_name = sheets{sheet_idx};
    end
    
    fprintf('\n선택된 시트: %s\n', sheet_name);
    fprintf('데이터 읽는 중...\n');
    
    % 선택된 시트에서 데이터 읽기
    [num_data, txt_data, raw_data] = xlsread(filename, sheet_name);
    [n_rows, n_cols] = size(raw_data);
    fprintf('데이터 크기: %d행 x %d열\n', n_rows, n_cols);
    
    % 헤더 행 찾기
    header_row = find_header_row(raw_data);
    headers = raw_data(header_row, :);
    data_body = raw_data(header_row+1:end, :);
    
    % ID 열과 발현역량 열 찾기
    [id_col, eval_cols] = find_data_columns(headers);
    
    % 데이터 추출 및 정제
    [employee_ids, evaluation_data, n_employees, n_periods] = extract_and_clean_data(data_body, id_col, eval_cols);
    
    % 평가 기간 레이블 생성
    period_labels = generate_period_labels(n_periods, headers, eval_cols);
    
    fprintf('\n데이터 로드 완료:\n');
    fprintf('  - 직원 수: %d명\n', n_employees);
    fprintf('  - 평가 기간: %d개\n', n_periods);
    
catch ME
    error('데이터 로드 오류: %s', ME.message);
end

%% 2. 성장 점수 계산
fprintf('\n성장 점수 계산 시작...\n');

% 결과 저장용 변수 초기화
growth_scores = zeros(n_employees, 1);
growth_patterns = cell(n_employees, 1);
score_details = cell(n_employees, 1);

% 프로그레스 바 설정
progress_interval = max(1, floor(n_employees / 20));

% 각 직원별 성장 점수 계산
for i = 1:n_employees
    % 진행률 표시
    if mod(i, progress_interval) == 0 || i == 1 || i == n_employees
        fprintf('진행: %d/%d (%.1f%%)\n', i, n_employees, i/n_employees*100);
    end
    
    % 개별 직원의 성장 점수 계산
    employee_evals = evaluation_data(i, :);
    [score, pattern, details] = calculate_growth_score(employee_evals);
    
    growth_scores(i) = score;
    growth_patterns{i} = pattern;
    score_details{i} = details;
end

fprintf('계산 완료!\n');

%% 3. 통계 분석
fprintf('\n통계 분석 중...\n');

% 유효한 데이터만 필터링
valid_indices = growth_scores ~= 0 | ~cellfun(@(x) all(strcmp(x, '#N/A')), evaluation_data);
valid_scores = growth_scores(valid_indices);

% 기본 통계 계산
stats = calculate_statistics(valid_scores);

% 패턴별 집계
[unique_patterns, ~, pattern_indices] = unique(growth_patterns(valid_indices));
pattern_counts = accumarray(pattern_indices, 1);

%% 4. 결과 출력
print_analysis_results(n_employees, sum(valid_indices), stats, unique_patterns, pattern_counts);

%% 5. 시각화
fprintf('\n그래프 생성 중...\n');
create_visualizations(valid_scores, stats, unique_patterns, pattern_counts, growth_scores);

%% 6. 결과 저장
fprintf('\n결과 저장 중...\n');
save_results(employee_ids, growth_scores, growth_patterns, evaluation_data, period_labels, n_periods);

%% 7. 개별 직원 상세 분석 (선택사항)
individual_analysis(employee_ids, evaluation_data, growth_scores, growth_patterns, score_details, period_labels);

fprintf('\n프로그램 종료.\n');

%% ========== 헬퍼 함수들 ==========

function header_row = find_header_row(raw_data)
    % 헤더 행을 찾는 함수
    [n_rows, ~] = size(raw_data);
    header_row = 0;
    
    for i = 1:min(5, n_rows)
        row_data = raw_data(i, :);
        if any(cellfun(@(x) ischar(x) && contains(x, '발현'), row_data)) || ...
           any(cellfun(@(x) ischar(x) && contains(x, 'H1'), row_data)) || ...
           any(cellfun(@(x) ischar(x) && contains(x, 'H2'), row_data))
            header_row = i;
            break;
        end
    end
    
    if header_row == 0
        header_row = 1;
        fprintf('헤더를 자동으로 찾을 수 없어 1행을 헤더로 가정합니다.\n');
    end
end

function [id_col, eval_cols] = find_data_columns(headers)
    % ID 열과 발현역량 열을 찾는 함수
    id_col = 0;
    eval_cols = [];
    
    for i = 1:length(headers)
        if ischar(headers{i})
            if contains(headers{i}, 'ID') || contains(headers{i}, '사번')
                id_col = i;
            elseif contains(headers{i}, '발현') || contains(headers{i}, 'H1') || ...
                   contains(headers{i}, 'H2')
                eval_cols = [eval_cols, i];
            end
        end
    end
    
    if id_col == 0
        id_col = 1;
        fprintf('ID 열을 찾을 수 없어 1열을 ID로 가정합니다.\n');
    end
    
    if isempty(eval_cols)
        eval_cols = (id_col+1):min(id_col+5, length(headers));
        fprintf('발현역량 열을 자동으로 찾을 수 없어 %d~%d열을 사용합니다.\n', ...
                eval_cols(1), eval_cols(end));
    end
end

function [employee_ids, evaluation_data, n_employees, n_periods] = extract_and_clean_data(data_body, id_col, eval_cols)
    % 데이터 추출 및 정제 함수
    employee_ids = data_body(:, id_col);
    evaluation_data = data_body(:, eval_cols);
    
    % ID가 숫자인 경우 처리
    if all(cellfun(@isnumeric, employee_ids))
        employee_ids = cell2mat(employee_ids);
    end
    
    % 유효한 데이터 필터링 (ID가 있는 행만)
    valid_rows = ~cellfun(@(x) isempty(x) || (isnumeric(x) && isnan(x)), ...
                          data_body(:, id_col));
    
    employee_ids = employee_ids(valid_rows);
    evaluation_data = evaluation_data(valid_rows, :);
    
    n_employees = length(employee_ids);
    n_periods = size(evaluation_data, 2);
end

function period_labels = generate_period_labels(n_periods, headers, eval_cols)
    % 평가 기간 레이블 생성 함수
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
end

function [score, pattern, details] = calculate_growth_score(evaluations)
    % 성장 점수 계산 함수
    score = 0;
    base_level = '열린(Lv2)';
    consecutive_higher = 0;
    score_history = zeros(1, length(evaluations));
    base_history = cell(1, length(evaluations));
    
    for period = 1:length(evaluations)
        current_eval = evaluations{period};
        period_score = 0;
        
        % 데이터 정제
        current_eval = clean_evaluation_data(current_eval);
        
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
    pattern = classify_growth_pattern(score);
    
    details.score_history = score_history;
    details.base_history = base_history;
    details.final_base_level = base_level;
end

function cleaned_eval = clean_evaluation_data(current_eval)
    % 평가 데이터 정제 함수
    if isnumeric(current_eval)
        if isnan(current_eval)
            cleaned_eval = '#N/A';
        else
            cleaned_eval = num2str(current_eval);
        end
    elseif ~ischar(current_eval)
        cleaned_eval = char(current_eval);
    else
        cleaned_eval = current_eval;
    end
    
    cleaned_eval = strtrim(cleaned_eval);
end

function pattern = classify_growth_pattern(score)
    % 성장 패턴 분류 함수
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
end

function stats = calculate_statistics(valid_scores)
    % 통계 계산 함수
    stats = struct();
    stats.mean = mean(valid_scores);
    stats.median = median(valid_scores);
    stats.std = std(valid_scores);
    stats.min = min(valid_scores);
    stats.max = max(valid_scores);
    stats.quartiles = quantile(valid_scores, [0.25, 0.5, 0.75]);
end

function print_analysis_results(n_employees, valid_count, stats, unique_patterns, pattern_counts)
    % 분석 결과 출력 함수
    fprintf('\n========== MIDAS 성장단계 평가 분석 결과 ==========\n');
    fprintf('분석 대상: %d명 (유효 데이터: %d명)\n', n_employees, valid_count);
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
            pattern_counts(i)/valid_count*100);
    end
end

function create_visualizations(valid_scores, stats, unique_patterns, pattern_counts, growth_scores)
    % 시각화 생성 함수
    figure('Name', 'MIDAS 성장단계 평가 분석', 'Position', [100, 100, 1400, 800]);

    % 1) 점수 분포 히스토그램
    subplot(2, 3, 1);
    histogram(valid_scores, 15, 'FaceColor', [0.2 0.4 0.6]);
    xlabel('성장 점수');
    ylabel('인원수');
    title('성장 점수 분포');
    grid on;
    hold on;
    xline(stats.mean, 'r-', 'LineWidth', 2);
    xline(stats.median, 'g--', 'LineWidth', 1.5);
    legend('분포', '평균', '중앙값', 'Location', 'best');

    % 2) 패턴별 분포 (파이 차트)
    subplot(2, 3, 2);
    pie(pattern_counts, unique_patterns);
    title('성장 패턴 분포');
    colormap(autumn);

    % 3) Box Plot
    subplot(2, 3, 3);
    boxplot(valid_scores, 'Labels', {'전체'});
    ylabel('성장 점수');
    title('성장 점수 Box Plot');
    grid on;

    % 4) 누적 분포 함수
    subplot(2, 3, 4);
    [f, x] = ecdf(valid_scores);
    plot(x, f, 'LineWidth', 2);
    xlabel('성장 점수');
    ylabel('누적 확률');
    title('누적 분포 함수 (CDF)');
    grid on;

    % 5) 점수 구간별 분포
    subplot(2, 3, 5);
    score_ranges = [-inf, 0, 10, 20, 30, inf];
    score_labels = {'퇴보(<0)', '정체(0)', '저성장(1-10)', ...
                '중성장(11-20)', '고성장(>20)'};
    score_groups = discretize(valid_scores, score_ranges);
    bar(accumarray(score_groups, 1));
    set(gca, 'XTickLabel', score_labels);
    xtickangle(45);
    ylabel('인원수');
    title('점수 구간별 분포');
    grid on;

    % 6) 상위 20명 성장 점수
    subplot(2, 3, 6);
    [sorted_scores, ~] = sort(growth_scores, 'descend');
    top_n = min(20, sum(sorted_scores > 0));
    bar(1:top_n, sorted_scores(1:top_n));
    xlabel('순위');
    ylabel('성장 점수');
    title('상위 20명 성장 점수');
    grid on;
end

function save_results(employee_ids, growth_scores, growth_patterns, evaluation_data, period_labels, n_periods)
    % 결과 저장 함수
    % 기본 결과 테이블 생성
    result_table = table(employee_ids, growth_scores, growth_patterns, ...
        'VariableNames', {'ID', 'GrowthScore', 'Pattern'});
    
    % 엑셀 파일로 저장
    output_filename = 'MIDAS_성장평가_결과.xlsx';
    writetable(result_table, output_filename);
    
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
    writetable(detailed_results, detailed_output_filename, 'Sheet', '분석결과');
    
    % MAT 파일로 저장
    save('MIDAS_growth_analysis.mat', 'result_table', 'detailed_results', ...
         'evaluation_data', 'period_labels');
    
    fprintf('저장 완료:\n');
    fprintf('  - 기본 결과: %s\n', output_filename);
    fprintf('  - 상세 결과: %s\n', detailed_output_filename);
    fprintf('  - 데이터 파일: MIDAS_growth_analysis.mat\n');
end

function individual_analysis(employee_ids, evaluation_data, growth_scores, growth_patterns, score_details, period_labels)
    % 개별 직원 상세 분석 함수
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
end
