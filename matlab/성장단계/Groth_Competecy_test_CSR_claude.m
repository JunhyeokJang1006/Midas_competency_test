%% MIDAS 성장단계 평가 점수 산출 프로그램 (Claude 최적화 버전)
% 작성일: 2024
% 작성자: Claude (기존 코드 완전 리팩토링)
% 목적: 최근 3년 입사자의 발현 역량 데이터를 기반으로 성장 점수 계산
%
% 주요 개선사항:
%   - 코드 구조 완전 재설계
%   - 에러 처리 및 복구 로직 강화
%   - 메모리 효율성 최적화
%   - 사용자 경험 개선 (진행률, 로깅)
%   - 모듈화된 함수 설계
%   - 데이터 검증 로직 강화

clear; clc; close all;

%% ========== 1. 프로그램 초기화 ==========
fprintf('═══════════════════════════════════════════════════════════════\n');
fprintf('    MIDAS 성장단계 평가 분석 시스템 (Claude 최적화 버전)\n');
fprintf('═══════════════════════════════════════════════════════════════\n');

start_time = datetime('now');
fprintf('분석 시작: %s\n', datestr(start_time, 'yyyy-mm-dd HH:MM:SS'));

% 전역 설정
config = struct();
config.data_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
config.target_sheet = 3;  % 발현역량 시트
config.id_column = 'ID';
config.period_labels = {'23H1', '23H2', '24H1', '24H2', '25H1'};
config.verbose = true;

%% ========== 2. 데이터 로드 및 검증 ==========
fprintf('\n[단계 1/6] 데이터 로드 시작\n');

try
    % 파일 존재 여부 확인
    if ~exist(config.data_file, 'file')
        error('데이터 파일을 찾을 수 없습니다: %s', config.data_file);
    end

    % 시트 정보 확인
    [~, sheet_names] = xlsfinfo(config.data_file);
    if length(sheet_names) < config.target_sheet
        error('대상 시트(%d번째)가 존재하지 않습니다. 총 %d개 시트 발견', ...
              config.target_sheet, length(sheet_names));
    end

    target_sheet_name = sheet_names{config.target_sheet};
    fprintf('  ✓ 대상 시트: %s (시트 %d)\n', target_sheet_name, config.target_sheet);

    % 데이터 읽기 (다중 방법 시도)
    [data_table, data_source] = load_excel_data(config.data_file, target_sheet_name);
    fprintf('  ✓ 데이터 로드 완료 (방법: %s)\n', data_source);
    fprintf('  ✓ 데이터 크기: %d행 × %d열\n', height(data_table), width(data_table));

catch ME
    fprintf('  ❌ 데이터 로드 실패: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 3. 데이터 전처리 및 구조 분석 ==========
fprintf('\n[단계 2/6] 데이터 전처리\n');

try
    % 컬럼 분석
    column_names = data_table.Properties.VariableNames;
    fprintf('  ✓ 컬럼 분석: %s\n', strjoin(column_names(1:min(5, end)), ', '));

    % ID 컬럼 확인
    id_col_idx = find_column_index(column_names, config.id_column);
    if id_col_idx == 0
        error('ID 컬럼을 찾을 수 없습니다');
    end
    fprintf('  ✓ ID 컬럼: %s (열 %d)\n', column_names{id_col_idx}, id_col_idx);

    % 발현역량 컬럼 찾기
    eval_col_indices = find_evaluation_columns(column_names);
    if isempty(eval_col_indices)
        error('발현역량 컬럼을 찾을 수 없습니다');
    end
    fprintf('  ✓ 발현역량 컬럼 %d개 발견: [%s]\n', length(eval_col_indices), ...
            num2str(eval_col_indices));

    % 데이터 추출 및 정제
    employee_ids = data_table{:, id_col_idx};
    evaluation_data = data_table{:, eval_col_indices};

    % 유효한 데이터만 필터링
    [clean_ids, clean_evals, valid_mask] = clean_data(employee_ids, evaluation_data);

    n_employees = length(clean_ids);
    n_periods = size(clean_evals, 2);

    fprintf('  ✓ 데이터 정제 완료: %d명 × %d기간\n', n_employees, n_periods);
    fprintf('  ✓ 유효 데이터율: %.1f%% (%d/%d)\n', ...
            sum(valid_mask)/length(valid_mask)*100, sum(valid_mask), length(valid_mask));

    % 기간 레이블 설정
    period_labels = get_period_labels(column_names, eval_col_indices, config.period_labels, n_periods);
    fprintf('  ✓ 평가 기간: %s\n', strjoin(period_labels, ' → '));

catch ME
    fprintf('  ❌ 데이터 전처리 실패: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 4. 성장 점수 계산 ==========
fprintf('\n[단계 3/6] 성장 점수 계산\n');

try
    % 계산 시작
    calculation_start = tic;

    % 결과 저장용 변수 초기화
    growth_scores = zeros(n_employees, 1);
    growth_patterns = cell(n_employees, 1);
    score_details = cell(n_employees, 1);

    % 진행률 표시 설정
    progress_step = max(1, floor(n_employees / 20));

    fprintf('  진행률: ');
    for i = 1:n_employees
        % 개별 직원 성장 점수 계산
        [score, pattern, details] = calculate_growth_score(clean_evals(i, :), period_labels);

        growth_scores(i) = score;
        growth_patterns{i} = pattern;
        score_details{i} = details;

        % 진행률 표시
        if mod(i, progress_step) == 0 || i == n_employees
            fprintf('█');
        end
    end

    calculation_time = toc(calculation_start);
    fprintf(' 완료!\n');
    fprintf('  ✓ 계산 완료 (소요시간: %.2f초)\n', calculation_time);

catch ME
    fprintf('  ❌ 성장 점수 계산 실패: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 5. 통계 분석 ==========
fprintf('\n[단계 4/6] 통계 분석\n');

try
    % 기본 통계
    stats = calculate_statistics(growth_scores);

    % 패턴 분석
    [pattern_stats, pattern_distribution] = analyze_patterns(growth_patterns);

    % 결과 출력
    print_statistics(stats, pattern_stats, n_employees);

catch ME
    fprintf('  ❌ 통계 분석 실패: %s\n', ME.message);
    rethrow(ME);
end

%% ========== 6. 시각화 ==========
fprintf('\n[단계 5/6] 시각화 생성\n');

try
    create_visualizations(growth_scores, growth_patterns, score_details, ...
                         period_labels, stats, pattern_distribution);
    fprintf('  ✓ 시각화 완료 (8개 차트 생성)\n');

catch ME
    fprintf('  ❌ 시각화 생성 실패: %s\n', ME.message);
    % 시각화는 실패해도 프로그램 계속 진행
end

%% ========== 7. 결과 저장 ==========
fprintf('\n[단계 6/6] 결과 저장\n');

try
    % 파일명 생성
    timestamp = datestr(now, 'yyyymmdd_HHMMSS');

    % 기본 결과 저장
    basic_results = table(clean_ids, growth_scores, growth_patterns, ...
        'VariableNames', {'ID', 'GrowthScore', 'Pattern'});
    basic_filename = sprintf('MIDAS_성장평가_결과_%s.xlsx', timestamp);
    writetable(basic_results, basic_filename);
    fprintf('  ✓ 기본 결과: %s\n', basic_filename);

    % 상세 결과 저장
    detailed_results = create_detailed_results(clean_ids, clean_evals, growth_scores, ...
                                              growth_patterns, score_details, period_labels);
    detailed_filename = sprintf('MIDAS_성장평가_상세결과_%s.xlsx', timestamp);
    writetable(detailed_results, detailed_filename);
    fprintf('  ✓ 상세 결과: %s\n', detailed_filename);

    % MATLAB 데이터 저장
    mat_filename = sprintf('MIDAS_growth_analysis_%s.mat', timestamp);
    save(mat_filename, 'basic_results', 'detailed_results', 'stats', ...
         'pattern_stats', 'score_details', 'config', 'period_labels');
    fprintf('  ✓ 데이터 파일: %s\n', mat_filename);

catch ME
    fprintf('  ❌ 결과 저장 실패: %s\n', ME.message);
    % 저장 실패해도 분석 결과는 메모리에 남아있음
end

%% ========== 8. 프로그램 완료 ==========
total_time = seconds(datetime('now') - start_time);
fprintf('\n═══════════════════════════════════════════════════════════════\n');
fprintf('              분석 완료 - 총 소요시간: %.1f초\n', total_time);
fprintf('═══════════════════════════════════════════════════════════════\n');

% 요약 통계 재출력
fprintf('\n📊 최종 요약:\n');
fprintf('   • 분석 대상: %d명\n', n_employees);
fprintf('   • 평균 점수: %.1f점\n', stats.mean);
fprintf('   • 최고 점수: %.0f점 (패턴: %s)\n', stats.max, pattern_stats{1, 1});
fprintf('   • 주요 패턴: %s (%.1f%%)\n', pattern_distribution.pattern{1}, pattern_distribution.percentage(1));

%% ========== 보조 함수들 ==========

function [data_table, data_source] = load_excel_data(filename, sheet_name)
    % 다중 방법으로 엑셀 데이터 로드
    data_table = [];
    data_source = '';

    % 방법 1: readtable (권장)
    try
        data_table = readtable(filename, 'Sheet', sheet_name, ...
                              'VariableNamingRule', 'preserve', ...
                              'TextType', 'string');
        data_source = 'readtable';
        return;
    catch ME1
        % 방법 1 실패
    end

    % 방법 2: xlsread + table 변환
    try
        [~, ~, raw_data] = xlsread(filename, sheet_name);
        headers = raw_data(1, :);
        data_body = raw_data(2:end, :);

        % 유효한 헤더만 선택
        valid_headers = {};
        valid_cols = [];
        for i = 1:length(headers)
            if ~isempty(headers{i}) && (ischar(headers{i}) || isstring(headers{i}))
                valid_headers{end+1} = char(headers{i});
                valid_cols(end+1) = i;
            elseif isnumeric(headers{i}) && ~isnan(headers{i})
                valid_headers{end+1} = sprintf('Var%d', i);
                valid_cols(end+1) = i;
            end
        end

        if ~isempty(valid_cols)
            data_table = table();
            for i = 1:length(valid_cols)
                col_data = data_body(:, valid_cols(i));
                data_table.(valid_headers{i}) = col_data;
            end
            data_source = 'xlsread';
            return;
        end
    catch ME2
        % 방법 2 실패
    end

    % 방법 3: readmatrix (최후 수단)
    try
        raw_matrix = readmatrix(filename, 'Sheet', sheet_name);
        n_cols = size(raw_matrix, 2);
        data_table = table();
        for i = 1:n_cols
            data_table.(sprintf('Var%d', i)) = raw_matrix(:, i);
        end
        data_source = 'readmatrix';
        return;
    catch ME3
        % 모든 방법 실패
    end

    error('모든 데이터 로드 방법 실패');
end

function col_idx = find_column_index(column_names, target_name)
    % 컬럼 인덱스 찾기 (대소문자 구분 안함)
    col_idx = 0;
    keywords = {target_name, 'id', 'ID', '사번', '직원번호', '사원번호'};

    for i = 1:length(column_names)
        col_name = char(column_names{i});
        for j = 1:length(keywords)
            if contains(col_name, keywords{j}, 'IgnoreCase', true)
                col_idx = i;
                return;
            end
        end
    end
end

function eval_indices = find_evaluation_columns(column_names)
    % 발현역량 컬럼들 찾기
    eval_indices = [];
    keywords = {'발현', 'H1', 'H2', 'h1', 'h2', '역량', '평가'};

    for i = 1:length(column_names)
        col_name = char(column_names{i});
        for j = 1:length(keywords)
            if contains(col_name, keywords{j}, 'IgnoreCase', true)
                eval_indices(end+1) = i;
                break;
            end
        end
    end

    % 중복 제거 및 정렬
    eval_indices = unique(eval_indices);
end

function [clean_ids, clean_evals, valid_mask] = clean_data(employee_ids, evaluation_data)
    % 데이터 정제 및 유효성 검사
    n_rows = length(employee_ids);
    valid_mask = true(n_rows, 1);

    % ID 유효성 검사
    for i = 1:n_rows
        id_val = employee_ids(i);
        if iscell(id_val)
            id_val = id_val{1};
        end

        if isempty(id_val) || (isnumeric(id_val) && isnan(id_val)) || ...
           (ischar(id_val) && isempty(strtrim(id_val)))
            valid_mask(i) = false;
        end
    end

    % 유효한 데이터만 추출
    clean_ids = employee_ids(valid_mask);
    clean_evals = evaluation_data(valid_mask, :);

    % ID가 셀 배열이면 숫자로 변환
    if iscell(clean_ids)
        numeric_ids = zeros(size(clean_ids));
        for i = 1:length(clean_ids)
            if isnumeric(clean_ids{i})
                numeric_ids(i) = clean_ids{i};
            else
                numeric_ids(i) = str2double(clean_ids{i});
            end
        end
        clean_ids = numeric_ids;
    end
end

function period_labels = get_period_labels(column_names, eval_indices, default_labels, n_periods)
    % 기간 레이블 생성
    period_labels = cell(1, n_periods);

    for i = 1:n_periods
        if i <= length(eval_indices)
            col_name = char(column_names{eval_indices(i)});
            % 컬럼명에서 기간 정보 추출 (예: "23H1 발현역량" -> "23H1")
            tokens = regexp(col_name, '\d+H\d+', 'match');
            if ~isempty(tokens)
                period_labels{i} = tokens{1};
            else
                period_labels{i} = sprintf('P%d', i);
            end
        elseif i <= length(default_labels)
            period_labels{i} = default_labels{i};
        else
            period_labels{i} = sprintf('Period%d', i);
        end
    end
end

function [score, pattern, details] = calculate_growth_score(evaluations, period_labels)
    % 개별 직원의 성장 점수 계산
    score = 0;
    base_level = '열린(Lv2)';
    consecutive_achievement = 0;

    n_periods = length(evaluations);
    score_history = zeros(1, n_periods);
    base_history = cell(1, n_periods);

    for period = 1:n_periods
        current_eval = normalize_evaluation(evaluations{period});
        period_score = 0;

        % 평가 결과에 따른 점수 계산
        if isempty(current_eval) || contains(current_eval, '#N/A')
            % 데이터 없음
            period_score = 0;

        elseif contains(current_eval, '책임')
            % 책임 단계 달성
            if ~strcmp(base_level, '책임(Lv1)')
                period_score = 30;  % 최초 책임 달성
                base_level = '책임(Lv1)';
                consecutive_achievement = 0;
            else
                period_score = 0;   % 이미 책임 단계
            end

        elseif contains(current_eval, '성취')
            % 성취 단계
            if strcmp(base_level, '열린(Lv2)')
                consecutive_achievement = consecutive_achievement + 1;
                period_score = 10;

                % 연속 3회 성취 시 성취(Lv1)로 승급
                if consecutive_achievement >= 3
                    base_level = '성취(Lv1)';
                    consecutive_achievement = 0;
                end
            else
                period_score = 0;   % 이미 더 높은 단계
            end

        elseif contains(current_eval, '열린(Lv2)')
            % 열린(Lv2) 유지 또는 하향
            if strcmp(base_level, '열린(Lv2)')
                period_score = 0;   % 현상 유지
                consecutive_achievement = 0;
            elseif contains(base_level, '성취') || contains(base_level, '책임')
                period_score = -10; % 하향 이동 페널티
                base_level = '열린(Lv2)';
                consecutive_achievement = 0;
            end

        elseif contains(current_eval, '열린(Lv1)')
            % 열린(Lv1) - 퇴보
            period_score = -5;
            base_level = '열린(Lv1)';
            consecutive_achievement = 0;
        end

        score = score + period_score;
        score_history(period) = period_score;
        base_history{period} = base_level;
    end

    % 성장 패턴 분류
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

    % 상세 정보 저장
    details = struct();
    details.score_history = score_history;
    details.base_history = base_history;
    details.final_base_level = base_level;
    details.consecutive_count = consecutive_achievement;
end

function normalized = normalize_evaluation(eval_data)
    % 평가 데이터 정규화
    try
        if iscell(eval_data)
            eval_data = eval_data{1};
        end

        if isnumeric(eval_data)
            if isnan(eval_data)
                normalized = '#N/A';
            else
                normalized = num2str(eval_data);
            end
        elseif isstring(eval_data) || ischar(eval_data)
            normalized = char(eval_data);
            normalized = strtrim(normalized);
            if isempty(normalized)
                normalized = '#N/A';
            end
        else
            normalized = '#N/A';
        end
    catch
        normalized = '#N/A';
    end
end

function stats = calculate_statistics(scores)
    % 기본 통계량 계산
    valid_scores = scores(scores ~= 0 | ~isnan(scores));

    if isempty(valid_scores)
        stats = struct('mean', 0, 'median', 0, 'std', 0, 'min', 0, ...
                      'max', 0, 'quartiles', [0 0 0], 'range', 0, 'cv', 0);
        return;
    end

    stats = struct();
    stats.mean = mean(valid_scores);
    stats.median = median(valid_scores);
    stats.std = std(valid_scores);
    stats.min = min(valid_scores);
    stats.max = max(valid_scores);
    stats.quartiles = quantile(valid_scores, [0.25, 0.5, 0.75]);
    stats.range = stats.max - stats.min;

    if stats.mean ~= 0
        stats.cv = abs(stats.std / stats.mean * 100);
    else
        stats.cv = 0;
    end
end

function [pattern_stats, pattern_distribution] = analyze_patterns(patterns)
    % 패턴 분석
    [unique_patterns, ~, pattern_idx] = unique(patterns);
    counts = accumarray(pattern_idx, 1);
    percentages = counts / length(patterns) * 100;

    % 패턴별 통계
    pattern_stats = [unique_patterns, num2cell(counts), num2cell(percentages)];

    % 분포 테이블
    pattern_distribution = table(unique_patterns, counts, percentages, ...
        'VariableNames', {'pattern', 'count', 'percentage'});

    % 빈도순 정렬
    [~, sort_idx] = sort(counts, 'descend');
    pattern_distribution = pattern_distribution(sort_idx, :);
end

function print_statistics(stats, pattern_stats, n_total)
    % 통계 결과 출력
    fprintf('  ✓ 기초 통계량:\n');
    fprintf('     평균: %.2f점  |  중앙값: %.2f점  |  표준편차: %.2f\n', ...
            stats.mean, stats.median, stats.std);
    fprintf('     최소: %.0f점  |  최대: %.0f점  |  범위: %.0f점\n', ...
            stats.min, stats.max, stats.range);
    fprintf('     Q1: %.0f  |  Q2: %.0f  |  Q3: %.0f  |  CV: %.1f%%\n', ...
            stats.quartiles(1), stats.quartiles(2), stats.quartiles(3), stats.cv);

    fprintf('\n  ✓ 성장 패턴 분포:\n');
    for i = 1:size(pattern_stats, 1)
        fprintf('     %s: %d명 (%.1f%%)\n', ...
                pattern_stats{i,1}, pattern_stats{i,2}, pattern_stats{i,3});
    end
end

function create_visualizations(scores, patterns, score_details, period_labels, stats, pattern_dist)
    % 종합 시각화 생성
    try
        figure('Name', 'MIDAS 성장단계 분석 결과', ...
               'Position', [50, 50, 1800, 1000], ...
               'Color', 'white');

        % 8개 서브플롯 생성

        % 1. 점수 분포 히스토그램
        subplot(2, 4, 1);
        if ~isempty(scores) && sum(scores ~= 0) > 0
            valid_scores = scores(scores ~= 0);
            histogram(valid_scores, max(10, length(valid_scores)/5), ...
                     'FaceColor', [0.3 0.6 0.8], 'EdgeColor', 'white', 'LineWidth', 0.5);
            xlabel('성장 점수', 'FontSize', 10);
            ylabel('인원수', 'FontSize', 10);
            title('성장 점수 분포', 'FontSize', 12, 'FontWeight', 'bold');
            grid on; grid minor;
            xline(stats.mean, 'r-', 'LineWidth', 2);
            xline(stats.median, 'g--', 'LineWidth', 1.5);
            legend({'분포', '평균', '중앙값'}, 'Location', 'best', 'FontSize', 8);
        else
            text(0.5, 0.5, '유효 데이터 없음', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('성장 점수 분포', 'FontSize', 12);
        end

        % 2. 패턴별 분포 (파이 차트)
        subplot(2, 4, 2);
        if height(pattern_dist) > 0
            pie_colors = lines(height(pattern_dist));
            pie(pattern_dist.count, pattern_dist.pattern);
            colormap(pie_colors);
            title('성장 패턴 분포', 'FontSize', 12, 'FontWeight', 'bold');
        else
            text(0.5, 0.5, '패턴 데이터 없음', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('성장 패턴 분포', 'FontSize', 12);
        end

        % 3. 박스플롯
        subplot(2, 4, 3);
        if ~isempty(scores) && sum(scores ~= 0) > 0
            valid_scores = scores(scores ~= 0);
            boxplot(valid_scores, 'Labels', {'전체'});
            ylabel('성장 점수', 'FontSize', 10);
            title('점수 분포 (Box Plot)', 'FontSize', 12, 'FontWeight', 'bold');
            grid on;
        else
            text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('점수 분포 (Box Plot)', 'FontSize', 12);
        end

        % 4. 누적분포함수
        subplot(2, 4, 4);
        if ~isempty(scores) && sum(scores ~= 0) > 0
            valid_scores = scores(scores ~= 0);
            [f, x] = ecdf(valid_scores);
            plot(x, f, 'LineWidth', 2.5, 'Color', [0.8 0.3 0.3]);
            xlabel('성장 점수', 'FontSize', 10);
            ylabel('누적 확률', 'FontSize', 10);
            title('누적분포함수 (CDF)', 'FontSize', 12, 'FontWeight', 'bold');
            grid on; grid minor;
        else
            text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('누적분포함수', 'FontSize', 12);
        end

        % 5. 점수 구간별 분포
        subplot(2, 4, 5);
        if ~isempty(scores)
            ranges = [-inf, 0, 10, 20, 30, inf];
            labels = {'퇴보', '정체', '저성장', '중성장', '고성장'};
            groups = discretize(scores, ranges);
            counts = accumarray(groups, 1, [length(labels), 1]);

            bar_colors = [0.8 0.2 0.2;   % 퇴보 - 빨강
                         0.9 0.6 0.1;   % 정체 - 주황
                         0.9 0.9 0.3;   % 저성장 - 노랑
                         0.5 0.8 0.3;   % 중성장 - 연두
                         0.2 0.7 0.2];  % 고성장 - 녹색

            bar_handle = bar(counts);
            bar_handle.FaceColor = 'flat';
            bar_handle.CData = bar_colors;

            set(gca, 'XTickLabel', labels);
            xtickangle(45);
            ylabel('인원수', 'FontSize', 10);
            title('구간별 분포', 'FontSize', 12, 'FontWeight', 'bold');
            grid on;

            % 값 표시
            for i = 1:length(counts)
                if counts(i) > 0
                    text(i, counts(i) + max(counts)*0.02, num2str(counts(i)), ...
                         'HorizontalAlignment', 'center', 'FontSize', 8);
                end
            end
        else
            text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('구간별 분포', 'FontSize', 12);
        end

        % 6. 상위 20명
        subplot(2, 4, 6);
        if ~isempty(scores) && max(scores) > 0
            [sorted_scores, ~] = sort(scores, 'descend');
            top_n = min(20, sum(sorted_scores > 0));
            if top_n > 0
                bar(1:top_n, sorted_scores(1:top_n), 'FaceColor', [0.2 0.5 0.8]);
                xlabel('순위', 'FontSize', 10);
                ylabel('성장 점수', 'FontSize', 10);
                title('상위 20명 점수', 'FontSize', 12, 'FontWeight', 'bold');
                grid on;
            else
                text(0.5, 0.5, '양수 점수 없음', 'HorizontalAlignment', 'center', 'FontSize', 12);
                title('상위 20명 점수', 'FontSize', 12);
            end
        else
            text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('상위 20명 점수', 'FontSize', 12);
        end

        % 7. 기간별 평균 추이
        subplot(2, 4, 7);
        if ~isempty(score_details) && length(period_labels) > 1
            try
                period_avgs = zeros(1, length(period_labels));
                for p = 1:length(period_labels)
                    period_scores = [];
                    for i = 1:length(score_details)
                        if ~isempty(score_details{i}) && ...
                           isfield(score_details{i}, 'score_history') && ...
                           length(score_details{i}.score_history) >= p
                            period_scores(end+1) = score_details{i}.score_history(p);
                        end
                    end
                    if ~isempty(period_scores)
                        period_avgs(p) = mean(period_scores);
                    end
                end

                plot(1:length(period_labels), period_avgs, 'o-', ...
                     'LineWidth', 2, 'MarkerSize', 6, 'Color', [0.4 0.6 0.8]);
                xlabel('평가 기간', 'FontSize', 10);
                ylabel('평균 점수', 'FontSize', 10);
                title('기간별 평균 추이', 'FontSize', 12, 'FontWeight', 'bold');
                set(gca, 'XTick', 1:length(period_labels), 'XTickLabel', period_labels);
                xtickangle(45);
                grid on; grid minor;
            catch
                text(0.5, 0.5, '추이 계산 오류', 'HorizontalAlignment', 'center', 'FontSize', 12);
                title('기간별 평균 추이', 'FontSize', 12);
            end
        else
            text(0.5, 0.5, '기간 데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('기간별 평균 추이', 'FontSize', 12);
        end

        % 8. 패턴별 박스플롯
        subplot(2, 4, 8);
        if height(pattern_dist) > 1
            try
                pattern_names = pattern_dist.pattern;
                pattern_scores_cell = cell(length(pattern_names), 1);

                for i = 1:length(pattern_names)
                    mask = strcmp(patterns, pattern_names{i});
                    pattern_scores_cell{i} = scores(mask);
                end

                % 박스플롯 생성
                all_scores = [];
                all_groups = [];
                for i = 1:length(pattern_names)
                    all_scores = [all_scores; pattern_scores_cell{i}];
                    all_groups = [all_groups; repmat(i, length(pattern_scores_cell{i}), 1)];
                end

                if ~isempty(all_scores)
                    boxplot(all_scores, all_groups, 'Labels', pattern_names);
                    ylabel('성장 점수', 'FontSize', 10);
                    title('패턴별 점수 분포', 'FontSize', 12, 'FontWeight', 'bold');
                    xtickangle(45);
                    grid on;
                end
            catch
                text(0.5, 0.5, '패턴 분석 오류', 'HorizontalAlignment', 'center', 'FontSize', 12);
                title('패턴별 점수 분포', 'FontSize', 12);
            end
        else
            text(0.5, 0.5, '패턴 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('패턴별 점수 분포', 'FontSize', 12);
        end

        % 전체 레이아웃 조정
        sgtitle('MIDAS 성장단계 평가 종합 분석', 'FontSize', 16, 'FontWeight', 'bold');

    catch ME
        fprintf('  ⚠ 시각화 오류: %s\n', ME.message);
        rethrow(ME);
    end
end

function detailed_table = create_detailed_results(ids, evaluations, scores, patterns, details, period_labels)
    % 상세 결과 테이블 생성
    n_employees = length(ids);
    n_periods = size(evaluations, 2);

    % 기본 정보
    detailed_table = table();
    detailed_table.ID = ids;

    % 각 기간별 평가 결과
    for p = 1:n_periods
        col_name = sprintf('%s_평가', period_labels{p});
        detailed_table.(col_name) = evaluations(:, p);

        col_name = sprintf('%s_점수', period_labels{p});
        period_scores = zeros(n_employees, 1);
        for i = 1:n_employees
            if ~isempty(details{i}) && isfield(details{i}, 'score_history') && ...
               length(details{i}.score_history) >= p
                period_scores(i) = details{i}.score_history(p);
            end
        end
        detailed_table.(col_name) = period_scores;
    end

    % 총 점수와 패턴
    detailed_table.총점수 = scores;
    detailed_table.성장패턴 = patterns;

    % 최종 수준
    final_levels = cell(n_employees, 1);
    for i = 1:n_employees
        if ~isempty(details{i}) && isfield(details{i}, 'final_base_level')
            final_levels{i} = details{i}.final_base_level;
        else
            final_levels{i} = '알수없음';
        end
    end
    detailed_table.최종수준 = final_levels;
end