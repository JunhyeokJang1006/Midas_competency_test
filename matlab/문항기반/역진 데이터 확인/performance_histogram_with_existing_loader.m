%% 기존 데이터 로더를 사용한 성과 문항 원점수 히스토그램 생성기
%
% 목적: 기존 corr_item_vs_comp_score.m의 데이터 로드 로직을 그대로 사용하여
%       각 시점별 성과 문항의 원점수 분포를 히스토그램으로 시각화
%
% 작성일: 2025년

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('성과 문항 원점수 히스토그램 생성기\n');
fprintf('(기존 데이터 로더 사용)\n');
fprintf('========================================\n\n');

%% 1. 기존 분석 결과 로드
fprintf('[1단계] 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

% 기존 역량진단 데이터 로드 (factor_analysis_by_period.m 결과)
consolidatedScores = [];
allData = struct();
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};

% MAT 파일에서 로드
matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    matFileName = matFiles(idx).name;

    try
        loadedData = load(matFileName);
        if isfield(loadedData, 'allData')
            allData = loadedData.allData;
            fprintf('✓ 역량진단 데이터 로드 완료\n');

            % allData 구조 확인
            fields = fieldnames(allData);
            fprintf('  - 로드된 Period 수: %d개\n', length(fields));
            for i = 1:length(fields)
                if isfield(allData.(fields{i}), 'selfData')
                    fprintf('  - %s: %d명\n', fields{i}, height(allData.(fields{i}).selfData));
                end
            end
        else
            fprintf('✗ allData를 찾을 수 없습니다\n');
        end

        if isfield(loadedData, 'consolidatedScores')
            consolidatedScores = loadedData.consolidatedScores;
        end

        if isfield(loadedData, 'periods')
            periods = loadedData.periods;
        end

    catch ME
        fprintf('✗ MAT 파일 로드 실패: %s\n', ME.message);
    end
end

% 역량검사 데이터 로드
fprintf('\n역량검사 데이터 로드 중...\n');
competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

% -------------------------------------------------------------
% 1. 역량검사 데이터 로드
% -------------------------------------------------------------
fprintf('\n역량검사 데이터 로드 중...\n');
competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    competencyTestData = readtable( competencyTestPath, ...
        'Sheet', '역량검사_종합점수', ...
        'VariableNamingRule', 'preserve');
    fprintf('✓ 역량검사 데이터 로드 완료: %d명\n', height(competencyTestData));
catch ME
    fprintf('✗ 역량검사 데이터 로드 실패: %s\n', ME.message);
    return;
end


% -------------------------------------------------------------
% 2. 신뢰가능성 필터링
% -------------------------------------------------------------
fprintf('\n【STEP 1-1】 신뢰가능성 필터링\n');
fprintf('────────────────────────────────────────────\n');

% 신뢰가능성 컬럼 찾기
reliability_col_idx = find(contains(competencyTestData.Properties.VariableNames, '신뢰가능성'), 1);

if ~isempty(reliability_col_idx)
    colName = competencyTestData.Properties.VariableNames{reliability_col_idx};
    fprintf('▶ 신뢰가능성 컬럼 발견: %s\n', colName);

    % 신뢰불가 데이터 찾기
    reliability_data = competencyTestData{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(competencyTestData), 1);
    end

    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));

    % 신뢰가능한 데이터만 유지
    competencyTestData = competencyTestData(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능한 데이터: %d명\n', height(competencyTestData));
else
    fprintf('  ⚠ 신뢰가능성 컬럼이 없습니다. 모든 데이터를 사용합니다.\n');
end

%% 2. 23년 상반기 데이터 추가 (선택적)
fprintf('\n[1-1단계] 23년 상반기 데이터 추가 확인\n');
fprintf('----------------------------------------\n');

fileName_23_1st = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터\23년_상반기_역량진단_응답데이터.xlsx';

if exist(fileName_23_1st, 'file')
    try
        fprintf('▶ 23년 상반기 데이터 로드 중...\n');

        % 기존 allData를 임시 저장
        temp_allData = allData;
        allData = struct();

        % 23년 상반기 데이터 로드
        rawData_23_1st = readtable(fileName_23_1st, 'Sheet', '하향진단', 'VariableNamingRule', 'preserve');

        % 중복 문항 처리 (Q1, Q1_1, Q1_2 형태의 중복 제거)
        fprintf('  ▶ 23년 상반기 중복 문항 처리 중...\n');
        colNames = rawData_23_1st.Properties.VariableNames;
        qCols = colNames(startsWith(colNames, 'Q'));

        % 기본 문항만 선택 (Q1, Q2, ... Q60, _1이나 _2 접미사 제거)
        baseQCols = {};
        for i = 1:length(qCols)
            colName = qCols{i};
            if ~contains(colName, '_1') && ~contains(colName, '_2')
                baseQCols{end+1} = colName;
            end
        end

        % 비Q 컬럼들과 기본 Q 컬럼들만 유지
        nonQCols = colNames(~startsWith(colNames, 'Q'));
        keepCols = [nonQCols, baseQCols];

        % 필터링된 데이터 생성
        allData.period1.selfData = rawData_23_1st(:, keepCols);

        fprintf('    원본 Q문항: %d개 → 처리 후: %d개 (중복 제거 완료)\n', ...
                length(qCols), length(baseQCols));

        try
            allData.period1.questionInfo = readtable(fileName_23_1st, 'Sheet', '문항 정보', 'VariableNamingRule', 'preserve');
        catch
            allData.period1.questionInfo = table();
        end

        % 기존 데이터를 뒤로 밀기 (period2-5)
        fieldNames = fieldnames(temp_allData);
        for i = 1:length(fieldNames)
            oldPeriodNum = str2double(fieldNames{i}(end));
            newPeriodNum = oldPeriodNum + 1;
            allData.(sprintf('period%d', newPeriodNum)) = temp_allData.(fieldNames{i});
        end

        % periods 배열 업데이트
        periods = {'23년_상반기', '23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};

        fprintf('  ✓ 23년 상반기 데이터 추가 완료: %d명\n', height(allData.period1.selfData));
        fprintf('  ✓ 전체 Period: %d개\n', length(periods));

    catch ME
        fprintf('  ✗ 23년 상반기 데이터 로드 실패: %s\n', ME.message);
        fprintf('  → 기존 4개 시점으로 계속 진행합니다\n');
    end
else
    fprintf('  • 23년 상반기 파일 없음 - 기존 4개 시점으로 진행\n');
end

%% 3. 성과 문항 정의 및 히스토그램 생성 준비
fprintf('\n[2단계] 성과 문항 히스토그램 생성\n');
fprintf('----------------------------------------\n');

% 성과 문항 정의
performanceQuestions = struct();
performanceQuestions.period1 = {'Q3','Q4','Q5','Q22','Q23','Q45','Q46','Q51'};   % 23년 상반기
performanceQuestions.period2 = {'Q4','Q21','Q23','Q25','Q32','Q33','Q34'};       % 23년 하반기
performanceQuestions.period3 = {'Q4','Q22','Q25','Q27','Q31','Q32','Q33'};       % 24년 상반기
performanceQuestions.period4 = {'Q4','Q22','Q25','Q27','Q31','Q32','Q33'};       % 24년 하반기
performanceQuestions.period5 = {'Q4','Q22','Q25','Q27','Q31','Q32','Q33'};       % 25년 상반기

% 출력 디렉토리 설정
outputPath = 'D:\project\HR데이터\결과\성과문항_원점수_히스토그램\';
if ~exist(outputPath, 'dir')
    mkdir(outputPath);
    fprintf('출력 디렉토리 생성: %s\n', outputPath);
end

% 타임스탬프
timestamp = datestr(now, 'yyyymmdd_HHMM');

fprintf('성과 문항 정의:\n');
for p = 1:5
    if p <= length(periods)
        questions = performanceQuestions.(sprintf('period%d', p));
        fprintf('  Period %d (%s): %s\n', p, periods{p}, strjoin(questions, ', '));
    end
end

%% 4. Period별 히스토그램 생성
fprintf('\n히스토그램 생성 시작...\n');

totalPeriods = length(periods);

for p = 1:totalPeriods
    periodName = periods{p};
    periodField = sprintf('period%d', p);

    fprintf('\n▶ %s 처리 중...\n', periodName);

    % Period 데이터 확인
    if ~isfield(allData, periodField) || ~isfield(allData.(periodField), 'selfData')
        fprintf('  ⚠️ 데이터가 없습니다. 스킵.\n');
        continue;
    end

    periodData = allData.(periodField).selfData;

    if isempty(periodData)
        fprintf('  ⚠️ 빈 데이터입니다. 스킵.\n');
        continue;
    end

    % 성과 문항 리스트
    if isfield(performanceQuestions, periodField)
        perfQuestions = performanceQuestions.(periodField);
    else
        fprintf('  ⚠️ 성과 문항 정의가 없습니다. 스킵.\n');
        continue;
    end

    fprintf('  성과 문항: %s\n', strjoin(perfQuestions, ', '));

    % 존재하는 성과 문항만 필터링
    availableCols = periodData.Properties.VariableNames;
    validPerfQuestions = perfQuestions(ismember(perfQuestions, availableCols));

    if isempty(validPerfQuestions)
        fprintf('  ⚠️ 유효한 성과 문항이 없습니다. 스킵.\n');
        continue;
    end

    fprintf('  유효한 성과 문항: %d개\n', length(validPerfQuestions));

    % 히스토그램 데이터 준비
    histData = struct();
    validQuestions = {};

    for q = 1:length(validPerfQuestions)
        qName = validPerfQuestions{q};

        try
            qData = periodData.(qName);
        catch
            fprintf('    %s: 컬럼 접근 실패 (스킵)\n', qName);
            continue;
        end

        % 숫자형으로 변환
        if iscell(qData) || ischar(qData) || isstring(qData)
            qData = str2double(qData);
        end

        % NaN 제거
        validData = qData(~isnan(qData));

        if isempty(validData)
            fprintf('    %s: 유효한 데이터 없음 (스킵)\n', qName);
            continue;
        end

        % 통계 계산
        histData.(qName).data = validData;
        histData.(qName).n = length(validData);
        histData.(qName).mean = mean(validData);
        histData.(qName).std = std(validData);
        histData.(qName).min = min(validData);
        histData.(qName).max = max(validData);

        validQuestions{end+1} = qName;

        fprintf('    %s: N=%d, 평균=%.2f, SD=%.2f, 범위=%.0f~%.0f\n', ...
            qName, histData.(qName).n, histData.(qName).mean, ...
            histData.(qName).std, histData.(qName).min, histData.(qName).max);
    end

    if isempty(validQuestions)
        fprintf('  ⚠️ 그릴 수 있는 문항이 없습니다. 스킵.\n');
        continue;
    end

    % Figure 생성
    numQuestions = length(validQuestions);

    % 서브플롯 배치 계산
    if numQuestions == 1
        nRows = 1; nCols = 1;
    elseif numQuestions <= 2
        nRows = 1; nCols = 2;
    elseif numQuestions <= 4
        nRows = 2; nCols = 2;
    elseif numQuestions <= 6
        nRows = 2; nCols = 3;
    elseif numQuestions <= 9
        nRows = 3; nCols = 3;
    else
        nRows = ceil(sqrt(numQuestions));
        nCols = ceil(numQuestions / nRows);
    end

    fig = figure('Position', [100, 100, 300*nCols, 250*nRows], 'Color', 'white');

    for q = 1:length(validQuestions)
        qName = validQuestions{q};
        qInfo = histData.(qName);

        subplot(nRows, nCols, q);

        % 히스토그램 빈 설정 (불연속 점수 대응)
        uniqueVals = unique(qInfo.data);

        if length(uniqueVals) <= 15
            % 불연속 점수 대응
            sortedVals = sort(uniqueVals);
            edges = [(sortedVals(1)-0.5), (sortedVals + 0.5)'];
            edges = unique(edges);
        else
            % 자동 빈 사용
            edges = [];
        end

        % 히스토그램 그리기
        if isempty(edges)
            h = histogram(qInfo.data, 'FaceColor', [0.3, 0.6, 0.9], 'EdgeColor', 'black', 'FaceAlpha', 0.7);
        else
            h = histogram(qInfo.data, edges, 'FaceColor', [0.3, 0.6, 0.9], 'EdgeColor', 'black', 'FaceAlpha', 0.7);
        end

        % 제목 및 축 라벨
        title(sprintf('%s (N=%d, 평균=%.2f, SD=%.2f, 범위=%.0f~%.0f)', ...
            qName, qInfo.n, qInfo.mean, qInfo.std, qInfo.min, qInfo.max), ...
            'FontSize', 10, 'Interpreter', 'none');
        xlabel('원점수', 'FontSize', 9);
        ylabel('빈도', 'FontSize', 9);
        grid on;

        % x축 범위 설정
        if qInfo.max > qInfo.min
            xlim([qInfo.min-1, qInfo.max+1]);
        end

        % y축 최대값 조정
        ylim([0, max(h.Values)*1.1]);
    end

    % Figure 전체 제목
    periodNameDisplay = strrep(periodName, '_', ' ');
    sgtitle(sprintf('%s 성과 문항 원점수 분포', periodNameDisplay), ...
        'FontSize', 14, 'FontWeight', 'bold');

    % 파일 저장
    baseFileName = sprintf('hist_perf_raw_%d_%s', p, timestamp);

    pngFile = fullfile(outputPath, [baseFileName, '.png']);
    figFile = fullfile(outputPath, [baseFileName, '.fig']);

    try
        saveas(fig, pngFile, 'png');
        saveas(fig, figFile, 'fig');

        fprintf('  ✓ 히스토그램 저장 완료:\n');
        fprintf('    - %s\n', pngFile);
        fprintf('    - %s\n', figFile);
    catch saveErr
        fprintf('  ✗ 파일 저장 실패: %s\n', saveErr.message);
    end

    close(fig);
end

%% 5. 완료 메시지
fprintf('\n========================================\n');
fprintf('성과 문항 원점수 히스토그램 생성 완료\n');
fprintf('========================================\n');

fprintf('📊 처리 결과:\n');
fprintf('  - 총 Period 수: %d개\n', totalPeriods);
fprintf('  - 출력 경로: %s\n', outputPath);
fprintf('  - 파일 형식: PNG, FIG\n');
fprintf('  - 타임스탬프: %s\n', timestamp);

if totalPeriods >= 5
    fprintf('  - 23년 상반기: 포함됨\n');
else
    fprintf('  - 23년 상반기: 미포함\n');
end

fprintf('\n✅ 모든 작업이 완료되었습니다.\n');