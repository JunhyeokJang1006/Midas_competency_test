%% 성과 문항 원점수 히스토그램 생성기 (23년 상반기 보장)
%
% 목적: 각 시점별 성과 문항의 원점수 분포를 히스토그램으로 시각화
%       23년 상반기 데이터가 없으면 엑셀에서 자동 로드하여 추가
%
% 입력:
%   - 최신 competency_correlation_workspace_*.mat 파일
%   - 23년 상반기 엑셀: D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터\23년_상반기_역량진단_응답데이터.xlsx
%
% 출력:
%   - Period별 성과 문항 원점수 히스토그램 (PNG, FIG)
%   - 저장 경로: D:\project\HR데이터\결과\성과문항_원점수_히스토그램\
%
% 사용법:
%   1. MATLAB에서 D:\project\HR데이터\matlab 디렉토리로 이동
%   2. performance_items_raw_histogram_generator 실행
%
% 작성일: 2025년

clear; clc; close all;

fprintf('========================================\n');
fprintf('성과 문항 원점수 히스토그램 생성기\n');
fprintf('(23년 상반기 자동 보장)\n');
fprintf('========================================\n\n');

%% 1. 기본 설정
basePath = 'D:\project\HR데이터\matlab';
outputPath = 'D:\project\HR데이터\결과\성과문항_원점수_히스토그램\';
excelPath23H1 = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터\23년_상반기_역량진단_응답데이터.xlsx';
sheetName23H1 = '하향진단';

% 출력 디렉토리 생성
if ~exist(outputPath, 'dir')
    mkdir(outputPath);
    fprintf('출력 디렉토리 생성: %s\n', outputPath);
end

% 작업 디렉토리 이동
cd(basePath);

%% 2. 성과 문항 정의
performanceQuestions = struct();
performanceQuestions.period1 = {'Q3','Q4','Q5','Q22','Q23','Q45','Q46','Q51'};   % 23년 상반기
performanceQuestions.period2 = {'Q4','Q21','Q23','Q25','Q32','Q33','Q34'};       % 23년 하반기
performanceQuestions.period3 = {'Q4','Q22','Q25','Q27','Q31','Q32','Q33'};       % 24년 상반기
performanceQuestions.period4 = {'Q4','Q22','Q25','Q27','Q31','Q32','Q33'};       % 24년 하반기
performanceQuestions.period5 = {'Q4','Q22','Q25','Q27','Q31','Q32','Q33'};       % 25년 상반기

defaultPeriods = {'23년_상반기','23년_하반기','24년_상반기','24년_하반기','25년_상반기'};

fprintf('성과 문항 정의 완료:\n');
for p = 1:5
    questions = performanceQuestions.(sprintf('period%d', p));
    fprintf('  Period %d (%s): %s\n', p, defaultPeriods{p}, strjoin(questions, ', '));
end

%% 3. MAT 파일 로드
fprintf('\n[1단계] MAT 파일 로드\n');
fprintf('----------------------------------------\n');

matFiles = dir('competency_correlation_workspace_*.mat');
if isempty(matFiles)
    error('MAT 파일을 찾을 수 없습니다: competency_correlation_workspace_*.mat');
end

[~, idx] = max([matFiles.datenum]);
matFileName = matFiles(idx).name;

try
    loadedData = load(matFileName);
    fprintf('✓ MAT 파일 로드 성공: %s\n', matFileName);

    if ~isfield(loadedData, 'allData')
        error('MAT 파일에 allData가 없습니다.');
    end

    allData = loadedData.allData;

    if isfield(loadedData, 'periods')
        periods = loadedData.periods;
    else
        periods = {};
        fprintf('  periods 필드가 없어 기본값으로 초기화합니다.\n');
    end

    fprintf('  기존 Period 수: %d개\n', length(fieldnames(allData)));

catch ME
    error('MAT 파일 로드 실패: %s', ME.message);
end

%% 4. 23년 상반기 보장 로직
fprintf('\n[2단계] 23년 상반기 데이터 보장\n');
fprintf('----------------------------------------\n');

need23H1 = false;

% period1 확인
if ~isfield(allData, 'period1')
    need23H1 = true;
    fprintf('⚠️ period1이 없습니다.\n');
elseif ~isfield(allData.period1, 'selfData') || isempty(allData.period1.selfData)
    need23H1 = true;
    fprintf('⚠️ period1.selfData가 비어있습니다.\n');
end

% periods 배열에서 23년_상반기 확인
if ~ismember('23년_상반기', periods)
    need23H1 = true;
    fprintf('⚠️ periods에 "23년_상반기"가 없습니다.\n');
end

if need23H1
    fprintf('📁 23년 상반기 데이터를 엑셀에서 로드합니다...\n');

    try
        % 23년 상반기 엑셀 로드
        if ~exist(excelPath23H1, 'file')
            error('23년 상반기 엑셀 파일이 없습니다: %s', excelPath23H1);
        end

        raw23H1 = readtable(excelPath23H1, 'Sheet', sheetName23H1, 'VariableNamingRule', 'preserve');
        fprintf('  ✓ 엑셀 로드 성공: %d행 x %d열\n', height(raw23H1), width(raw23H1));

        % 중복 Q문항 정리
        colNames = raw23H1.Properties.VariableNames;

        % Q문항 찾기 (Q로 시작하고 _text가 아닌 것)
        qCols = colNames(startsWith(colNames, 'Q') & ~contains(colNames, '_text'));

        % 중복 제거: Q1_1, Q1_2가 있으면 Q1만 남기기
        uniqueQCols = {};
        for i = 1:length(qCols)
            colName = qCols{i};
            % 기본형 추출 (Q숫자 부분만)
            baseMatch = regexp(colName, '^Q\d+', 'match');
            if ~isempty(baseMatch)
                baseQ = baseMatch{1};
                if ~ismember(baseQ, uniqueQCols)
                    uniqueQCols{end+1} = baseQ;
                end
            end
        end

        % 실제 존재하는 기본형 Q컬럼만 선택
        actualQCols = {};
        for i = 1:length(uniqueQCols)
            if ismember(uniqueQCols{i}, colNames)
                actualQCols{end+1} = uniqueQCols{i};
            end
        end

        fprintf('  정리된 Q문항: %d개 (%s)\n', length(actualQCols), ...
            strjoin(actualQCols(1:min(5, length(actualQCols))), ', '));

        % 비Q컬럼 + 정리된 Q컬럼만 선택
        nonQCols = colNames(~startsWith(colNames, 'Q'));
        finalCols = [nonQCols, actualQCols];

        % allData.period1에 추가
        allData.period1.selfData = raw23H1(:, finalCols);

        % periods 배열 업데이트
        if isempty(periods)
            periods = defaultPeriods;
        elseif ~ismember('23년_상반기', periods)
            periods = ['23년_상반기', periods];
        end

        fprintf('  ✓ 23년 상반기 데이터 추가 완료\n');
        fprintf('  ✓ 최종 컬럼 수: %d개\n', length(finalCols));

    catch ME
        warning('23년 상반기 로드 실패: %s', ME.message);
        fprintf('  기존 데이터로 계속 진행합니다.\n');
    end
else
    fprintf('✓ 23년 상반기 데이터가 이미 존재합니다.\n');
end

%% 5. 각 Period별 히스토그램 생성
fprintf('\n[3단계] Period별 히스토그램 생성\n');
fprintf('----------------------------------------\n');

timestamp = datestr(now, 'yyyymmdd_HHMM');
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
        qData = periodData.(qName);

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

        % 히스토그램 빈 설정
        uniqueVals = unique(qInfo.data);

        % getQuestionScaleInfo 함수 존재 확인
        if exist('getQuestionScaleInfo', 'file') == 2
            try
                scaleInfo = getQuestionScaleInfo({qName}, p);
                if isfield(scaleInfo, qName) && isfield(scaleInfo.(qName), 'theoreticalMin') && ...
                   isfield(scaleInfo.(qName), 'theoreticalMax')
                    minVal = scaleInfo.(qName).theoreticalMin;
                    maxVal = scaleInfo.(qName).theoreticalMax;
                    edges = (minVal-0.5):(maxVal+0.5);
                else
                    edges = [];
                end
            catch
                edges = [];
            end
        else
            edges = [];
        end

        % 대체 빈 전략
        if isempty(edges)
            if length(uniqueVals) <= 15
                % 불연속 점수 대응
                sortedVals = sort(uniqueVals);
                edges = [(sortedVals(1)-0.5), (sortedVals + 0.5)'];
                edges = unique(edges);
            else
                % 자동 빈 사용
                edges = [];
            end
        end

        % 히스토그램 그리기
        if isempty(edges)
            histogram(qInfo.data, 'FaceColor', [0.3, 0.6, 0.9], 'EdgeColor', 'black', 'FaceAlpha', 0.7);
        else
            histogram(qInfo.data, edges, 'FaceColor', [0.3, 0.6, 0.9], 'EdgeColor', 'black', 'FaceAlpha', 0.7);
        end

        % 제목 및 축 라벨
        title(sprintf('%s (N=%d, 평균=%.2f, SD=%.2f, 범위=%.0f~%.0f)', ...
            qName, qInfo.n, qInfo.mean, qInfo.std, qInfo.min, qInfo.max), ...
            'FontSize', 10, 'Interpreter', 'none');
        xlabel('원점수', 'FontSize', 9);
        ylabel('빈도', 'FontSize', 9);
        grid on;

        % x축 범위 설정
        xlim([qInfo.min-1, qInfo.max+1]);
    end

    % Figure 전체 제목
    periodNameDisplay = strrep(periodName, '_', ' ');
    sgtitle(sprintf('%s 성과 문항 원점수 분포', periodNameDisplay), ...
        'FontSize', 14, 'FontWeight', 'bold');

    % 파일 저장
    baseFileName = sprintf('hist_perf_raw_%d_%s', p, timestamp);

    pngFile = fullfile(outputPath, [baseFileName, '.png']);
    figFile = fullfile(outputPath, [baseFileName, '.fig']);

    saveas(fig, pngFile, 'png');
    saveas(fig, figFile, 'fig');

    fprintf('  ✓ 히스토그램 저장 완료:\n');
    fprintf('    - %s\n', pngFile);
    fprintf('    - %s\n', figFile);

    close(fig);
end

%% 6. 완료 메시지
fprintf('\n========================================\n');
fprintf('성과 문항 원점수 히스토그램 생성 완료\n');
fprintf('========================================\n');

fprintf('📊 처리 결과:\n');
fprintf('  - 총 Period 수: %d개\n', totalPeriods);
fprintf('  - 출력 경로: %s\n', outputPath);
fprintf('  - 파일 형식: PNG, FIG\n');
fprintf('  - 타임스탬프: %s\n', timestamp);

if need23H1
    fprintf('  - 23년 상반기: 엑셀에서 자동 로드됨\n');
else
    fprintf('  - 23년 상반기: 기존 MAT에서 사용됨\n');
end

fprintf('\n✅ 모든 작업이 완료되었습니다.\n');