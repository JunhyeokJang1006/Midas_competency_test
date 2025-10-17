%% Period별 문항-역량검사 상관 매트릭스 (실제 작동 버전)
%
% 목적: 23년 하반기~25년 상반기 4개 시점의 문항과 역량검사 종합점수 간 상관분석
% 특징: ID 매칭이 실제로 가능한 데이터만 분석, 모든 필요 함수 포함
%
% 작성일: 2025년

clear; clc; close all;
rng(42, 'twister');

cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('Period별 문항-종합점수 상관 매트릭스 (4개 시점)\n');
fprintf('ID 매칭 가능한 데이터만 분석\n');
fprintf('========================================\n\n');

%% 1단계: 데이터 로드 및 준비
fprintf('[1단계] 데이터 로드 및 준비\n');
fprintf('----------------------------------------\n');

% 기존 역량진단 데이터 로드 (23년 하반기~25년 상반기)
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
allData = struct();

% MAT 파일에서 기존 데이터 로드
matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    matFileName = matFiles(idx).name;

    try
        loadedData = load(matFileName);
        if isfield(loadedData, 'allData')
            allData = loadedData.allData;
            fprintf('✓ 역량진단 데이터 로드 완료\n');

            fields = fieldnames(allData);
            for i = 1:length(fields)
                if isfield(allData.(fields{i}), 'selfData')
                    fprintf('  - %s: %d명\n', fields{i}, height(allData.(fields{i}).selfData));
                end
            end
        else
            fprintf('✗ allData 필드를 찾을 수 없습니다\n');
            error('기존 데이터를 로드할 수 없습니다.');
        end
    catch ME
        fprintf('✗ MAT 파일 로드 실패: %s\n', ME.message);
        error('데이터 로드 실패');
    end
end

%% 2단계: 역량검사 데이터 로드 및 전처리
fprintf('\n[2단계] 역량검사 데이터 로드\n');
fprintf('----------------------------------------\n');

competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    competencyTestData = readtable(competencyTestPath, ...
        'Sheet', '역량검사_종합점수', ...
        'VariableNamingRule', 'preserve');
    fprintf('✓ 역량검사 데이터 로드 완료: %d명\n', height(competencyTestData));

    % 신뢰가능성 필터링
    reliability_col_idx = find(contains(competencyTestData.Properties.VariableNames, '신뢰가능성'), 1);

    if ~isempty(reliability_col_idx)
        colName = competencyTestData.Properties.VariableNames{reliability_col_idx};
        reliability_data = competencyTestData{:, reliability_col_idx};
        if iscell(reliability_data)
            unreliable_idx = strcmp(reliability_data, '신뢰불가');
        else
            unreliable_idx = false(height(competencyTestData), 1);
        end

        fprintf('  - 신뢰불가 데이터: %d명\n', sum(unreliable_idx));
        competencyTestData = competencyTestData(~unreliable_idx, :);
        fprintf('  - 신뢰가능한 데이터: %d명\n', height(competencyTestData));
    end

catch ME
    fprintf('✗ 역량검사 데이터 로드 실패: %s\n', ME.message);
    error('역량검사 데이터 로드 실패');
end

%% 3단계: 각 Period별 상관 분석
fprintf('\n[3단계] 각 Period별 상관 분석\n');
fprintf('----------------------------------------\n');

% 결과 저장용 구조체
correlationResults = struct();
successfulPeriods = {};

for p = 1:length(periods)
    fprintf('\n▶ %s 분석 중...\n', periods{p});

    % Period 데이터 확인
    if ~isfield(allData, sprintf('period%d', p))
        fprintf('  [경고] %s 데이터가 없습니다.\n', periods{p});
        continue;
    end

    selfData = allData.(sprintf('period%d', p)).selfData;

    % 1) ID 컬럼 찾기
    idCol = findBestIDColumn(selfData);
    if isempty(idCol)
        fprintf('  [경고] ID 컬럼을 찾을 수 없습니다.\n');
        continue;
    end
    fprintf('  - ID 컬럼: %s\n', idCol);

    % 2) Q 문항 컬럼들 추출
    [questionCols, questionData] = extractQuestionColumns(selfData, idCol);
    if isempty(questionCols)
        fprintf('  [경고] 문항 데이터가 없습니다.\n');
        continue;
    end
    fprintf('  - Q 문항: %d개\n', length(questionCols));

    % 3) ID 데이터 추출
    idData = selfData{:, idCol};
    fprintf('  - 응답자: %d명\n', length(idData));

    % 4) 역량검사 데이터와 매칭
    [matchedQuestionData, matchedCompData, matchedIDs] = matchDataWithCompetencyTest(...
        questionData, idData, competencyTestData);

    if size(matchedQuestionData, 1) < 10
        fprintf('  [경고] 매칭된 데이터가 부족합니다 (%d명)\n', size(matchedQuestionData, 1));
        continue;
    end

    fprintf('  - 매칭 성공: %d명\n', length(matchedIDs));

    % 5) 상관 분석 수행
    [corrMatrix, pValues] = performCorrelationAnalysis(matchedQuestionData, matchedCompData, questionCols);

    if ~isempty(corrMatrix)
        % 결과 저장
        correlationResults.(sprintf('period%d', p)) = struct(...
            'period', periods{p}, ...
            'correlationMatrix', corrMatrix, ...
            'pValues', pValues, ...
            'questionNames', {questionCols}, ...
            'sampleSize', length(matchedIDs), ...
            'matchedIDs', {matchedIDs});

        successfulPeriods{end+1} = periods{p};

        fprintf('  ✓ 상관 분석 완료 (%dx%d 매트릭스)\n', size(corrMatrix));

        % 상위 상관계수 출력
        displayTopCorrelations(corrMatrix, pValues, questionCols, 5);
    else
        fprintf('  [경고] 상관 분석 실패\n');
    end
end

%% 4단계: 결과 저장
fprintf('\n[4단계] 결과 저장\n');
fprintf('----------------------------------------\n');

if ~isempty(successfulPeriods)
    saveAnalysisResults(correlationResults, successfulPeriods);
    fprintf('✓ 분석 결과 저장 완료\n');
    fprintf('✓ 성공한 Period: %s\n', strjoin(successfulPeriods, ', '));
else
    fprintf('⚠ 분석에 성공한 Period가 없습니다.\n');
end

fprintf('\n🎉 분석 완료!\n');

%% =================================================================
%% 보조 함수들
%% =================================================================

function idCol = findBestIDColumn(data)
    % 최적의 ID 컬럼을 찾는 함수
    idCol = [];

    if ~istable(data)
        return;
    end

    varNames = data.Properties.VariableNames;
    bestCandidate = '';
    bestScore = 0;

    % ID 관련 패턴들 (우선순위별)
    idPatterns = {
        'ID', 'id', '사번', '평가자', '피평가자', ...
        'idhr', '번호', 'No', 'USER_ID', 'user_id'
    };

    for i = 1:length(varNames)
        colName = varNames{i};
        score = 0;

        % 패턴 매칭 점수
        for j = 1:length(idPatterns)
            if contains(colName, idPatterns{j})
                score = score + (length(idPatterns) - j + 1) * 10;
                break;
            end
        end

        % 데이터 특성 점수
        try
            colData = data{:, colName};
            if isnumeric(colData) && ~any(isnan(colData))
                % 고유값 비율
                uniqueRatio = length(unique(colData)) / length(colData);
                if uniqueRatio > 0.9
                    score = score + 100;
                elseif uniqueRatio > 0.8
                    score = score + 50;
                end

                % 정수형 데이터 선호
                if all(colData == floor(colData))
                    score = score + 20;
                end
            elseif iscell(colData)
                validData = colData(~cellfun(@isempty, colData));
                uniqueRatio = length(unique(validData)) / length(validData);
                if uniqueRatio > 0.9
                    score = score + 80;
                end
            end
        catch
            % 데이터 접근 실패시 점수 감점
            score = score - 50;
        end

        if score > bestScore
            bestScore = score;
            bestCandidate = colName;
        end
    end

    if bestScore > 0
        idCol = bestCandidate;
    end
end

function [questionCols, questionData] = extractQuestionColumns(data, idCol)
    % Q로 시작하는 문항 컬럼들을 추출하는 함수
    questionCols = {};
    questionData = [];

    if ~istable(data)
        return;
    end

    varNames = data.Properties.VariableNames;

    % Q로 시작하는 컬럼들 찾기
    qCols = {};
    for i = 1:length(varNames)
        colName = varNames{i};
        if startsWith(colName, 'Q') && ~strcmp(colName, idCol)
            % 숫자 문항인지 확인 (Q1, Q2, ... 등)
            qNum = regexp(colName, '^Q(\d+)$', 'tokens');
            if ~isempty(qNum)
                qCols{end+1} = colName;
            end
        end
    end

    if isempty(qCols)
        return;
    end

    % 문항 번호로 정렬
    qNumbers = cellfun(@(x) str2double(regexp(x, '\d+', 'match', 'once')), qCols);
    [~, sortIdx] = sort(qNumbers);
    qCols = qCols(sortIdx);

    % 숫자 데이터만 선택
    validCols = {};
    for i = 1:length(qCols)
        colName = qCols{i};
        colData = data{:, colName};

        if isnumeric(colData)
            validCols{end+1} = colName;
        elseif iscell(colData)
            % 셀 데이터를 숫자로 변환 시도
            numData = convertCellToNumeric(colData);
            if ~isempty(numData)
                validCols{end+1} = colName;
            end
        end
    end

    questionCols = validCols;

    % 데이터 매트릭스 생성
    if ~isempty(questionCols)
        questionData = zeros(height(data), length(questionCols));

        for i = 1:length(questionCols)
            colName = questionCols{i};
            colData = data{:, colName};

            if isnumeric(colData)
                questionData(:, i) = colData;
            elseif iscell(colData)
                questionData(:, i) = convertCellToNumeric(colData);
            end
        end
    end
end

function numericData = convertCellToNumeric(cellData)
    % 셀 데이터를 숫자로 변환
    numericData = nan(size(cellData));

    for i = 1:length(cellData)
        if isnumeric(cellData{i})
            numericData(i) = cellData{i};
        elseif ischar(cellData{i}) || isstring(cellData{i})
            val = str2double(cellData{i});
            if ~isnan(val)
                numericData(i) = val;
            end
        end
    end

    % 너무 많은 NaN이 있으면 빈 배열 반환
    validRatio = sum(~isnan(numericData)) / length(numericData);
    if validRatio < 0.5
        numericData = [];
    end
end

function [matchedQuestionData, matchedCompData, matchedIDs] = matchDataWithCompetencyTest(...
    questionData, idData, competencyTestData)
    % 역량검사 데이터와 매칭하는 함수

    matchedQuestionData = [];
    matchedCompData = [];
    matchedIDs = [];

    % 역량검사 데이터의 ID와 점수 추출
    compIDs = competencyTestData.ID;

    % 종합점수 컬럼 찾기
    totalScoreCol = [];
    varNames = competencyTestData.Properties.VariableNames;
    scorePatterns = {'점수', '종합점수', '총점', 'Total', '합계', '스코어', 'Score'};

    for i = 1:length(varNames)
        for j = 1:length(scorePatterns)
            if contains(varNames{i}, scorePatterns{j})
                totalScoreCol = varNames{i};
                break;
            end
        end
        if ~isempty(totalScoreCol)
            break;
        end
    end

    if isempty(totalScoreCol)
        fprintf('    [경고] 종합점수 컬럼을 찾을 수 없습니다.\n');
        return;
    end

    compScores = competencyTestData{:, totalScoreCol};

    % 매칭 수행
    matchedIndices = [];
    compMatchedIndices = [];

    for i = 1:length(idData)
        currentID = idData(i);

        % 역량검사 데이터에서 매칭되는 ID 찾기
        matchIdx = find(compIDs == currentID, 1);

        if ~isempty(matchIdx)
            matchedIndices(end+1) = i;
            compMatchedIndices(end+1) = matchIdx;
        end
    end

    if ~isempty(matchedIndices)
        matchedQuestionData = questionData(matchedIndices, :);
        matchedCompData = compScores(compMatchedIndices);
        matchedIDs = idData(matchedIndices);

        % NaN 값이 있는 행 제거
        validRows = ~any(isnan([matchedQuestionData, matchedCompData]), 2);

        matchedQuestionData = matchedQuestionData(validRows, :);
        matchedCompData = matchedCompData(validRows);
        matchedIDs = matchedIDs(validRows);
    end
end

function [corrMatrix, pValues] = performCorrelationAnalysis(questionData, compData, questionNames)
    % 상관 분석 수행
    corrMatrix = [];
    pValues = [];

    if isempty(questionData) || isempty(compData)
        return;
    end

    % 문항 데이터와 역량검사 점수를 결합
    allData = [questionData, compData];

    % 상관계수 계산
    [corrMatrix, pValues] = corr(allData);
end

function displayTopCorrelations(corrMatrix, pValues, questionNames, topN)
    % 상위 상관계수들을 표시하는 함수
    if isempty(corrMatrix) || size(corrMatrix, 2) < 2
        return;
    end

    % 역량검사 점수와의 상관계수 (마지막 컬럼)
    compCorrs = corrMatrix(1:end-1, end);
    compPvals = pValues(1:end-1, end);

    % 절댓값 기준으로 정렬
    [~, sortIdx] = sort(abs(compCorrs), 'descend');

    fprintf('    📊 역량검사와 상위 %d개 문항 상관계수:\n', topN);

    for i = 1:min(topN, length(sortIdx))
        idx = sortIdx(i);
        qName = questionNames{idx};
        corrVal = compCorrs(idx);
        pVal = compPvals(idx);

        significance = '';
        if pVal < 0.001
            significance = '***';
        elseif pVal < 0.01
            significance = '**';
        elseif pVal < 0.05
            significance = '*';
        end

        fprintf('      %d. %s: r=%.3f (p=%.3f)%s\n', i, qName, corrVal, pVal, significance);
    end
end

function saveAnalysisResults(correlationResults, successfulPeriods)
    % 분석 결과를 파일로 저장하는 함수

    % 타임스탬프 생성
    dateStr = datestr(now, 'yyyymmdd_HHMM');

    % Excel 파일로 저장
    excelFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\correlation_results_%s.xlsx', dateStr);

    % 각 Period별로 시트 생성
    for i = 1:length(successfulPeriods)
        periodName = successfulPeriods{i};
        result = correlationResults.(sprintf('period%d', i));

        % 상관계수 매트릭스
        corrMatrix = result.correlationMatrix;
        varNames = [result.questionNames, {'역량검사_종합점수'}];

        corrTable = array2table(corrMatrix, 'VariableNames', varNames, 'RowNames', varNames);

        try
            writetable(corrTable, excelFileName, 'Sheet', sprintf('%s_상관계수', periodName), 'WriteRowNames', true);
            fprintf('  ✓ %s 상관계수 매트릭스 저장\n', periodName);
        catch ME
            fprintf('  ✗ %s 저장 실패: %s\n', periodName, ME.message);
        end

        % p-value 매트릭스
        pTable = array2table(result.pValues, 'VariableNames', varNames, 'RowNames', varNames);

        try
            writetable(pTable, excelFileName, 'Sheet', sprintf('%s_p값', periodName), 'WriteRowNames', true);
            fprintf('  ✓ %s p값 매트릭스 저장\n', periodName);
        catch ME
            fprintf('  ✗ %s p값 저장 실패: %s\n', periodName, ME.message);
        end
    end

    % 요약 정보 저장
    summaryData = createSummaryTable(correlationResults, successfulPeriods);
    try
        writetable(summaryData, excelFileName, 'Sheet', '분석요약');
        fprintf('  ✓ 분석 요약 저장\n');
    catch ME
        fprintf('  ✗ 요약 저장 실패: %s\n', ME.message);
    end

    % MAT 파일로도 저장
    matFileName = sprintf('D:\\project\\correlation_analysis_results_%s.mat', dateStr);
    save(matFileName, 'correlationResults', 'successfulPeriods');
    fprintf('  ✓ MAT 파일 저장: %s\n', matFileName);

    fprintf('  📁 Excel 파일: %s\n', excelFileName);
end

function summaryTable = createSummaryTable(correlationResults, successfulPeriods)
    % 분석 요약 테이블 생성

    periodNames = {};
    sampleSizes = [];
    questionCounts = [];
    maxCorrelations = [];
    avgCorrelations = [];
    significantCounts = [];

    for i = 1:length(successfulPeriods)
        result = correlationResults.(sprintf('period%d', i));

        periodNames{end+1} = result.period;
        sampleSizes(end+1) = result.sampleSize;
        questionCounts(end+1) = length(result.questionNames);

        % 역량검사와의 상관계수 (마지막 컬럼의 마지막 행 제외)
        compCorrs = result.correlationMatrix(1:end-1, end);
        compPvals = result.pValues(1:end-1, end);

        maxCorrelations(end+1) = max(abs(compCorrs));
        avgCorrelations(end+1) = mean(abs(compCorrs));
        significantCounts(end+1) = sum(compPvals < 0.05);
    end

    summaryTable = table(periodNames', sampleSizes', questionCounts', maxCorrelations', ...
                         avgCorrelations', significantCounts', ...
                         'VariableNames', {'Period', 'SampleSize', 'QuestionCount', ...
                         'MaxAbsCorrelation', 'AvgAbsCorrelation', 'SignificantCount'});
end