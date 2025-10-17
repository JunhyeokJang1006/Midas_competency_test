%% 각 Period별 문항과 역량검사 종합점수 간 상관 매트릭스 생성 (완전 버전)
%
% 목적: 각 시점별로 수집된 문항들과 역량검사 종합점수 간의
%       전체 상관 매트릭스를 생성하고 분석
%
% 작성일: 2025년
% 특징: 5개 Period 지원, 모든 필요 함수 포함

clear; clc; close all;
rng(42, 'twister');

cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('Period별 문항-종합점수 상관 매트릭스 생성 (5개 시점)\n');
fprintf('========================================\n\n');

%% 1. 기존 분석 결과 로드
fprintf('[1단계] 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

% 기존 역량진단 데이터 로드
consolidatedScores = [];
allData = struct();
periods = {'23년_상반기', '23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};

% MAT 파일에서 로드 시도
matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    matFileName = matFiles(idx).name;

    try
        loadedData = load(matFileName);
        if isfield(loadedData, 'allData')
            existingAllData = loadedData.allData;
            fprintf('✓ 기존 역량진단 데이터 로드 완료\n');

            % 기존 데이터 구조 확인
            fields = fieldnames(existingAllData);
            fprintf('  - 기존 로드된 Period 수: %d개\n', length(fields));
            for i = 1:length(fields)
                if isfield(existingAllData.(fields{i}), 'selfData')
                    fprintf('  - %s: %d명\n', fields{i}, height(existingAllData.(fields{i}).selfData));
                end
            end

            % 기존 데이터를 period2-5로 이동 (23년 하반기부터)
            existingFields = fieldnames(existingAllData);
            for i = 1:length(existingFields)
                newPeriodNum = i + 1; % period1은 23년 상반기용으로 비워둠
                allData.(sprintf('period%d', newPeriodNum)) = existingAllData.(existingFields{i});
            end
        else
            fprintf('✗ allData를 찾을 수 없습니다\n');
        end
    catch ME
        fprintf('✗ MAT 파일 로드 실패: %s\n', ME.message);
    end
end

%% 2. 23년 상반기 데이터 추가
fprintf('\n[2단계] 23년 상반기 데이터 추가\n');
fprintf('----------------------------------------\n');

fileName_23_1st = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터\23년_상반기_역량진단_응답데이터.xlsx';

if exist(fileName_23_1st, 'file')
    try
        fprintf('▶ 23년 상반기 데이터 로드 중...\n');

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

        fprintf('  ✓ 23년 상반기 데이터 추가 완료: %d명\n', height(allData.period1.selfData));

    catch ME
        fprintf('  ✗ 23년 상반기 데이터 로드 실패: %s\n', ME.message);
        fprintf('  → 4개 시점으로 계속 진행합니다\n');
        periods = periods(2:end); % 첫 번째 시점 제거
    end
else
    fprintf('  • 23년 상반기 파일 없음 - 4개 시점으로 진행\n');
    periods = periods(2:end); % 첫 번째 시점 제거

    % 기존 데이터의 period 번호를 조정
    if exist('existingAllData', 'var')
        allData = existingAllData; % 원래 번호 유지
    end
end

%% 3. 역량검사 데이터 로드
fprintf('\n[3단계] 역량검사 데이터 로드\n');
fprintf('----------------------------------------\n');

competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    competencyTestData = readtable(competencyTestPath, ...
        'Sheet', '역량검사_종합점수', ...
        'VariableNamingRule', 'preserve');
    fprintf('✓ 역량검사 데이터 로드 완료: %d명\n', height(competencyTestData));
catch ME
    fprintf('✗ 역량검사 데이터 로드 실패: %s\n', ME.message);
    return;
end

% 신뢰가능성 필터링
reliability_col_idx = find(contains(competencyTestData.Properties.VariableNames, '신뢰가능성'), 1);

if ~isempty(reliability_col_idx)
    colName = competencyTestData.Properties.VariableNames{reliability_col_idx};
    fprintf('▶ 신뢰가능성 컬럼 발견: %s\n', colName);

    reliability_data = competencyTestData{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(competencyTestData), 1);
    end

    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));

    competencyTestData = competencyTestData(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능한 데이터: %d명\n', height(competencyTestData));
else
    fprintf('  ⚠ 신뢰가능성 컬럼이 없습니다. 모든 데이터를 사용합니다.\n');
end

%% 4. 각 Period별 성과 관련 문항 정의
performanceQuestions = struct();

if length(periods) == 5
    performanceQuestions.period1 = {'Q3', 'Q4', 'Q5', 'Q22', 'Q23', 'Q45', 'Q46', 'Q51'};  % 23년 상반기
    performanceQuestions.period2 = {'Q4', 'Q21', 'Q23', 'Q25', 'Q32', 'Q33', 'Q34'};       % 23년 하반기
    performanceQuestions.period3 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};       % 24년 상반기
    performanceQuestions.period4 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};       % 24년 하반기
    performanceQuestions.period5 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};       % 25년 상반기
else
    performanceQuestions.period1 = {'Q4', 'Q21', 'Q23', 'Q25', 'Q32', 'Q33', 'Q34'};       % 23년 하반기
    performanceQuestions.period2 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};       % 24년 상반기
    performanceQuestions.period3 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};       % 24년 하반기
    performanceQuestions.period4 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};       % 25년 상반기
end

%% 5. 각 Period별 상관 분석
fprintf('\n[4단계] 각 Period별 문항 데이터 분석\n');
fprintf('----------------------------------------\n');

% 결과 저장용 구조체
correlationMatrices = struct();

for p = 1:length(periods)
    fprintf('\n▶ %s 처리 중...\n', periods{p});

    % Period별 데이터 확인
    if ~isfield(allData, sprintf('period%d', p))
        fprintf('  [경고] %s 데이터를 찾을 수 없습니다.\n', periods{p});
        continue;
    end

    selfData = allData.(sprintf('period%d', p)).selfData;

    % ID 컬럼 찾기
    idCol = findIDColumn(selfData);
    if isempty(idCol)
        fprintf('  [경고] ID 컬럼을 찾을 수 없습니다.\n');
        continue;
    end

    % 문항 컬럼들 추출
    [questionCols, questionData] = extractQuestionData(selfData, idCol);

    if isempty(questionCols)
        fprintf('  [경고] 문항 데이터를 찾을 수 없습니다.\n');
        continue;
    end

    % ID 추출 및 표준화
    responseIDs = extractAndStandardizeIDs(selfData{:, idCol});

    fprintf('  - 발견된 문항: %d개\n', length(questionCols));
    fprintf('  - 응답자: %d명\n', length(responseIDs));

    % 역량검사 데이터와 매칭
    [matchedData, matchedIDs, sampleSize] = matchWithCompetencyTest(questionData, responseIDs, competencyTestData);

    if sampleSize < 5
        fprintf('  [경고] 매칭된 데이터가 부족합니다 (%d명)\n', sampleSize);
        continue;
    end

    fprintf('  - 매칭된 응답자: %d명\n', sampleSize);

    % 상관 분석 수행
    [correlationMatrix, pValues, cleanData, variableNames] = performCorrelationAnalysis(matchedData, questionCols);

    if ~isempty(correlationMatrix)
        % 결과 저장
        correlationMatrices.(sprintf('period%d', p)) = struct(...
            'correlationMatrix', correlationMatrix, ...
            'pValues', pValues, ...
            'variableNames', {variableNames}, ...
            'questionNames', {questionCols}, ...
            'sampleSize', size(cleanData, 1), ...
            'cleanData', cleanData, ...
            'cleanIDs', {matchedIDs});

        fprintf('  ✓ 상관 매트릭스 계산 완료 (%dx%d)\n', ...
            size(correlationMatrix, 1), size(correlationMatrix, 2));

        % 주요 상관계수 출력
        displayTopCorrelations(correlationMatrix, pValues, questionCols);

        % 성과 관련 문항의 상관계수 별도 출력
        fprintf('\n');
        perfQuestions = performanceQuestions.(sprintf('period%d', p));
        displayPerformanceQuestionCorrelations(correlationMatrix, pValues, questionCols, perfQuestions);
    end
end

%% 6. 결과 저장
fprintf('\n[5단계] 결과 저장\n');
fprintf('----------------------------------------\n');

% 백업 및 저장 로직
saveResultsToFiles(correlationMatrices, periods);

fprintf('\n🎉 분석이 성공적으로 완료되었습니다!\n');

%% =================================================================
%% 보조 함수들
%% =================================================================

function idCol = findIDColumn(data)
    % ID 컬럼을 찾는 함수
    idCol = [];

    if ~istable(data)
        return;
    end

    varNames = data.Properties.VariableNames;

    % ID 관련 가능한 컬럼명들 (더 포괄적으로)
    idPatterns = {'ID', 'id', 'Id', '사번', '번호', 'No', 'USER_ID', 'idhr', '평가자', '피평가자', '사원번호'};

    for i = 1:length(varNames)
        varName = varNames{i};
        for j = 1:length(idPatterns)
            if contains(varName, idPatterns{j})
                % 실제 데이터가 ID로 보이는지 확인
                colData = data{:, varName};
                if isnumeric(colData) || (iscell(colData) && ~isempty(colData))
                    % 고유값 비율이 높으면 ID 컬럼으로 판단
                    if isnumeric(colData)
                        uniqueRatio = length(unique(colData(~isnan(colData)))) / length(colData(~isnan(colData)));
                    else
                        validData = colData(~cellfun(@isempty, colData));
                        uniqueRatio = length(unique(validData)) / length(validData);
                    end

                    if uniqueRatio > 0.8  % 80% 이상이 고유값이면 ID로 간주
                        idCol = varName;
                        return;
                    end
                end
            end
        end
    end

    % 첫 번째 컬럼이 ID일 가능성
    if isempty(idCol) && length(varNames) > 0
        firstCol = data{:, 1};
        if (isnumeric(firstCol) && length(unique(firstCol)) == length(firstCol)) || ...
           (iscell(firstCol) && length(unique(firstCol)) == length(firstCol))
            idCol = varNames{1};
        end
    end

    % 모든 방법이 실패하면 첫 번째 컬럼 사용
    if isempty(idCol) && length(varNames) > 0
        fprintf('    [주의] ID 컬럼을 찾지 못해 첫 번째 컬럼 "%s"을 사용합니다.\n', varNames{1});
        idCol = varNames{1};
    end
end

function [questionCols, questionData] = extractQuestionData(data, idCol)
    % 문항 데이터를 추출하는 함수
    questionCols = {};
    questionData = [];

    if ~istable(data)
        return;
    end

    varNames = data.Properties.VariableNames;

    % Q로 시작하는 컬럼들 찾기
    qCols = varNames(startsWith(varNames, 'Q'));

    % 숫자형 또는 변환 가능한 컬럼들만 선택
    for i = 1:length(qCols)
        colName = qCols{i};
        colData = data.(colName);

        if isnumeric(colData)
            questionCols{end+1} = colName;
        elseif iscell(colData)
            % 셀 데이터가 숫자로 변환 가능한지 확인
            converted = convertCellToNumeric(colData);
            if ~isempty(converted)
                questionCols{end+1} = colName;
            end
        end
    end

    % 문항 데이터 매트릭스 생성
    if ~isempty(questionCols)
        questionData = zeros(height(data), length(questionCols));

        for i = 1:length(questionCols)
            colName = questionCols{i};
            colData = data.(colName);

            if isnumeric(colData)
                questionData(:, i) = colData;
            elseif iscell(colData)
                questionData(:, i) = convertCellToNumeric(colData);
            end
        end
    end
end

function numericData = convertCellToNumeric(cellData)
    % 셀 데이터를 숫자로 변환하는 함수
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

function standardizedIDs = extractAndStandardizeIDs(idData)
    % ID를 추출하고 표준화하는 함수
    standardizedIDs = {};

    if isnumeric(idData)
        for i = 1:length(idData)
            standardizedIDs{i} = sprintf('%.0f', idData(i));
        end
    elseif iscell(idData)
        for i = 1:length(idData)
            if isnumeric(idData{i})
                standardizedIDs{i} = sprintf('%.0f', idData{i});
            elseif ischar(idData{i}) || isstring(idData{i})
                standardizedIDs{i} = char(idData{i});
            else
                standardizedIDs{i} = sprintf('ID_%d', i);
            end
        end
    else
        % 기본값
        for i = 1:length(idData)
            standardizedIDs{i} = sprintf('ID_%d', i);
        end
    end
end

function standardizedData = standardizeQuestionScales(data, questions, periodNum)
    % 리커트 척도를 표준화하는 함수
    standardizedData = data;

    % 각 문항의 범위를 0-1로 표준화
    for i = 1:size(data, 2)
        colData = data(:, i);
        validData = colData(~isnan(colData));

        if ~isempty(validData)
            minVal = min(validData);
            maxVal = max(validData);

            if maxVal > minVal
                standardizedData(:, i) = (colData - minVal) / (maxVal - minVal);
            else
                standardizedData(:, i) = colData * 0; % 모든 값이 같으면 0
            end
        end
    end
end

function [matchedData, matchedIDs, sampleSize] = matchWithCompetencyTest(questionData, responseIDs, competencyTestData)
    % 역량검사 데이터와 매칭하는 함수
    matchedData = [];
    matchedIDs = {};
    sampleSize = 0;

    % 역량검사 데이터의 ID 컬럼 찾기
    compIdCol = findIDColumn(competencyTestData);
    if isempty(compIdCol)
        fprintf('  [경고] 역량검사 데이터에서 ID 컬럼을 찾을 수 없습니다.\n');
        return;
    end

    % 역량검사 ID 표준화
    compIDs = extractAndStandardizeIDs(competencyTestData{:, compIdCol});

    % 종합점수 컬럼 찾기
    totalScoreCol = [];
    varNames = competencyTestData.Properties.VariableNames;
    scorePatterns = {'종합점수', '총점', 'Total', '합계'};

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
        fprintf('  [경고] 종합점수 컬럼을 찾을 수 없습니다.\n');
        return;
    end

    % 매칭되는 ID 찾기
    matchedIndices = [];
    compMatchedIndices = [];

    for i = 1:length(responseIDs)
        respID = responseIDs{i};

        for j = 1:length(compIDs)
            if strcmp(respID, compIDs{j})
                matchedIndices(end+1) = i;
                compMatchedIndices(end+1) = j;
                break;
            end
        end
    end

    if isempty(matchedIndices)
        fprintf('  [경고] 매칭되는 ID가 없습니다.\n');
        return;
    end

    % 매칭된 데이터 생성
    matchedQuestionData = questionData(matchedIndices, :);
    matchedTotalScores = competencyTestData{compMatchedIndices, totalScoreCol};

    % 숫자 변환
    if iscell(matchedTotalScores)
        numericTotalScores = convertCellToNumeric(matchedTotalScores);
    else
        numericTotalScores = matchedTotalScores;
    end

    % NaN 제거
    validRows = ~any(isnan([matchedQuestionData, numericTotalScores]), 2);

    matchedData = [matchedQuestionData(validRows, :), numericTotalScores(validRows)];
    matchedIDs = responseIDs(matchedIndices(validRows));
    sampleSize = sum(validRows);
end

function [correlationMatrix, pValues, cleanData, variableNames] = performCorrelationAnalysis(matchedData, questionCols)
    % 상관 분석을 수행하는 함수
    correlationMatrix = [];
    pValues = [];
    cleanData = [];
    variableNames = {};

    if isempty(matchedData) || size(matchedData, 2) < 2
        return;
    end

    % 변수명 설정
    variableNames = [questionCols, {'종합점수'}];

    % 완전한 케이스만 사용
    completeRows = ~any(isnan(matchedData), 2);
    cleanData = matchedData(completeRows, :);

    if size(cleanData, 1) < 3
        return;
    end

    % 상관계수 및 p값 계산
    [correlationMatrix, pValues] = corr(cleanData);
end

function displayTopCorrelations(corrMatrix, pValues, questionCols)
    % 상위 상관계수를 표시하는 함수
    if isempty(corrMatrix)
        return;
    end

    % 종합점수와의 상관계수만 추출 (마지막 컬럼)
    if size(corrMatrix, 2) > 1
        totalScoreCorr = corrMatrix(1:end-1, end);
        totalScorePvals = pValues(1:end-1, end);

        % 절댓값 기준 정렬
        [sortedCorr, sortIdx] = sort(abs(totalScoreCorr), 'descend');

        fprintf('  📊 종합점수와의 상위 상관계수:\n');
        topN = min(5, length(sortedCorr));

        for i = 1:topN
            idx = sortIdx(i);
            qName = questionCols{idx};
            corrVal = totalScoreCorr(idx);
            pVal = totalScorePvals(idx);

            significance = '';
            if pVal < 0.001
                significance = '***';
            elseif pVal < 0.01
                significance = '**';
            elseif pVal < 0.05
                significance = '*';
            end

            fprintf('    %d. %s: r=%.3f (p=%.3f)%s\n', i, qName, corrVal, pVal, significance);
        end
    end
end

function displayPerformanceQuestionCorrelations(corrMatrix, pValues, questionCols, perfQuestions)
    % 성과 관련 문항의 상관계수를 표시하는 함수
    if isempty(corrMatrix) || isempty(perfQuestions)
        return;
    end

    fprintf('  🎯 성과 관련 문항과 종합점수의 상관계수:\n');

    % 종합점수와의 상관계수 (마지막 컬럼)
    if size(corrMatrix, 2) > 1
        totalScoreCorr = corrMatrix(1:end-1, end);
        totalScorePvals = pValues(1:end-1, end);

        perfCount = 0;
        for i = 1:length(questionCols)
            qName = questionCols{i};

            % 성과 관련 문항인지 확인
            if any(strcmp(qName, perfQuestions))
                corrVal = totalScoreCorr(i);
                pVal = totalScorePvals(i);

                significance = '';
                if pVal < 0.001
                    significance = '***';
                elseif pVal < 0.01
                    significance = '**';
                elseif pVal < 0.05
                    significance = '*';
                end

                fprintf('    - %s: r=%.3f (p=%.3f)%s\n', qName, corrVal, pVal, significance);
                perfCount = perfCount + 1;
            end
        end

        if perfCount == 0
            fprintf('    (해당하는 성과 문항이 없습니다)\n');
        end
    end
end

function summaryTable = createSummaryTable(correlationMatrices, periods)
    % 요약 테이블을 생성하는 함수
    periodNames = {};
    sampleSizes = [];
    questionCounts = [];
    maxCorrelations = [];
    minCorrelations = [];
    avgCorrelations = [];
    significantCorrelations = [];

    fieldNames = fieldnames(correlationMatrices);

    for i = 1:length(fieldNames)
        fieldName = fieldNames{i};
        periodNum = str2double(fieldName(end));
        result = correlationMatrices.(fieldName);

        periodNames{end+1} = periods{periodNum};
        sampleSizes(end+1) = result.sampleSize;
        questionCounts(end+1) = length(result.questionNames);

        % 종합점수와의 상관계수 (마지막 컬럼, 대각선 제외)
        if size(result.correlationMatrix, 2) > 1
            totalCorr = result.correlationMatrix(1:end-1, end);
            totalPvals = result.pValues(1:end-1, end);

            maxCorrelations(end+1) = max(totalCorr);
            minCorrelations(end+1) = min(totalCorr);
            avgCorrelations(end+1) = mean(totalCorr);
            significantCorrelations(end+1) = sum(totalPvals < 0.05);
        else
            maxCorrelations(end+1) = NaN;
            minCorrelations(end+1) = NaN;
            avgCorrelations(end+1) = NaN;
            significantCorrelations(end+1) = 0;
        end
    end

    summaryTable = table(periodNames', sampleSizes', questionCounts', ...
                         maxCorrelations', minCorrelations', avgCorrelations', significantCorrelations', ...
                         'VariableNames', {'Period', 'SampleSize', 'QuestionCount', ...
                         'MaxCorrelation', 'MinCorrelation', 'AvgCorrelation', 'SignificantCount'});
end

function saveResultsToFiles(correlationMatrices, periods)
    % 결과를 파일로 저장하는 함수

    % 백업 폴더 확인 및 생성
    backupDir = 'D:\project\HR데이터\결과\성과종합점수&역검\backup';
    if ~exist(backupDir, 'dir')
        mkdir(backupDir);
        fprintf('✓ 백업 폴더 생성: %s\n', backupDir);
    end

    % 기존 파일들을 백업 폴더로 이동
    existingFiles = dir('D:\project\HR데이터\결과\성과종합점수&역검\correlation_matrices_by_period_*.xlsx');
    if ~isempty(existingFiles)
        fprintf('▶ 기존 파일 백업 중...\n');
        for i = 1:length(existingFiles)
            oldFile = fullfile(existingFiles(i).folder, existingFiles(i).name);
            newFile = fullfile(backupDir, existingFiles(i).name);
            try
                movefile(oldFile, newFile);
                fprintf('  • %s → 백업 완료\n', existingFiles(i).name);
            catch ME
                fprintf('  ✗ %s 백업 실패: %s\n', existingFiles(i).name, ME.message);
            end
        end
    end

    % MAT 파일도 백업
    existingMatFiles = dir('D:\project\correlation_matrices_workspace_*.mat');
    if ~isempty(existingMatFiles)
        fprintf('▶ 기존 MAT 파일 백업 중...\n');
        for i = 1:length(existingMatFiles)
            oldFile = fullfile(existingMatFiles(i).folder, existingMatFiles(i).name);
            newFile = fullfile(backupDir, existingMatFiles(i).name);
            try
                movefile(oldFile, newFile);
                fprintf('  • %s → 백업 완료\n', existingMatFiles(i).name);
            catch ME
                fprintf('  ✗ %s 백업 실패: %s\n', existingMatFiles(i).name, ME.message);
            end
        end
    end

    % 새 결과 파일명 생성
    dateStr = datestr(now, 'yyyymmdd_HHMM');
    outputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\correlation_matrices_by_period_%s.xlsx', dateStr);

    % 각 Period별 상관 매트릭스를 별도 시트에 저장
    savedSheets = {};
    fieldNames = fieldnames(correlationMatrices);

    for i = 1:length(fieldNames)
        fieldName = fieldNames{i};
        periodNum = str2double(fieldName(end));
        result = correlationMatrices.(fieldName);

        % 상관 매트릭스를 테이블로 변환
        corrTable = array2table(result.correlationMatrix, ...
            'VariableNames', result.variableNames, ...
            'RowNames', result.variableNames);

        % p-value 매트릭스를 테이블로 변환
        pTable = array2table(result.pValues, ...
            'VariableNames', result.variableNames, ...
            'RowNames', result.variableNames);

        % 시트명 설정
        corrSheetName = sprintf('%s_상관계수', periods{periodNum});
        pSheetName = sprintf('%s_p값', periods{periodNum});

        try
            % 상관계수 매트릭스 저장
            writetable(corrTable, outputFileName, 'Sheet', corrSheetName, 'WriteRowNames', true);

            % p-value 매트릭스 저장
            writetable(pTable, outputFileName, 'Sheet', pSheetName, 'WriteRowNames', true);

            savedSheets{end+1} = corrSheetName;
            savedSheets{end+1} = pSheetName;

            fprintf('✓ %s 매트릭스 저장 완료\n', periods{periodNum});

        catch ME
            fprintf('✗ %s 매트릭스 저장 실패: %s\n', periods{periodNum}, ME.message);
        end
    end

    % 요약 테이블 생성 및 저장
    if ~isempty(fieldNames)
        summaryTable = createSummaryTable(correlationMatrices, periods);

        try
            writetable(summaryTable, outputFileName, 'Sheet', '분석요약');
            savedSheets{end+1} = '분석요약';
            fprintf('✓ 분석 요약 저장 완료\n');
        catch ME
            fprintf('✗ 분석 요약 저장 실패: %s\n', ME.message);
        end
    else
        summaryTable = table();
    end

    % MAT 파일로도 저장
    matFileName = sprintf('D:\\project\\correlation_matrices_workspace_%s.mat', dateStr);
    if ~isempty(fieldNames)
        save(matFileName, 'correlationMatrices', 'periods', 'summaryTable');
        fprintf('✓ MAT 파일 저장 완료: %s\n', matFileName);
    end

    fprintf('\n📁 저장된 시트: %s\n', strjoin(savedSheets, ', '));
    fprintf('📁 Excel 파일: %s\n', outputFileName);
end