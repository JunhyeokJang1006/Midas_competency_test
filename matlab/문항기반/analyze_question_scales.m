%% 문항별 척도 분석 및 탐지 스크립트
%
% 목적: 각 문항의 리커트 척도를 자동으로 탐지하고 분석
% 작성일: 2025년
%
% 사용법: analyze_question_scales()

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('문항별 척도 자동 탐지 및 분석\n');
fprintf('========================================\n\n');

%% 1. 기존 데이터 로드
fprintf('[1단계] 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

% MAT 파일에서 로드
matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    matFileName = matFiles(idx).name;

    try
        loadedData = load(matFileName);
        if isfield(loadedData, 'allData')
            allData = loadedData.allData;
            fprintf('✓ 역량진단 데이터 로드 완료: %s\n', matFileName);
        else
            fprintf('✗ allData를 찾을 수 없습니다\n');
            return;
        end

        if isfield(loadedData, 'periods')
            periods = loadedData.periods;
        else
            periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
        end

    catch ME
        fprintf('✗ MAT 파일 로드 실패: %s\n', ME.message);
        return;
    end
else
    fprintf('✗ 분석 결과 파일을 찾을 수 없습니다\n');
    fprintf('  먼저 factor_analysis_by_period.m을 실행하세요\n');
    return;
end

%% 2. 각 Period별 문항 척도 분석
fprintf('\n[2단계] Period별 문항 척도 분석\n');
fprintf('----------------------------------------\n');

allScaleAnalysis = table();
periodFields = fieldnames(allData);

for p = 1:length(periodFields)
    fieldName = periodFields{p};
    periodNum = str2double(fieldName(end));

    if periodNum > length(periods)
        continue;
    end

    fprintf('\n▶ %s 분석 중...\n', periods{periodNum});

    selfData = allData.(fieldName).selfData;

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

    fprintf('  - 발견된 문항: %d개\n', length(questionCols));

    % 문항별 척도 분석
    periodAnalysis = analyzeQuestionScales(questionData, questionCols, periods{periodNum});
    allScaleAnalysis = [allScaleAnalysis; periodAnalysis];
end

%% 3. 전체 문항 척도 요약
fprintf('\n[3단계] 전체 문항 척도 요약\n');
fprintf('----------------------------------------\n');

if height(allScaleAnalysis) > 0
    % 문항별 척도 통합 분석
    uniqueQuestions = unique(allScaleAnalysis.Question);
    fprintf('✓ 전체 고유 문항: %d개\n', length(uniqueQuestions));

    % 문항별 척도 일관성 확인
    inconsistentQuestions = {};
    for i = 1:length(uniqueQuestions)
        qName = uniqueQuestions{i};
        qData = allScaleAnalysis(strcmp(allScaleAnalysis.Question, qName), :);

        if height(qData) > 1
            % 여러 시점에서 나타나는 문항
            scales = unique(qData.ScaleType);
            if length(scales) > 1
                inconsistentQuestions{end+1} = qName;
                fprintf('  ⚠️  %s: 시점별 척도 불일치 (%s)\n', qName, strjoin(scales, ', '));
            end
        end
    end

    if isempty(inconsistentQuestions)
        fprintf('✓ 모든 문항의 척도가 시점별로 일관됩니다\n');
    else
        fprintf('⚠️  척도 불일치 문항: %d개\n', length(inconsistentQuestions));
    end

    % 척도 유형별 통계
    fprintf('\n▶ 척도 유형별 분포:\n');
    scaleTypes = unique(allScaleAnalysis.ScaleType);
    for i = 1:length(scaleTypes)
        sType = scaleTypes{i};
        count = sum(strcmp(allScaleAnalysis.ScaleType, sType));
        fprintf('  - %s: %d개 문항\n', sType, count);
    end
end

%% 4. 결과 저장
fprintf('\n[4단계] 분석 결과 저장\n');
fprintf('----------------------------------------\n');

dateStr = datestr(now, 'yyyymmdd_HHMM');
outputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\question_scale_analysis_%s.xlsx', dateStr);

try
    % 전체 분석 결과 저장
    writetable(allScaleAnalysis, outputFileName, 'Sheet', '전체문항척도분석');

    % 문항별 요약 생성
    if ~isempty(uniqueQuestions)
        summaryTable = createQuestionScaleSummary(allScaleAnalysis, uniqueQuestions);
        writetable(summaryTable, outputFileName, 'Sheet', '문항별척도요약');
    end

    % 척도 매핑 테이블 생성 (코드에서 사용할 수 있는 형태)
    if ~isempty(uniqueQuestions)
        mappingTable = createScaleMappingTable(allScaleAnalysis, uniqueQuestions);
        writetable(mappingTable, outputFileName, 'Sheet', '척도매핑테이블');
    end

    fprintf('✓ 분석 결과 저장 완료: %s\n', outputFileName);

    % MAT 파일로도 저장
    matOutputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\question_scale_analysis_%s.mat', dateStr);
    save(matOutputFileName, 'allScaleAnalysis', 'uniqueQuestions', 'periods', 'inconsistentQuestions');

    fprintf('✓ MAT 파일 저장 완료: %s\n', matOutputFileName);

catch ME
    fprintf('✗ 결과 저장 실패: %s\n', ME.message);
end

%% 5. 개선된 표준화 함수 생성
fprintf('\n[5단계] 개선된 표준화 함수 생성\n');
fprintf('----------------------------------------\n');

createImprovedStandardizationFunction(allScaleAnalysis, outputFileName);

fprintf('\n========================================\n');
fprintf('문항별 척도 분석 완료\n');
fprintf('========================================\n');

if ~isempty(inconsistentQuestions)
    fprintf('⚠️  주의: %d개 문항에서 시점별 척도 불일치 발견\n', length(inconsistentQuestions));
    fprintf('   불일치 문항: %s\n', strjoin(inconsistentQuestions, ', '));
end

fprintf('\n📁 생성된 파일:\n');
fprintf('  • 분석 결과: %s\n', outputFileName);
if exist('matOutputFileName', 'var')
    fprintf('  • MAT 파일: %s\n', matOutputFileName);
end

fprintf('\n✅ 다음 단계: corr_item_vs_comp_score_improved.m 파일을 확인하세요\n');

%% ===== 보조 함수들 =====

function idCol = findIDColumn(dataTable)
    idCol = [];
    colNames = dataTable.Properties.VariableNames;

    for col = 1:width(dataTable)
        colName = lower(colNames{col});
        colData = dataTable{:, col};

        if contains(colName, {'id', '사번', 'empno', 'employee'}) && ...
           ((isnumeric(colData) && ~all(isnan(colData))) || ...
            (iscell(colData) && ~all(cellfun(@isempty, colData))) || ...
            (isstring(colData) && ~all(ismissing(colData))))
            idCol = col;
            break;
        end
    end
end

function [questionCols, questionData] = extractQuestionData(selfData, idCol)
    colNames = selfData.Properties.VariableNames;
    questionCols = {};
    questionData = [];

    for col = 1:width(selfData)
        if col == idCol
            continue;
        end

        colName = colNames{col};
        colData = selfData{:, col};

        if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
            questionCols{end+1} = colName;
            questionData = [questionData, colData];
        end
    end
end

function scaleAnalysis = analyzeQuestionScales(questionData, questionNames, periodName)
    scaleAnalysis = table();

    fprintf('  문항별 척도 분석:\n');

    for i = 1:size(questionData, 2)
        columnData = questionData(:, i);
        validData = columnData(~isnan(columnData));

        if isempty(validData)
            continue;
        end

        % 기본 통계
        minVal = min(validData);
        maxVal = max(validData);
        uniqueVals = unique(validData);
        numUnique = length(uniqueVals);
        meanVal = mean(validData);

        % 척도 유형 추정
        if all(mod(validData, 1) == 0) && numUnique <= 10
            scaleType = sprintf('%d~%d점', minVal, maxVal);
        else
            scaleType = '연속형/기타';
        end

        % 데이터 품질 확인
        qualityIssues = {};
        if minVal < 1
            qualityIssues{end+1} = sprintf('최소값이상(%.1f)', minVal);
        end
        if maxVal > 10
            qualityIssues{end+1} = sprintf('최대값이상(%.1f)', maxVal);
        end
        if numUnique == 1
            qualityIssues{end+1} = '단일값';
        end

        qualityStatus = '정상';
        if ~isempty(qualityIssues)
            qualityStatus = strjoin(qualityIssues, ', ');
        end

        % 결과 테이블에 추가
        newRow = table();
        newRow.Period = {periodName};
        newRow.Question = {questionNames{i}};
        newRow.Min = minVal;
        newRow.Max = maxVal;
        newRow.Mean = meanVal;
        newRow.NumUnique = numUnique;
        newRow.ScaleType = {scaleType};
        newRow.QualityStatus = {qualityStatus};
        newRow.SampleSize = length(validData);
        newRow.UniqueValues = {mat2str(uniqueVals(1:min(10, length(uniqueVals))))};

        scaleAnalysis = [scaleAnalysis; newRow];

        % 콘솔 출력
        if strcmp(qualityStatus, '정상')
            statusText = '';
        else
            statusText = ['[' qualityStatus ']'];
        end
        fprintf('    %-6s: %s (N=%d, 평균=%.1f) %s\n', ...
            questionNames{i}, scaleType, length(validData), meanVal, statusText);
    end
end

function summaryTable = createQuestionScaleSummary(allScaleAnalysis, uniqueQuestions)
    summaryTable = table();

    for i = 1:length(uniqueQuestions)
        qName = uniqueQuestions{i};
        qData = allScaleAnalysis(strcmp(allScaleAnalysis.Question, qName), :);

        % 가장 빈번한 척도 유형 선택
        scales = qData.ScaleType;
        [uniqueScales, ~, idx] = unique(scales);
        counts = accumarray(idx, 1);
        [~, maxIdx] = max(counts);
        mostCommonScale = uniqueScales{maxIdx};

        % 범위 정보
        allMins = qData.Min;
        allMaxs = qData.Max;
        avgMin = mean(allMins);
        avgMax = mean(allMaxs);

        % 나타난 시점들
        periods = qData.Period;
        numPeriods = length(periods);

        % 일관성 확인
        isConsistent = length(uniqueScales) == 1;

        newRow = table();
        newRow.Question = {qName};
        newRow.RecommendedScale = {mostCommonScale};
        newRow.AvgMin = avgMin;
        newRow.AvgMax = avgMax;
        newRow.NumPeriods = numPeriods;
        newRow.IsConsistent = isConsistent;
        newRow.AllScales = {strjoin(uniqueScales, ', ')};
        newRow.Periods = {strjoin(periods, ', ')};

        summaryTable = [summaryTable; newRow];
    end
end

function mappingTable = createScaleMappingTable(allScaleAnalysis, uniqueQuestions)
    mappingTable = table();

    fprintf('  척도 매핑 테이블 생성 중...\n');

    for i = 1:length(uniqueQuestions)
        qName = uniqueQuestions{i};
        qData = allScaleAnalysis(strcmp(allScaleAnalysis.Question, qName), :);

        % 가장 적절한 척도 선택 (가장 빈번하고 품질이 좋은 것)
        qualityData = qData(strcmp(qData.QualityStatus, '정상'), :);

        if height(qualityData) > 0
            % 품질 좋은 데이터 우선
            targetData = qualityData;
        else
            % 품질 이슈가 있어도 사용
            targetData = qData;
        end

        % 평균값을 사용하여 권장 척도 결정
        avgMin = round(mean(targetData.Min));
        avgMax = round(mean(targetData.Max));

        % 표준 리커트 척도로 보정
        if avgMax <= 4 && avgMin >= 1
            recommendedMin = 1; recommendedMax = 4;
        elseif avgMax <= 5 && avgMin >= 1
            recommendedMin = 1; recommendedMax = 5;
        elseif avgMax <= 7 && avgMin >= 1
            recommendedMin = 1; recommendedMax = 7;
        elseif avgMax <= 10 && avgMin >= 1
            recommendedMin = 1; recommendedMax = 10;
        else
            recommendedMin = avgMin; recommendedMax = avgMax;
        end

        newRow = table();
        newRow.Question = {qName};
        newRow.RecommendedMin = recommendedMin;
        newRow.RecommendedMax = recommendedMax;
        newRow.DataMin = avgMin;
        newRow.DataMax = avgMax;
        newRow.NumOccurrences = height(qData);

        mappingTable = [mappingTable; newRow];

        fprintf('    %s: %d~%d점 권장 (데이터: %.1f~%.1f)\n', ...
            qName, recommendedMin, recommendedMax, avgMin, avgMax);
    end
end

function createImprovedStandardizationFunction(allScaleAnalysis, outputFileName)
    % 개선된 표준화 함수를 별도 파일로 생성

    try
        % 척도 매핑 정보 추출
        uniqueQuestions = unique(allScaleAnalysis.Question);
        mappingTable = createScaleMappingTable(allScaleAnalysis, uniqueQuestions);

        % 개선된 corr_item_vs_comp_score.m 파일 생성
        improvedFileName = 'D:\project\HR데이터\matlab\문항기반\corr_item_vs_comp_score_improved.m';
        createImprovedCorrelationScript(improvedFileName, mappingTable);

        fprintf('✓ 개선된 상관분석 스크립트 생성: corr_item_vs_comp_score_improved.m\n');

        % 척도 매핑을 Excel에도 저장
        writetable(mappingTable, outputFileName, 'Sheet', '코드용_척도매핑');

    catch ME
        fprintf('✗ 개선된 함수 생성 실패: %s\n', ME.message);
    end
end

function createImprovedCorrelationScript(fileName, mappingTable)
    % 개선된 상관분석 스크립트 파일 생성

    % 척도 매핑을 MATLAB 코드로 변환
    scaleMapping = containers.Map();
    for i = 1:height(mappingTable)
        qName = mappingTable.Question{i};
        minVal = mappingTable.RecommendedMin(i);
        maxVal = mappingTable.RecommendedMax(i);
        scaleMapping(qName) = [minVal, maxVal];
    end

    % 기존 파일 읽기
    originalFileName = 'D:\project\HR데이터\matlab\문항기반\corr_item_vs_comp_score.m';
    if exist(originalFileName, 'file')
        fid = fopen(originalFileName, 'r');
        originalContent = fread(fid, '*char')';
        fclose(fid);

        % 개선된 표준화 함수 추가
        improvedContent = addImprovedStandardizationFunction(originalContent, mappingTable);

        % 새 파일에 저장
        fid = fopen(fileName, 'w');
        fprintf(fid, '%s', improvedContent);
        fclose(fid);
    end
end

function improvedContent = addImprovedStandardizationFunction(originalContent, mappingTable)
    % 기존 내용에 개선된 표준화 함수 추가

    % 표준화 함수 교체 부분 생성
    newFunction = generateImprovedStandardizeFunction(mappingTable);

    % 기존 함수 교체
    pattern = 'function standardizedData = standardizeQuestionScales.*?end';
    improvedContent = regexprep(originalContent, pattern, newFunction, 'dotexceptnewline');

    % 교체가 안된 경우 맨 뒤에 추가
    if strcmp(originalContent, improvedContent)
        improvedContent = [originalContent, newline, newline, newFunction];
    end
end

function functionCode = generateImprovedStandardizeFunction(mappingTable)
    % 개선된 표준화 함수 코드 생성

    functionCode = sprintf(['function standardizedData = standardizeQuestionScales(questionData, questionNames, periodNum)\n'...
        '    %% 개선된 문항별 리커트 척도 표준화 함수 (자동 생성)\n'...
        '    %% 생성일: %s\n'...
        '    \n'...
        '    standardizedData = questionData;\n'...
        '    \n'...
        '    %% 자동 탐지된 척도 매핑 테이블\n'...
        '    scaleMapping = containers.Map();\n'], datestr(now));

    % 척도 매핑 추가
    for i = 1:height(mappingTable)
        qName = mappingTable.Question{i};
        minVal = mappingTable.RecommendedMin(i);
        maxVal = mappingTable.RecommendedMax(i);

        functionCode = [functionCode, sprintf('    scaleMapping(''%s'') = [%d, %d];\n', qName, minVal, maxVal)];
    end

    % 나머지 함수 코드 추가
    functionCode = [functionCode, sprintf(['\n'...
        '    fprintf(''\\n=== 개선된 문항별 척도 표준화 ===\\n'');\n'...
        '    \n'...
        '    for i = 1:size(questionData, 2)\n'...
        '        questionName = questionNames{i};\n'...
        '        columnData = questionData(:, i);\n'...
        '        validData = columnData(~isnan(columnData));\n'...
        '        \n'...
        '        if isempty(validData)\n'...
        '            continue;\n'...
        '        end\n'...
        '        \n'...
        '        %% 척도 정보 가져오기\n'...
        '        if isKey(scaleMapping, questionName)\n'...
        '            scaleInfo = scaleMapping(questionName);\n'...
        '            minScale = scaleInfo(1);\n'...
        '            maxScale = scaleInfo(2);\n'...
        '            fprintf(''%-6s: 사전정의 %%d~%%d점 척도 사용\\n'', questionName, minScale, maxScale);\n'...
        '        else\n'...
        '            %% 자동 탐지\n'...
        '            actualMin = min(validData);\n'...
        '            actualMax = max(validData);\n'...
        '            \n'...
        '            if all(mod(validData, 1) == 0) && length(unique(validData)) <= 10\n'...
        '                %% 정수형 리커트 척도로 추정\n'...
        '                if actualMax <= 4 && actualMin >= 1\n'...
        '                    minScale = 1; maxScale = 4;\n'...
        '                elseif actualMax <= 5 && actualMin >= 1\n'...
        '                    minScale = 1; maxScale = 5;\n'...
        '                elseif actualMax <= 7 && actualMin >= 1\n'...
        '                    minScale = 1; maxScale = 7;\n'...
        '                elseif actualMax <= 10 && actualMin >= 1\n'...
        '                    minScale = 1; maxScale = 10;\n'...
        '                else\n'...
        '                    minScale = actualMin; maxScale = actualMax;\n'...
        '                end\n'...
        '                fprintf(''%-6s: 자동탐지 %%d~%%d점 척도\\n'', questionName, minScale, maxScale);\n'...
        '            else\n'...
        '                %% 연속형 데이터\n'...
        '                minScale = actualMin; maxScale = actualMax;\n'...
        '                fprintf(''%-6s: 연속형 %%.2f~%%.2f 범위\\n'', questionName, minScale, maxScale);\n'...
        '            end\n'...
        '        end\n'...
        '        \n'...
        '        %% Min-Max 표준화 적용\n'...
        '        if maxScale > minScale\n'...
        '            standardizedData(:, i) = (columnData - minScale) / (maxScale - minScale);\n'...
        '        else\n'...
        '            standardizedData(:, i) = 0.5 * ones(size(columnData));\n'...
        '            standardizedData(isnan(columnData), i) = NaN;\n'...
        '        end\n'...
        '        \n'...
        '        %% 결과 검증\n'...
        '        normalizedVals = standardizedData(~isnan(standardizedData(:, i)), i);\n'...
        '        if ~isempty(normalizedVals)\n'...
        '            minNorm = min(normalizedVals);\n'...
        '            maxNorm = max(normalizedVals);\n'...
        '            if minNorm < -0.001 || maxNorm > 1.001\n'...
        '                fprintf(''  ❌ %%s: 표준화 오류 [%%.3f, %%.3f]\\n'', questionName, minNorm, maxNorm);\n'...
        '            end\n'...
        '        end\n'...
        '    end\n'...
        '    \n'...
        '    fprintf(''✓ %%d개 문항 개선된 표준화 완료 ([0,1] 범위)\\n'', size(questionData, 2));\n'...
        'end\n'])];
end