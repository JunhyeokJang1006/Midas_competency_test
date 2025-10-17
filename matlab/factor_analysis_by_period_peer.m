%% 시점별 개별 요인분석 기반 역량진단 성과점수 산출 및 성과기여도 상관분석
% [수평 평가 데이터 활용 버전 - 개선된 집계 및 신뢰도 분석]
% 2023년 하반기 ~ 2025년 상반기 (4개 시점) 데이터 개별 분석
%
% 작성일: 2025년
% 목적: 수평 평가(동료 평가) 데이터를 활용한 각 시점별 독립적인 요인분석 수행

cd('D:\project\HR데이터\matlab')
clear; clc; close all;
diary('D:\project\matlab_runlog\runlog_horizontal_enhanced.txt');


addpath(genpath('D:\project\HR데이터\matlab\refactored'))

%% 1. 초기 설정 및 전역 변수
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

% 결과 저장용 구조체
allData = struct();
periodResults = struct();
consolidatedScores = table();

fprintf('========================================\n');
fprintf('시점별 개별 요인분석 기반 성과점수 산출 시작\n');
fprintf('[수평 평가 데이터 활용 - 개선된 버전]\n');
fprintf('========================================\n\n');

%% 2. 데이터 로드 (수평 평가 데이터 중심)
fprintf('[1단계] 모든 시점 데이터 로드 (수평 평가)\n');
fprintf('----------------------------------------\n');

for p = 1:length(periods)
    fprintf('▶ %s 데이터 로드 중...\n', periods{p});
    fileName = fullfile(dataPath, fileNames{p});

    try
        % 기본 데이터 로드
        allData.(sprintf('period%d', p)).masterIDs = ...
            readtable(fileName, 'Sheet', '기준인원 검토', 'VariableNamingRule', 'preserve');
        
        % 수평 진단 데이터 로드 (동료 평가)
        allData.(sprintf('period%d', p)).horizontalData = ...
            readtable(fileName, 'Sheet', '수평 진단', 'VariableNamingRule', 'preserve');
        
        % 문항 정보 로드
        allData.(sprintf('period%d', p)).questionInfo = ...
            readtable(fileName, 'Sheet', '문항 정보_타인진단', 'VariableNamingRule', 'preserve');

        fprintf('  ✓ 마스터ID: %d명, 수평진단: %d건\n', ...
            height(allData.(sprintf('period%d', p)).masterIDs), ...
            height(allData.(sprintf('period%d', p)).horizontalData));

        % 수평 평가 데이터 구조 분석 (개선된 버전)
        horizontalData = allData.(sprintf('period%d', p)).horizontalData;
        
        % 진단자와 대상자 ID 컬럼 식별 (더 정확한 방법)
        [evaluatorCol, targetCol] = identifyEvaluationColumns(horizontalData);
        
        if ~isempty(evaluatorCol) && ~isempty(targetCol)
            fprintf('  진단자 컬럼: %s\n', horizontalData.Properties.VariableNames{evaluatorCol});
            fprintf('  대상자 컬럼: %s\n', horizontalData.Properties.VariableNames{targetCol});
            
            % 대상자별 평가자 수 계산
            [uniqueTargets, evaluatorCounts, evaluatorStats] = ...
                analyzeEvaluationStructure(horizontalData, evaluatorCol, targetCol);
            
            fprintf('  수평 평가 대상자: %d명\n', length(uniqueTargets));
            fprintf('  평가자 수: 평균 %.1f명 (범위: %d-%d명, 중위수: %.0f명)\n', ...
                evaluatorStats.mean, evaluatorStats.min, evaluatorStats.max, evaluatorStats.median);
            
            % 평가 완성도 분석
            completionRate = analyzeCompletionRate(horizontalData, evaluatorCol, targetCol);
            fprintf('  평가 완성도: %.1f%% (완료된 평가 비율)\n', completionRate);
        else
            fprintf('  [경고] 진단자 또는 대상자 컬럼을 찾을 수 없습니다.\n');
        end

        % 면담용 문항 제외 처리
        fprintf('\n▶ %s: 면담용 문항 제외 처리\n', periods{p});
        
        [cleanedData, excludedInfo] = removeInterviewItems(horizontalData, ...
            allData.(sprintf('period%d', p)).questionInfo, p);
        
        % 원본 데이터 백업 및 정리된 데이터 저장
        allData.(sprintf('period%d', p)).originalHorizontalData = horizontalData;
        allData.(sprintf('period%d', p)).horizontalData = cleanedData;
        allData.(sprintf('period%d', p)).excludedInfo = excludedInfo;
        
        fprintf('  ✓ 면담용/제외 문항 제거 완료\n');
        fprintf('    - 제거 전: %d개 문항\n', excludedInfo.originalCount);
        fprintf('    - 제거 후: %d개 문항\n', excludedInfo.finalCount);

    catch ME
        fprintf('  ✗ %s 데이터 로드 실패: %s\n', periods{p}, ME.message);
        return;
    end
end

fprintf('\n모든 시점 데이터 로드 완료\n');

%% 3. 수평 평가 데이터 집계 및 품질 관리
fprintf('\n[2단계] 수평 평가 데이터 집계 및 품질 관리\n');
fprintf('========================================\n');

% 전체 마스터 ID 리스트 생성
allMasterIDs = [];
for p = 1:length(periods)
    if ~isempty(allData.(sprintf('period%d', p)).masterIDs)
        ids = extractMasterIDs(allData.(sprintf('period%d', p)).masterIDs);
        allMasterIDs = [allMasterIDs; ids];
    end
end
allMasterIDs = unique(allMasterIDs);
fprintf('전체 고유 마스터 ID: %d명\n\n', length(allMasterIDs));

% 결과 저장을 위한 테이블 초기화
consolidatedScores = table();
consolidatedScores.ID = allMasterIDs;

% 각 시점별 분석
for p = 1:length(periods)
    fprintf('========================================\n');
    fprintf('[%s] 분석 시작\n', periods{p});
    fprintf('========================================\n');

    horizontalData = allData.(sprintf('period%d', p)).horizontalData;
    questionInfo = allData.(sprintf('period%d', p)).questionInfo;

    %% 3-1. 수평 평가 데이터 고급 집계
    fprintf('▶ 수평 평가 데이터 고급 집계\n');

    % 대상자 ID 추출
    [evaluatorCol, targetCol] = identifyEvaluationColumns(horizontalData);
    if isempty(targetCol)
        fprintf('  [경고] 대상자 ID 컬럼을 찾을 수 없습니다. 건너뜀.\n\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('NO_TARGET_COLUMN');
     
    end

    % 고급 수평 평가 데이터 집계
    aggregationOptions = struct();
    aggregationOptions.minEvaluators = 2;  % 최소 평가자 수
    aggregationOptions.outlierMethod = 'iqr';  % 이상값 처리 방법
    aggregationOptions.aggregationMethod = 'trimmed_mean';  % 집계 방법
    aggregationOptions.trimPercent = 10;  % 절삭 평균에서 제거할 비율 (상하위 10%)
    
    aggregatedData = aggregateHorizontalEvaluationsAdvanced(horizontalData, ...
        evaluatorCol, targetCol, aggregationOptions);
    
    fprintf('  고급 집계 완료:\n');
    fprintf('    - 대상자 수: %d명\n', length(aggregatedData.targetIDs));
    fprintf('    - 평균 평가자 수: %.1f명\n', mean(aggregatedData.evaluatorCounts));
    fprintf('    - 최소 평가자 기준 충족: %d명\n', sum(aggregatedData.validTargets));
    fprintf('    - 집계 방법: %s\n', aggregationOptions.aggregationMethod);
    
    % 평가자 간 일치도 상세 분석
    reliabilityAnalysis = calculateAdvancedReliability(horizontalData, ...
        evaluatorCol, targetCol, aggregatedData.questionCols);
    
    fprintf('  평가자 간 일치도 분석:\n');
    fprintf('    - ICC(2,k): %.3f\n', reliabilityAnalysis.ICC_2k);
    fprintf('    - ICC(2,1): %.3f\n', reliabilityAnalysis.ICC_21);
    fprintf('    - 평균 상관계수: %.3f\n', reliabilityAnalysis.avgCorrelation);
    fprintf('    - Fleiss'' Kappa: %.3f\n', reliabilityAnalysis.fleissKappa);

    %% 3-2. 집계된 데이터로 응답 매트릭스 생성
    fprintf('\n▶ 분석용 데이터 준비\n');

    % 유효한 대상자만 선택
    validTargetIndices = aggregatedData.validTargets;
    responseData = aggregatedData.aggregatedScores(validTargetIndices, :);
    responseIDs = aggregatedData.targetIDs(validTargetIndices);
    questionCols = aggregatedData.questionCols;
    evaluatorCounts = aggregatedData.evaluatorCounts(validTargetIndices);

    fprintf('  유효 응답자: %d명\n', length(responseIDs));
    fprintf('  평가자 수 분포: %.1f±%.1f (범위: %d-%d)\n', ...
        mean(evaluatorCounts), std(evaluatorCounts), ...
        min(evaluatorCounts), max(evaluatorCounts));

    if size(responseData, 1) < 10
        fprintf('  [경고] 표본 크기가 너무 작습니다. 건너뜀.\n\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('INSUFFICIENT_SAMPLE');
       
    end

    %% 3-3. 수평 평가 특화 데이터 품질 검사
    fprintf('\n▶ 수평 평가 특화 데이터 품질 검사\n');

    qualityResults = performHorizontalDataQualityCheck(responseData, questionCols, ...
        evaluatorCounts, aggregatedData.evaluationVariability);
    
    fprintf('  데이터 품질 검사 결과:\n');
    fprintf('    - 최종 문항 수: %d개\n', qualityResults.finalItemCount);
    fprintf('    - 제거된 문항 수: %d개\n', qualityResults.removedItemCount);
    fprintf('    - 평가자 간 변산성: %.3f\n', qualityResults.avgVariability);
    fprintf('    - 데이터 품질 등급: %s\n', qualityResults.qualityGrade);
    
    % 품질 검사 후 데이터 업데이트
    responseData = qualityResults.cleanedData;
    questionCols = qualityResults.finalQuestionCols;

    %% 3-4. 요인분석 수행 (수평 평가 특화)
    fprintf('\n▶ 수평 평가 특화 요인분석\n');

    % 최적 요인 수 결정 (수평 평가 특성 반영)
    factorAnalysisOptions = struct();
    factorAnalysisOptions.evaluatorWeights = evaluatorCounts;  % 평가자 수를 가중치로 활용
    factorAnalysisOptions.reliabilityThreshold = 0.7;  % 신뢰도 임계값
    
    factorResults = performWeightedFactorAnalysis(responseData, questionCols, ...
        questionInfo, factorAnalysisOptions);
    
    fprintf('  요인분석 결과:\n');
    fprintf('    - 추출된 요인 수: %d개\n', factorResults.numFactors);
    fprintf('    - 성과 요인: %d번째\n', factorResults.performanceFactorIdx);
    fprintf('    - 총 분산 설명률: %.2f%%\n', factorResults.totalVarianceExplained);
    fprintf('    - 분석 방법: %s\n', factorResults.method);

    %% 3-5. 개인별 성과점수 산출 (신뢰도 가중)
    fprintf('\n▶ 신뢰도 가중 성과점수 산출\n');

    % 평가자 수와 일치도를 반영한 가중 점수 계산
    weightedScores = calculateWeightedPerformanceScores(factorResults.factorScores, ...
        factorResults.performanceFactorIdx, evaluatorCounts, reliabilityAnalysis);
    
    fprintf('  성과점수 산출 완료:\n');
    fprintf('    - 평균 점수: %.3f\n', mean(weightedScores.performanceScores, 'omitnan'));
    fprintf('    - 표준편차: %.3f\n', std(weightedScores.performanceScores, 'omitnan'));
    fprintf('    - 신뢰도 가중 적용: %s\n', weightedScores.weightingApplied);

    %% 3-6. 종합 품질 평가
    fprintf('\n▶ 종합 품질 평가\n');

    overallQuality = assessOverallQuality(qualityResults, factorResults, ...
        reliabilityAnalysis, weightedScores);
    
    fprintf('  종합 품질 평가:\n');
    fprintf('    - 데이터 품질: %s\n', overallQuality.dataQuality);
    fprintf('    - 측정 신뢰도: %s\n', overallQuality.measurementReliability);
    fprintf('    - 구인 타당도: %s\n', overallQuality.constructValidity);
    fprintf('    - 전체 등급: %s\n', overallQuality.overallGrade);

    %% 3-7. 결과 저장
    periodResults.(sprintf('period%d', p)) = struct(...
        'loadings', factorResults.loadings, ...
        'factorScores', factorResults.factorScores, ...
        'performanceScores', weightedScores.performanceScores, ...
        'standardizedScores', weightedScores.standardizedScores, ...
        'percentileRanks', weightedScores.percentileRanks, ...
        'numFactors', factorResults.numFactors, ...
        'performanceFactorIdx', factorResults.performanceFactorIdx, ...
        'questionNames', {questionCols}, ...
        'responseIDs', {responseIDs}, ...
        'evaluatorCounts', evaluatorCounts, ...
        'dataQualityFlag', overallQuality.overallGrade, ...
        'scoreReliability', overallQuality.measurementReliability, ...
        'aggregatedData', aggregatedData, ...
        'reliabilityAnalysis', reliabilityAnalysis, ...
        'qualityResults', qualityResults, ...
        'factorResults', factorResults, ...
        'weightedScores', weightedScores, ...
        'overallQuality', overallQuality, ...
        'evaluationType', 'horizontal_advanced');

    %% 3-8. 통합 테이블에 점수 추가
    consolidatedScores = updateConsolidatedScores(consolidatedScores, responseIDs, weightedScores, ...
        evaluatorCounts, overallQuality, p);

    fprintf('  ✓ %s 분석 완료 (등급: %s)\n', periods{p}, overallQuality.overallGrade);
    fprintf('\n');
end


%% ========== 보조 함수들 (Helper Functions) ==========

function [evaluatorCol, targetCol] = identifyEvaluationColumns(dataTable)
    % 진단자와 대상자 컬럼을 더 정확하게 식별
    colNames = dataTable.Properties.VariableNames;
    evaluatorCol = [];
    targetCol = [];
    
    % 첫 번째와 세 번째 ID 컬럼을 찾기 (엑셀 구조 기반)
    idColumns = [];
    for i = 1:length(colNames)
        if contains(lower(colNames{i}), {'id', '사번', 'empno'})
            idColumns(end+1) = i;
        end
    end
    
    if length(idColumns) >= 2
        evaluatorCol = idColumns(1);  % 첫 번째 ID: 진단자
        targetCol = idColumns(2);     % 두 번째 ID: 대상자
    end
end

function [uniqueTargets, evaluatorCounts, stats] = analyzeEvaluationStructure(data, evaluatorCol, targetCol)
    % 평가 구조를 상세히 분석
    targetIDs = data{:, targetCol};
    uniqueTargets = unique(targetIDs);
    uniqueTargets = uniqueTargets(~isnan(uniqueTargets));
    evaluatorCounts = zeros(length(uniqueTargets), 1);
    
    for i = 1:length(uniqueTargets)
        targetID = uniqueTargets(i);
        evaluatorCounts(i) = sum(targetIDs == targetID);
    end
    
    stats = struct();
    stats.mean = mean(evaluatorCounts);
    stats.median = median(evaluatorCounts);
    stats.min = min(evaluatorCounts);
    stats.max = max(evaluatorCounts);
    stats.std = std(evaluatorCounts);
end

function completionRate = analyzeCompletionRate(data, evaluatorCol, targetCol)
    % 평가 완성도 분석
    totalRows = height(data);
    completedRows = sum(~isnan(data{:, evaluatorCol}) & ~isnan(data{:, targetCol}));
    completionRate = (completedRows / totalRows) * 100;
end

function [cleanedData, excludedInfo] = removeInterviewItems(data, questionInfo, periodIndex)
    % 면담용 문항 제거 (개선된 버전)
    excludedInfo = struct();
    
    % Q로 시작하는 문항 컬럼 추출
    colNames = data.Properties.VariableNames;
    questionCols = {};
    questionIndices = [];
    
    for col = 1:width(data)
        colName = colNames{col};
        % Q 또는 q로 시작하고 숫자형 데이터인 컬럼만 선택
        if (startsWith(upper(colName), 'Q') && isnumeric(data{:, col}))
            questionCols{end+1} = colName;
            questionIndices(end+1) = col;
        end
    end
    
    excludedInfo.originalCount = length(questionCols);
    fprintf('    발견된 문항 컬럼: %d개\n', excludedInfo.originalCount);
    
    % 면담용 키워드 - 더 포괄적으로
    interviewKeywords = {
        '점수 산출에 활용되지 않으며', 
        '면담을 위해', 
        '면담용',
        '인터뷰',
        '참고용',
        '점수산출에 활용되지 않',
        '점수에 반영되지 않',
        '평가에 활용되지 않'
    };
    
    % 시점별 특정 제외 패턴
    excludePatterns = {};
    excludeByNumber = [];
    
    if periodIndex == 4  % 25년 상반기
        % Q40-46 제외 (숫자로 직접 지정)
        excludeByNumber = [40, 41, 42, 43, 44, 45, 46];
        fprintf('    25년 상반기: Q40-46 제외 대상\n');
    end
    
    % 모든 시점 공통 제외 문항 (있다면)
    commonExcludeNumbers = []; % 예: [28, 29, 30] 등
    
    % 제외할 문항 찾기
    excludeColumns = {};
    excludeReasons = {};
    
    % 1) 문항 번호로 직접 제외
    for i = 1:length(questionCols)
        qName = questionCols{i};
        
        % 문항 번호 추출 (Q 다음의 숫자)
        qNumMatch = regexp(qName, '^[Qq](\d+)', 'tokens');
        if ~isempty(qNumMatch)
            qNum = str2double(qNumMatch{1}{1});
            
            % 제외 번호에 포함되는지 확인
            if ismember(qNum, excludeByNumber) || ismember(qNum, commonExcludeNumbers)
                excludeColumns{end+1} = qName;
                excludeReasons{end+1} = sprintf('문항 번호 %d (직접 지정)', qNum);
                fprintf('      - %s 제외 (번호 기준)\n', qName);
                continue;
            end
        end
    end
    
    % 2) 문항 정보 시트에서 키워드 기반 제외
    if height(questionInfo) > 0
        fprintf('    문항 정보에서 면담용 키워드 검색 중...\n');
        
        % 문항 정보의 모든 텍스트 컬럼 검사
        for row = 1:height(questionInfo)
            % 전체 행의 텍스트 결합
            rowText = '';
            for col = 1:width(questionInfo)
                cellData = questionInfo{row, col};
                if iscell(cellData)
                    cellData = cellData{1};
                end
                if ischar(cellData) || isstring(cellData)
                    rowText = [rowText, ' ', char(cellData)];
                end
            end
            rowText = lower(rowText);
            
            % 키워드 검사
            for k = 1:length(interviewKeywords)
                if contains(rowText, lower(interviewKeywords{k}))
                    % 이 행의 문항 코드 찾기
                    qCode = questionInfo{row, 1};
                    if iscell(qCode), qCode = qCode{1}; end
                    qCode = char(qCode);
                    
                    % 해당하는 컬럼 찾기
                    for i = 1:length(questionCols)
                        if contains(questionCols{i}, qCode) || strcmp(questionCols{i}, qCode)
                            if ~ismember(questionCols{i}, excludeColumns)
                                excludeColumns{end+1} = questionCols{i};
                                excludeReasons{end+1} = sprintf('키워드: "%s"', interviewKeywords{k});
                                fprintf('      - %s 제외 (키워드: %s)\n', questionCols{i}, interviewKeywords{k});
                            end
                            break;
                        end
                    end
                    break; % 키워드 찾으면 다음 행으로
                end
            end
        end
    end
    
    % 3) 문항 이름 패턴으로 제외 (보조 방법)
    % 예: Q40_1, Q40a 같은 변형도 잡기
    for i = 1:length(questionCols)
        qName = questionCols{i};
        for num = excludeByNumber
            pattern = sprintf('^[Qq]%d[^0-9]?', num); % Q40, Q40_, Q40a 등
            if regexp(qName, pattern)
                if ~ismember(qName, excludeColumns)
                    excludeColumns{end+1} = qName;
                    excludeReasons{end+1} = sprintf('패턴 매칭 Q%d', num);
                    fprintf('      - %s 제외 (패턴)\n', qName);
                end
            end
        end
    end
    
    % 중복 제거
    [excludeColumns, uniqueIdx] = unique(excludeColumns);
    if ~isempty(excludeReasons)
        excludeReasons = excludeReasons(uniqueIdx);
    end
    
    % 데이터에서 제외 컬럼 제거
    cleanedData = data;
    if ~isempty(excludeColumns)
        % 실제로 존재하는 컬럼만 제거
        existingExcludeCols = {};
        for i = 1:length(excludeColumns)
            if ismember(excludeColumns{i}, colNames)
                existingExcludeCols{end+1} = excludeColumns{i};
            end
        end
        
        if ~isempty(existingExcludeCols)
            cleanedData = removevars(cleanedData, existingExcludeCols);
            fprintf('    실제 제거된 컬럼: %d개\n', length(existingExcludeCols));
            
            % 제거 결과 요약
            for i = 1:min(10, length(existingExcludeCols))
                idx = find(strcmp(excludeColumns, existingExcludeCols{i}));
                if ~isempty(idx) && ~isempty(excludeReasons)
                    fprintf('      %s (%s)\n', existingExcludeCols{i}, excludeReasons{idx(1)});
                else
                    fprintf('      %s\n', existingExcludeCols{i});
                end
            end
            if length(existingExcludeCols) > 10
                fprintf('      ... 외 %d개\n', length(existingExcludeCols) - 10);
            end
        end
    else
        fprintf('    제외할 문항이 없습니다.\n');
    end
    
    % 최종 문항 수 계산
    finalQuestionCols = {};
    finalColNames = cleanedData.Properties.VariableNames;
    for col = 1:width(cleanedData)
        colName = finalColNames{col};
        if startsWith(upper(colName), 'Q') && isnumeric(cleanedData{:, col})
            finalQuestionCols{end+1} = colName;
        end
    end
    
    excludedInfo.excludedColumns = excludeColumns;
    excludedInfo.excludeReasons = excludeReasons;
    excludedInfo.finalCount = length(finalQuestionCols);
    excludedInfo.removedCount = excludedInfo.originalCount - excludedInfo.finalCount;
    
    % 디버깅: 남은 문항 확인
    if excludedInfo.finalCount > 0
        fprintf('    남은 문항 예시 (처음 5개): ');
        for i = 1:min(5, length(finalQuestionCols))
            fprintf('%s ', finalQuestionCols{i});
        end
        fprintf('\n');
    end
end

function aggregatedData = aggregateHorizontalEvaluationsAdvanced(data, evaluatorCol, targetCol, options)
    % 고급 수평 평가 데이터 집계
    
    % Q로 시작하는 문항 컬럼 추출
    colNames = data.Properties.VariableNames;
    questionCols = {};
    for col = 1:width(data)
        colName = colNames{col};
        if startsWith(colName, 'Q') && isnumeric(data{:, col})
            questionCols{end+1} = colName;
        end
    end
    
    % 대상자 ID 추출 및 표준화
    rawTargetIDs = data{:, targetCol};
    uniqueTargetIDs = unique(rawTargetIDs);
    uniqueTargetIDs = uniqueTargetIDs(~isnan(uniqueTargetIDs));
    
    % 대상자 ID를 cell array로 표준화
    if isnumeric(uniqueTargetIDs)
        targetIDs = arrayfun(@(x) sprintf('%.0f', x), uniqueTargetIDs, 'UniformOutput', false);
    else
        targetIDs = cellstr(string(uniqueTargetIDs));
    end
    
    numTargets = length(targetIDs);
    numQuestions = length(questionCols);
    
    % 결과 구조체 초기화
    aggregatedData = struct();
    aggregatedData.targetIDs = targetIDs;
    aggregatedData.questionCols = questionCols;
    aggregatedData.aggregatedScores = zeros(numTargets, numQuestions);
    aggregatedData.evaluatorCounts = zeros(numTargets, 1);
    aggregatedData.evaluationVariability = zeros(numTargets, numQuestions);
    aggregatedData.validTargets = false(numTargets, 1);
    
    % 각 대상자별로 집계
    for t = 1:numTargets
        % 원본 데이터에서 해당 대상자 찾기
        if isnumeric(rawTargetIDs)
            currentTargetID = str2double(targetIDs{t});
            targetRows = rawTargetIDs == currentTargetID;
        else
            targetRows = strcmp(string(rawTargetIDs), targetIDs{t});
        end
        
        aggregatedData.evaluatorCounts(t) = sum(targetRows);
        
        if sum(targetRows) >= options.minEvaluators
            aggregatedData.validTargets(t) = true;
            
            % 해당 대상자의 모든 평가 점수
            evaluationScores = table2array(data(targetRows, questionCols));
            
            % 각 문항별 집계
            for q = 1:numQuestions
                questionScores = evaluationScores(:, q);
                validScores = questionScores(~isnan(questionScores));
                
                if ~isempty(validScores)
                    % 집계 방법에 따른 점수 계산
                    switch options.aggregationMethod
                        case 'mean'
                            aggregatedData.aggregatedScores(t, q) = mean(validScores);
                        case 'median'
                            aggregatedData.aggregatedScores(t, q) = median(validScores);
                        case 'trimmed_mean'
                            if length(validScores) >= 4
                                trimmedScores = trimmean(validScores, options.trimPercent);
                                aggregatedData.aggregatedScores(t, q) = trimmedScores;
                            else
                                aggregatedData.aggregatedScores(t, q) = mean(validScores);
                            end
                        otherwise
                            aggregatedData.aggregatedScores(t, q) = mean(validScores);
                    end
                    
                    % 평가자 간 변산성 계산
                    aggregatedData.evaluationVariability(t, q) = std(validScores);
                else
                    aggregatedData.aggregatedScores(t, q) = NaN;
                    aggregatedData.evaluationVariability(t, q) = NaN;
                end
            end
        else
            aggregatedData.aggregatedScores(t, :) = NaN;
            aggregatedData.evaluationVariability(t, :) = NaN;
        end
    end
end

function reliabilityAnalysis = calculateAdvancedReliability(data, evaluatorCol, targetCol, questionCols)
    % 평가자 간 일치도 고급 분석
    
    reliabilityAnalysis = struct();
    
    % ICC(2,k) 및 ICC(2,1) 계산
    [icc2k, icc21] = calculateICC(data, evaluatorCol, targetCol, questionCols);
    reliabilityAnalysis.ICC_2k = icc2k;
    reliabilityAnalysis.ICC_21 = icc21;
    
    % 평가자 간 평균 상관계수
    avgCorr = calculateAverageInterRaterCorrelation(data, evaluatorCol, targetCol, questionCols);
    reliabilityAnalysis.avgCorrelation = avgCorr;
    
    % Fleiss' Kappa (다중 평가자 일치도)
    fleissKappa = calculateFleissKappa(data, evaluatorCol, targetCol, questionCols);
    reliabilityAnalysis.fleissKappa = fleissKappa;
    
    % 평가자별 편향 분석
    raterBias = analyzeRaterBias(data, evaluatorCol, targetCol, questionCols);
    reliabilityAnalysis.raterBias = raterBias;
end

function [icc2k, icc21] = calculateICC(data, evaluatorCol, targetCol, questionCols)
    % Intraclass Correlation Coefficient 계산
    
    targetIDs = unique(data{:, targetCol});
    targetIDs = targetIDs(~isnan(targetIDs));
    
    allScores = [];
    targetLabels = [];
    raterLabels = [];
    
    for t = 1:length(targetIDs)
        targetID = targetIDs(t);
        targetRows = data{:, targetCol} == targetID;
        
        if sum(targetRows) >= 2
            evaluatorIDs = data{targetRows, evaluatorCol};
            targetScores = table2array(data(targetRows, questionCols));
            
            % 총점 계산
            totalScores = sum(targetScores, 2, 'omitnan');
            
            for r = 1:length(totalScores)
                allScores(end+1) = totalScores(r);
                targetLabels(end+1) = t;
                raterLabels(end+1) = r;
            end
        end
    end
    
    if length(allScores) >= 4
        % ICC 계산
        [icc2k, icc21] = calculateICCFromScores(allScores, targetLabels, raterLabels);
    else
        icc2k = NaN;
        icc21 = NaN;
    end
end

function [icc2k, icc21] = calculateICCFromScores(scores, targets, raters)
    % ICC 공식 계산
    
    % 데이터 정리
    uniqueTargets = unique(targets);
    k = length(unique(raters)); % 평가자 수
    n = length(uniqueTargets);  % 대상자 수
    
    % 분산 성분 계산
    grandMean = mean(scores);
    
    % Between-target variance
    targetMeans = zeros(n, 1);
    for i = 1:n
        targetIdx = targets == uniqueTargets(i);
        targetMeans(i) = mean(scores(targetIdx));
    end
    MSB = k * var(targetMeans);
    
    % Within-target variance
    MSW = 0;
    validTargets = 0;
    for i = 1:n
        targetIdx = targets == uniqueTargets(i);
        targetScores = scores(targetIdx);
        if length(targetScores) > 1
            MSW = MSW + var(targetScores);
            validTargets = validTargets + 1;
        end
    end
    if validTargets > 0
        MSW = MSW / validTargets;
    end
    
    % ICC 계산
    if MSB > MSW
        icc2k = (MSB - MSW) / MSB;
        icc21 = (MSB - MSW) / (MSB + (k-1)*MSW);
    else
        icc2k = 0;
        icc21 = 0;
    end
    
    % 범위 제한
    icc2k = max(0, min(1, icc2k));
    icc21 = max(0, min(1, icc21));
end

function avgCorr = calculateAverageInterRaterCorrelation(data, evaluatorCol, targetCol, questionCols)
    % 평가자 간 평균 상관계수 계산
    
    correlations = [];
    
    % 충분한 데이터가 있는 대상자들 찾기
    targetIDs = unique(data{:, targetCol});
    targetIDs = targetIDs(~isnan(targetIDs));
    
    for t = 1:length(targetIDs)
        targetID = targetIDs(t);
        targetRows = data{:, targetCol} == targetID;
        
        if sum(targetRows) >= 3  % 최소 3명의 평가자
            targetScores = table2array(data(targetRows, questionCols));
            
            % 평가자 간 상관행렬 계산
            if size(targetScores, 1) >= 3
                R = corrcoef(targetScores', 'Rows', 'pairwise');
                
                % 대각선 제외한 상관계수들
                upperTri = triu(R, 1);
                validCorr = upperTri(upperTri ~= 0 & ~isnan(upperTri));
                
                correlations = [correlations; validCorr(:)];
            end
        end
    end
    
    if ~isempty(correlations)
        avgCorr = mean(correlations);
    else
        avgCorr = NaN;
    end
end

function fleissKappa = calculateFleissKappa(data, evaluatorCol, targetCol, questionCols)
    % Fleiss' Kappa 계산 (다중 평가자 일치도)
    
    allKappas = [];
    
    for q = 1:length(questionCols)
        questionCol = questionCols{q};
        questionData = data{:, questionCol};
        
        % 범주별 일치도 계산을 위한 데이터 준비
        targetIDs = unique(data{:, targetCol});
        targetIDs = targetIDs(~isnan(targetIDs));
        categories = 1:7;  % 1-7 척도 가정
        
        agreementMatrix = [];
        
        for t = 1:length(targetIDs)
            targetID = targetIDs(t);
            targetRows = data{:, targetCol} == targetID;
            
            if sum(targetRows) >= 3
                targetScores = questionData(targetRows);
                targetScores = targetScores(~isnan(targetScores));
                
                if length(targetScores) >= 3
                    % 각 범주별 평가자 수 계산
                    categoryCount = zeros(1, length(categories));
                    for c = 1:length(categories)
                        categoryCount(c) = sum(targetScores == categories(c));
                    end
                    agreementMatrix = [agreementMatrix; categoryCount];
                end
            end
        end
        
        if size(agreementMatrix, 1) >= 3
            kappa = calculateFleissKappaFromMatrix(agreementMatrix);
            if ~isnan(kappa)
                allKappas(end+1) = kappa;
            end
        end
    end
    
    if ~isempty(allKappas)
        fleissKappa = mean(allKappas);
    else
        fleissKappa = NaN;
    end
end

function kappa = calculateFleissKappaFromMatrix(agreementMatrix)
    % Fleiss' Kappa 공식 계산
    
    [n, k] = size(agreementMatrix);  % n: 대상자 수, k: 범주 수
    m = sum(agreementMatrix(1, :));  % 평가자 수 (일정하다고 가정)
    
    if m == 0
        kappa = NaN;
        return;
    end
    
    % 관찰된 일치도
    Pe_obs = 0;
    for i = 1:n
        for j = 1:k
            Pe_obs = Pe_obs + agreementMatrix(i, j) * (agreementMatrix(i, j) - 1);
        end
    end
    Pe_obs = Pe_obs / (n * m * (m - 1));
    
    % 기대 일치도
    pj = sum(agreementMatrix, 1) / (n * m);  % 각 범주의 marginal probability
    Pe_exp = sum(pj.^2);
    
    % Kappa 계산
    if Pe_exp < 1
        kappa = (Pe_obs - Pe_exp) / (1 - Pe_exp);
    else
        kappa = NaN;
    end
end

function raterBias = analyzeRaterBias(data, evaluatorCol, targetCol, questionCols)
    % 평가자별 편향 분석
    
    evaluatorIDs = unique(data{:, evaluatorCol});
    evaluatorIDs = evaluatorIDs(~isnan(evaluatorIDs));
    
    raterBias = struct();
    raterBias.evaluatorStats = table();
    raterBias.evaluatorStats.EvaluatorID = evaluatorIDs;
    raterBias.evaluatorStats.MeanScore = zeros(length(evaluatorIDs), 1);
    raterBias.evaluatorStats.StdScore = zeros(length(evaluatorIDs), 1);
    raterBias.evaluatorStats.NumEvaluations = zeros(length(evaluatorIDs), 1);
    raterBias.evaluatorStats.Severity = zeros(length(evaluatorIDs), 1);
    
    % 전체 평균 계산
    allScores = [];
    for q = 1:length(questionCols)
        questionData = data{:, questionCols{q}};
        allScores = [allScores; questionData(~isnan(questionData))];
    end
    grandMean = mean(allScores);
    
    % 각 평가자별 통계 계산
    for e = 1:length(evaluatorIDs)
        evaluatorID = evaluatorIDs(e);
        evaluatorRows = data{:, evaluatorCol} == evaluatorID;
        
        evaluatorScores = [];
        for q = 1:length(questionCols)
            questionData = data{evaluatorRows, questionCols{q}};
            evaluatorScores = [evaluatorScores; questionData(~isnan(questionData))];
        end
        
        if ~isempty(evaluatorScores)
            raterBias.evaluatorStats.MeanScore(e) = mean(evaluatorScores);
            raterBias.evaluatorStats.StdScore(e) = std(evaluatorScores);
            raterBias.evaluatorStats.NumEvaluations(e) = length(evaluatorScores);
            raterBias.evaluatorStats.Severity(e) = mean(evaluatorScores) - grandMean;
        end
    end
    
    % 편향 지표 계산
    raterBias.severityRange = range(raterBias.evaluatorStats.Severity);
    raterBias.severityStd = std(raterBias.evaluatorStats.Severity);
    raterBias.variabilityRange = range(raterBias.evaluatorStats.StdScore);
end

function qualityResults = performHorizontalDataQualityCheck(responseData, questionCols, evaluatorCounts, evaluationVariability)
    % 수평 평가 특화 데이터 품질 검사
    
    qualityResults = struct();
    qualityResults.originalData = responseData;
    qualityResults.originalQuestionCols = questionCols;
    
    cleanedData = responseData;
    finalQuestionCols = questionCols;
    removedItems = 0;
    
    % 1. 평가자 수 기반 가중치 적용된 분산 검사
    weightedVariances = zeros(1, size(responseData, 2));
    for q = 1:size(responseData, 2)
        questionScores = responseData(:, q);
        validIdx = ~isnan(questionScores);
        
        if sum(validIdx) > 0
            weights = evaluatorCounts(validIdx) / mean(evaluatorCounts(validIdx));
            weightedVar = sum(weights .* (questionScores(validIdx) - mean(questionScores(validIdx))).^2) / sum(weights);
            weightedVariances(q) = weightedVar;
        end
    end
    
    % 저분산 문항 제거 (가중 분산 기준)
    lowVarianceThreshold = 0.1;
    lowVarianceItems = weightedVariances < lowVarianceThreshold;
    
    if any(lowVarianceItems)
        cleanedData(:, lowVarianceItems) = [];
        finalQuestionCols(lowVarianceItems) = [];
        removedItems = removedItems + sum(lowVarianceItems);
    end
    
    % 2. 평가자 간 변산성 기반 품질 검사
    avgVariability = mean(evaluationVariability, 1, 'omitnan');
    
    % 변산성이 너무 큰 문항 제거
    if ~isempty(avgVariability)
        highVariabilityThreshold = prctile(avgVariability, 95);
        highVariabilityItems = avgVariability > highVariabilityThreshold;
        
        if any(highVariabilityItems) && sum(~highVariabilityItems) >= 5
            cleanedData(:, highVariabilityItems) = [];
            finalQuestionCols(highVariabilityItems) = [];
            removedItems = removedItems + sum(highVariabilityItems);
        end
    end
    
    % 3. 다중공선성 검사
    if size(cleanedData, 2) > 1
        R = corrcoef(cleanedData, 'Rows', 'pairwise');
        [row, col] = find(abs(R) > 0.90 & R ~= 1);
        
        if ~isempty(row)
            validPairs = row < col;
            row = row(validPairs);
            col = col(validPairs);
            
            removeIndices = unique(col);
            if ~isempty(removeIndices) && (size(cleanedData, 2) - length(removeIndices)) >= 3
                cleanedData(:, removeIndices) = [];
                finalQuestionCols(removeIndices) = [];
                removedItems = removedItems + length(removeIndices);
            end
        end
    end
    
    % 4. 품질 등급 판정
    avgVariabilityScore = mean(avgVariability, 'omitnan');
    
    if size(cleanedData, 2) >= 10 && avgVariabilityScore < 1.5
        qualityGrade = 'HIGH';
    elseif size(cleanedData, 2) >= 6 && avgVariabilityScore < 2.0
        qualityGrade = 'MODERATE';
    elseif size(cleanedData, 2) >= 3
        qualityGrade = 'LOW';
    else
        qualityGrade = 'UNUSABLE';
    end
    
    % 결과 저장
    qualityResults.cleanedData = cleanedData;
    qualityResults.finalQuestionCols = finalQuestionCols;
    qualityResults.finalItemCount = length(finalQuestionCols);
    qualityResults.removedItemCount = removedItems;
    qualityResults.avgVariability = avgVariabilityScore;
    qualityResults.qualityGrade = qualityGrade;
    qualityResults.weightedVariances = weightedVariances;
end

function factorResults = performWeightedFactorAnalysis(responseData, questionCols, questionInfo, options)
    % 수평 평가 특화 가중 요인분석
    
    factorResults = struct();
    
    % 평가자 수 가중치 적용 여부 결정
    useWeights = length(options.evaluatorWeights) == size(responseData, 1);
    
    try
        % PCA로 최적 요인 수 결정
        [coeff, score, latent] = pca(responseData);
        
        numFactorsKaiser = sum(latent > 1);
        numFactorsScree = findElbowPoint(latent);
        numFactorsParallel = parallelAnalysis(responseData, 50);
        
        suggestedFactors = [numFactorsKaiser, numFactorsScree, numFactorsParallel];
        optimalNumFactors = median(suggestedFactors);
        optimalNumFactors = max(1, min(optimalNumFactors, min(5, size(responseData, 2)-1)));
        
        % 요인분석 수행
        if useWeights
            % 가중 요인분석
            weights = options.evaluatorWeights / mean(options.evaluatorWeights);
            weightedData = responseData .* sqrt(weights);
            
            [loadings, specificVar, T, stats, factorScores] = ...
                factoran(weightedData, optimalNumFactors, 'rotate', 'promax', 'scores', 'regression');
            
            factorResults.method = 'weighted_factor_analysis';
        else
            [loadings, specificVar, T, stats, factorScores] = ...
                factoran(responseData, optimalNumFactors, 'rotate', 'promax', 'scores', 'regression');
            
            factorResults.method = 'standard_factor_analysis';
        end
        
        % 성과 요인 식별
        performanceFactorIdx = identifyPerformanceFactorAdvanced(loadings, questionCols, questionInfo);
        
        % 분산 설명률 계산
        totalVarianceExplained = 100 * (1 - mean(specificVar));
        
        factorResults.loadings = loadings;
        factorResults.factorScores = factorScores;
        factorResults.numFactors = optimalNumFactors;
        factorResults.performanceFactorIdx = performanceFactorIdx;
        factorResults.totalVarianceExplained = totalVarianceExplained;
        factorResults.specificVar = specificVar;
        factorResults.useWeights = useWeights;
        
    catch ME
        % 요인분석 실패 시 PCA 사용
        numPCs = min(3, size(score, 2));
        factorResults.loadings = coeff(:, 1:numPCs);
        factorResults.factorScores = score(:, 1:numPCs);
        factorResults.numFactors = numPCs;
        factorResults.performanceFactorIdx = 1;
        factorResults.totalVarianceExplained = sum(latent(1:numPCs)) / sum(latent) * 100;
        factorResults.method = 'PCA_fallback';
        factorResults.useWeights = false;
        factorResults.error = ME.message;
    end
end

function weightedScores = calculateWeightedPerformanceScores(factorScores, performanceFactorIdx, evaluatorCounts, reliabilityAnalysis)
    % 신뢰도 가중 성과점수 계산
    
    weightedScores = struct();
    
    if size(factorScores, 2) >= performanceFactorIdx
        rawScores = factorScores(:, performanceFactorIdx);
        
        % 평가자 수 기반 가중치
        evaluatorWeights = sqrt(evaluatorCounts / mean(evaluatorCounts));
        
        % 신뢰도 기반 가중치 (ICC 활용)
        if ~isnan(reliabilityAnalysis.ICC_2k) && reliabilityAnalysis.ICC_2k > 0
            reliabilityWeight = reliabilityAnalysis.ICC_2k;
        else
            reliabilityWeight = 0.5;  % 기본값
        end
        
        % 최종 가중치 계산
        finalWeights = evaluatorWeights * reliabilityWeight;
        
        % 가중 점수 계산
        adjustedScores = rawScores .* finalWeights;
        
        % 표준화
        standardizedScores = zscore(adjustedScores);
        percentileRanks = 100 * tiedrank(adjustedScores) / length(adjustedScores);
        
        weightedScores.performanceScores = adjustedScores;
        weightedScores.standardizedScores = standardizedScores;
        weightedScores.percentileRanks = percentileRanks;
        weightedScores.weights = finalWeights;
        weightedScores.weightingApplied = 'evaluator_count_and_reliability';
        
    else
        weightedScores.performanceScores = NaN(length(evaluatorCounts), 1);
        weightedScores.standardizedScores = NaN(length(evaluatorCounts), 1);
        weightedScores.percentileRanks = NaN(length(evaluatorCounts), 1);
        weightedScores.weights = ones(length(evaluatorCounts), 1);
        weightedScores.weightingApplied = 'none';
    end
end

function overallQuality = assessOverallQuality(qualityResults, factorResults, reliabilityAnalysis, weightedScores)
    % 종합 품질 평가
    
    overallQuality = struct();
    
    % 데이터 품질 평가
    overallQuality.dataQuality = qualityResults.qualityGrade;
    
    % 측정 신뢰도 평가
    if ~isnan(reliabilityAnalysis.ICC_2k)
        if reliabilityAnalysis.ICC_2k > 0.8
            measurementReliability = 'HIGH';
        elseif reliabilityAnalysis.ICC_2k > 0.6
            measurementReliability = 'MODERATE';
        else
            measurementReliability = 'LOW';
        end
    else
        measurementReliability = 'UNKNOWN';
    end
    overallQuality.measurementReliability = measurementReliability;
    
    % 구인 타당도 평가 (분산 설명률 기준)
    if factorResults.totalVarianceExplained > 60
        constructValidity = 'HIGH';
    elseif factorResults.totalVarianceExplained > 40
        constructValidity = 'MODERATE';
    else
        constructValidity = 'LOW';
    end
    overallQuality.constructValidity = constructValidity;
    
    % 전체 등급 결정
    qualityScores = struct();
    qualityScores.data = assignQualityScore(qualityResults.qualityGrade);
    qualityScores.reliability = assignQualityScore(measurementReliability);
    qualityScores.validity = assignQualityScore(constructValidity);
    
    avgQualityScore = mean([qualityScores.data, qualityScores.reliability, qualityScores.validity]);
    
    if avgQualityScore >= 3
        overallGrade = 'HIGH';
    elseif avgQualityScore >= 2
        overallGrade = 'MODERATE';
    else
        overallGrade = 'LOW';
    end
    
    overallQuality.overallGrade = overallGrade;
    overallQuality.qualityScores = qualityScores;
    overallQuality.avgQualityScore = avgQualityScore;
end

function score = assignQualityScore(grade)
    % 품질 등급을 숫자 점수로 변환
    switch grade
        case 'HIGH'
            score = 3;
        case 'MODERATE'
            score = 2;
        case 'LOW'
            score = 1;
        otherwise
            score = 0;
    end
end

function consolidatedScores = updateConsolidatedScores(consolidatedScores, responseIDs, weightedScores, evaluatorCounts, overallQuality, periodIndex)
    % 통합 테이블에 점수 추가
    
    colName = sprintf('Period%d_Score', periodIndex);
    colNameStd = sprintf('Period%d_StdScore', periodIndex);
    colNamePct = sprintf('Period%d_Percentile', periodIndex);
    colNameEvaluators = sprintf('Period%d_Evaluators', periodIndex);
    colNameQuality = sprintf('Period%d_Quality', periodIndex);
    
    % 컬럼 초기화
    consolidatedScores.(colName) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNameStd) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNamePct) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNameEvaluators) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNameQuality) = repmat({overallQuality.overallGrade}, height(consolidatedScores), 1);
    
    % 데이터 할당
    if ~strcmp(overallQuality.overallGrade, 'UNUSABLE')
        for i = 1:length(responseIDs)
            % ID를 문자열로 변환하여 비교
            if iscell(consolidatedScores.ID)
                targetID = responseIDs{i};
                idx = strcmp(consolidatedScores.ID, targetID);
            else
                % 숫자형 ID인 경우
                if iscell(responseIDs)
                    targetID = str2double(responseIDs{i});
                else
                    targetID = responseIDs(i);
                end
                idx = consolidatedScores.ID == targetID;
            end
            
            if any(idx)
                consolidatedScores.(colName)(idx) = weightedScores.performanceScores(i);
                consolidatedScores.(colNameStd)(idx) = weightedScores.standardizedScores(i);
                consolidatedScores.(colNamePct)(idx) = weightedScores.percentileRanks(i);
                consolidatedScores.(colNameEvaluators)(idx) = evaluatorCounts(i);
            end
        end
    end
end

function emptyResult = createEmptyPeriodResult(reason)
    % 빈 결과 구조체 생성
    emptyResult = struct(...
        'loadings', [], ...
        'factorScores', [], ...
        'performanceScores', [], ...
        'standardizedScores', [], ...
        'percentileRanks', [], ...
        'numFactors', 0, ...
        'performanceFactorIdx', 1, ...
        'questionNames', {{}}, ...
        'responseIDs', {{}}, ...
        'evaluatorCounts', [], ...
        'dataQualityFlag', 'FAILED', ...
        'scoreReliability', 'UNUSABLE', ...
        'evaluationType', 'horizontal_advanced', ...
        'failureReason', reason);
end

function masterIDs = extractMasterIDs(masterTable)
    % 마스터 테이블에서 ID 추출
    masterIDs = {};
    if height(masterTable) == 0, return; end
    
    for col = 1:width(masterTable)
        colName = masterTable.Properties.VariableNames{col};
        if contains(lower(colName), {'id', '사번', 'empno'})
            ids = masterTable{:, col};
            if isnumeric(ids)
                masterIDs = arrayfun(@(x) sprintf('%.0f', x), ids, 'UniformOutput', false);
            else
                masterIDs = cellstr(ids);
            end
            masterIDs = masterIDs(~cellfun(@isempty, masterIDs));
            break;
        end
    end
end

function elbowPoint = findElbowPoint(eigenValues)
    % Scree plot에서 elbow point 찾기
    if length(eigenValues) < 3
        elbowPoint = 1;
        return;
    end
    
    diffs = diff(eigenValues);
    secondDiffs = diff(diffs);
    [~, elbowPoint] = max(abs(secondDiffs));
    elbowPoint = min(elbowPoint + 1, length(eigenValues));
    elbowPoint = max(1, min(elbowPoint, length(eigenValues)));
end

function numFactors = parallelAnalysis(data, numIterations)
    % Parallel Analysis를 통한 요인 수 결정
    [n, p] = size(data);
    realEigenValues = eig(cov(data));
    realEigenValues = sort(realEigenValues, 'descend');
    
    randomEigenValues = zeros(numIterations, p);
    for iter = 1:numIterations
        randomData = randn(n, p);
        randomEigenValues(iter, :) = sort(eig(cov(randomData)), 'descend');
    end
    
    randomEigenThreshold = prctile(randomEigenValues, 95);
    numFactors = sum(realEigenValues > randomEigenThreshold');
    numFactors = max(1, numFactors);
end

function performanceIdx = identifyPerformanceFactorAdvanced(loadings, questionNames, questionInfo)
    % 성과 요인 식별 (개선된 버전)
    performanceKeywords = {'성과', '목표', '달성', '결과', '효과', '기여', '창출', '개선', '수행'};
    numFactors = size(loadings, 2);
    performanceScores = zeros(numFactors, 1);
    
    for f = 1:numFactors
        highLoadingItems = find(abs(loadings(:, f)) > 0.3);
        
        for itemIdx = 1:length(highLoadingItems)
            item = highLoadingItems(itemIdx);
            questionName = questionNames{item};
            
            try
                if height(questionInfo) > 0
                    for row = 1:height(questionInfo)
                        qCode = questionInfo{row, 1};
                        if iscell(qCode), qCode = qCode{1}; end
                        qCode = char(qCode);
                        
                        if contains(questionName, qCode)
                            questionText = questionInfo{row, 2};
                            if iscell(questionText), questionText = questionText{1}; end
                            questionText = char(questionText);
                            
                            for k = 1:length(performanceKeywords)
                                if contains(lower(questionText), performanceKeywords{k})
                                    performanceScores(f) = performanceScores(f) + abs(loadings(item, f));
                                end
                            end
                            break;
                        end
                    end
                end
            catch
           
            end
        end
        
        % 높은 적재값을 가진 문항 수도 고려
        performanceScores(f) = performanceScores(f) + 0.1 * sum(abs(loadings(:, f)) > 0.5);
    end
    
    [~, performanceIdx] = max(performanceScores);
    if all(performanceScores == 0)
        performanceIdx = 1;  % 기본값: 첫 번째 요인
    end
end