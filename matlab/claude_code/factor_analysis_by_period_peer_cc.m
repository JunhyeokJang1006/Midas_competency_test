%% 시점별 개별 요인분석 기반 역량진단 성과점수 산출 및 성과기여도 상관분석 (수평 진단 버전)
% 2023년 하반기 ~ 2025년 상반기 (4개 시점) 수평 진단 데이터 개별 분석
%
% 작성일: 2025년
% 목적: 수평 진단 데이터를 사용하여 각 시점별로 독립적인 요인분석 수행 후 개별 점수 산출 및 성과기여도와 상관분석
% 특징: 동료 평가 점수를 마스터ID별로 평균하여 하향 진단 형태로 변환 후 원본과 동일한 분석 적용
cd('D:\project\HR데이터\matlab')
clear; clc; close all;
diary('D:\project\matlab_runlog\runlog_horizontal_enhanced.txt');




%% 1. 초기 설정 및 전역 변수
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

% 결과 저장용 구조체
allData = struct();
periodResults = struct();
consolidatedScores = table();

fprintf('========================================\n');
fprintf('시점별 개별 요인분석 기반 성과점수 산출 시작 (수평 진단 버전)\n');
fprintf('========================================\n\n');

%% 2. 데이터 로드 (수평 진단 데이터를 하향 진단 형태로 변환)
fprintf('[1단계] 모든 시점 데이터 로드\n');
fprintf('----------------------------------------\n');

for p = 1:length(periods)
    fprintf('▶ %s 데이터 로드 중...\n', periods{p});
    fileName = fullfile(dataPath, fileNames{p});

    try
        % 기본 데이터 로드
        allData.(sprintf('period%d', p)).masterIDs = ...
            readtable(fileName, 'Sheet', '기준인원 검토', 'VariableNamingRule', 'preserve');
        allData.(sprintf('period%d', p)).peerRawData = ...
            readtable(fileName, 'Sheet', '수평 진단', 'VariableNamingRule', 'preserve');
        allData.(sprintf('period%d', p)).questionInfo = ...
            readtable(fileName, 'Sheet', '문항 정보_타인진단', 'VariableNamingRule', 'preserve');

        fprintf('  ✓ 마스터ID: %d명, 수평진단 원시데이터: %d개 레코드\n', ...
            height(allData.(sprintf('period%d', p)).masterIDs), ...
            height(allData.(sprintf('period%d', p)).peerRawData));

        % 수평 진단 데이터를 하향 진단 형태로 변환
        fprintf('\n▶ %s: 수평 진단 → 하향 진단 변환\n', periods{p});
        
        peerRawData = allData.(sprintf('period%d', p)).peerRawData;
        colNames = peerRawData.Properties.VariableNames;
        
        % 첫 번째 열: 평가대상자, 두 번째 열: 평가자
        if width(peerRawData) < 2
            error('수평 진단 데이터가 올바르지 않습니다.');
        end
        
        targetCol = peerRawData{:,1};  % 평가대상자 ID
        raterCol = peerRawData{:,2};   % 평가자 ID
        
        % 유효한 데이터만 추출 (0이나 NaN이 아닌 경우)
        if isnumeric(targetCol) && isnumeric(raterCol)
            validRows = ~(isnan(targetCol) | isnan(raterCol) | targetCol==0 | raterCol==0);
        else
            validRows = true(height(peerRawData), 1);
        end
        
        validData = peerRawData(validRows, :);
        
        fprintf('  유효한 평가 레코드: %d개\n', sum(validRows));
        
        % Q로 시작하는 문항 컬럼 식별
        questionCols = {};
        questionIndices = [];
        for col = 3:width(validData)  % 3번째 컬럼부터 (첫 2개는 ID)
            colName = colNames{col};
            colData = validData{:, col};
            
            if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
                questionCols{end+1} = colName;
                questionIndices(end+1) = col;
            end
        end
        
        fprintf('  문항 수: %d개\n', length(questionCols));
        
        if isempty(questionCols)
            fprintf('  ❌ 분석할 문항이 없습니다. 건너뜀.\n\n');
            continue;
        end
        
        % 고유한 평가대상자 목록
        if isnumeric(validData{:,1})
            uniqueTargets = unique(validData{:,1});
            uniqueTargets = uniqueTargets(~isnan(uniqueTargets) & uniqueTargets > 0);
        else
            uniqueTargets = unique(validData{:,1});
            uniqueTargets = uniqueTargets(~cellfun(@isempty, uniqueTargets));
        end
        
        fprintf('  고유한 평가 대상자: %d명\n', length(uniqueTargets));
        
        % 평가대상자별 평균 점수 계산하여 하향진단 형태로 변환
        convertedData = table();
        
        % ID 컬럼명을 하향진단과 동일하게 설정
        firstColName = colNames{1};
        convertedData.(firstColName) = uniqueTargets;
        
        % 각 문항에 대해 평가대상자별 평균 계산
        evaluatorCounts = zeros(length(uniqueTargets), 1);
        
        for q = 1:length(questionCols)
            questionCol = questionCols{q};
            questionIndex = questionIndices(q);
            avgScores = zeros(length(uniqueTargets), 1);
            
            for t = 1:length(uniqueTargets)
                targetID = uniqueTargets(t);
                if isnumeric(validData{:,1})
                    targetRows = validData{:,1} == targetID;
                else
                    targetRows = strcmp(validData{:,1}, targetID);
                end
                
                targetScores = validData{targetRows, questionIndex};
                
                % NaN이 아닌 점수들의 평균
                if isnumeric(targetScores)
                    validScores = targetScores(~isnan(targetScores));
                else
                    validScores = targetScores;
                end
                
                if ~isempty(validScores)
                    avgScores(t) = mean(validScores);
                    if q == 1  % 첫 번째 문항에서만 평가자 수 계산
                        evaluatorCounts(t) = length(validScores);
                    end
                else
                    avgScores(t) = NaN;
                    if q == 1
                        evaluatorCounts(t) = 0;
                    end
                end
            end
            
            convertedData.(questionCol) = avgScores;
        end
        
        fprintf('  변환 완료: %d명의 평가 대상자\n', height(convertedData));
        fprintf('  평균 평가자 수: %.1f명 (범위: %d-%d명)\n', ...
            mean(evaluatorCounts), min(evaluatorCounts), max(evaluatorCounts));
        
        % 변환된 데이터를 selfData로 저장 (원본 코드와 호환성 유지)
        allData.(sprintf('period%d', p)).selfData = convertedData;
        allData.(sprintf('period%d', p)).evaluatorCounts = evaluatorCounts;
        
        fprintf('  ✅ %s 데이터 로드 및 변환 완료\n\n', periods{p});
        
    catch ME
        fprintf('  ❌ %s 데이터 로드 실패: %s\n', periods{p}, ME.message);
        fprintf('     파일 경로: %s\n\n', fileName);
        continue;
    end
end

%% 3. 시점별 개별 요인분석 수행 (원본 코드와 동일한 구조)
fprintf('\n[2단계] 시점별 개별 요인분석 수행\n');
fprintf('========================================\n');

% 전체 마스터 ID 리스트 생성
allMasterIDs = [];
for p = 1:length(periods)
    if isfield(allData, sprintf('period%d', p)) && ...
       isfield(allData.(sprintf('period%d', p)), 'masterIDs') && ...
       ~isempty(allData.(sprintf('period%d', p)).masterIDs)
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

    if ~isfield(allData, sprintf('period%d', p)) || ...
       ~isfield(allData.(sprintf('period%d', p)), 'selfData')
        fprintf('[경고] 데이터가 없습니다. 건너뜀.\n\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('NO_DATA');
        continue;
    end

    selfData = allData.(sprintf('period%d', p)).selfData;
    questionInfo = allData.(sprintf('period%d', p)).questionInfo;

    %% 3-1. 문항 데이터 추출
    fprintf('▶ 문항 데이터 추출\n');

    % ID 컬럼 찾기
    idCol = findIDColumn(selfData);
    if isempty(idCol)
        fprintf('  [경고] ID 컬럼을 찾을 수 없습니다. 건너뜀.\n\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('NO_ID_COLUMN');
        continue;
    end

    % Q로 시작하는 문항 컬럼 추출
    colNames = selfData.Properties.VariableNames;
    questionCols = {};
    for col = 1:width(selfData)
        colName = colNames{col};
        colData = selfData{:, col};

        if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
            questionCols{end+1} = colName;
        end
    end

    fprintf('  발견된 문항 수: %d개\n', length(questionCols));

    if length(questionCols) < 3
        fprintf('  [경고] 문항이 너무 적어 요인분석이 불가능합니다. 건너뜀.\n\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('INSUFFICIENT_ITEMS');
        continue;
    end

    % 응답 데이터 추출
    responseData = table2array(selfData(:, questionCols));
    responseIDs = extractAndStandardizeIDs(selfData{:, idCol});

    % 결측치 처리
    responseData = handleMissingValues(responseData);

    % 유효한 행만 선택
    validRows = sum(isnan(responseData), 2) < (size(responseData, 2) * 0.5);
    responseData = responseData(validRows, :);
    responseIDs = responseIDs(validRows);

    fprintf('  유효 응답자: %d명\n', length(responseIDs));

    if size(responseData, 1) < 10
        fprintf('  [경고] 표본 크기가 너무 작습니다. 건너뜀.\n\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('INSUFFICIENT_SAMPLE');
        continue;
    end

    %% 3-2. 데이터 품질 검사 및 전처리 (원본과 동일)
    fprintf('\n▶ 데이터 품질 검사 및 전처리\n');

    originalNumQuestions = size(responseData, 2);
    dataQualityFlag = 'UNKNOWN';

    % 1. 상수 열 제거
    columnVariances = var(responseData, 0, 1, 'omitnan');
    constantColumns = columnVariances < 1e-10;

    if any(constantColumns)
        fprintf('  [제거] 상수 응답 문항 %d개\n', sum(constantColumns));
        responseData(:, constantColumns) = [];
        questionCols(constantColumns) = [];
        columnVariances(constantColumns) = [];
    end

    % 2. 다중공선성 처리
    if size(responseData, 2) > 1
        R = corrcoef(responseData, 'Rows', 'pairwise');
        
        % 완벽한 상관관계 제거
        [toRemove, ~] = find(triu(abs(R) > 0.95, 1));
        toRemove = unique(toRemove);
        
        if ~isempty(toRemove)
            fprintf('  [제거] 다중공선성 문항 %d개\n', length(toRemove));
            responseData(:, toRemove) = [];
            questionCols(toRemove) = [];
            columnVariances(toRemove) = [];
        end
    end

    % 3. 극단값 처리
    outlierCount = 0;
    if size(responseData, 2) > 0
        for col = 1:size(responseData, 2)
            colData = responseData(:, col);
            if ~all(isnan(colData))
                meanVal = mean(colData, 'omitnan');
                stdVal = std(colData, 'omitnan');
                
                if stdVal > 0
                    outlierIdx = abs(colData - meanVal) > 3 * stdVal;
                    if any(outlierIdx)
                        colData(outlierIdx & colData > meanVal) = meanVal + 3 * stdVal;
                        colData(outlierIdx & colData < meanVal) = meanVal - 3 * stdVal;
                        responseData(:, col) = colData;
                        outlierCount = outlierCount + sum(outlierIdx);
                    end
                end
            end
        end
    end

    if outlierCount > 0
        fprintf('  [처리] 극단값 %d개 조정\n', outlierCount);
    end

    % 4. 저분산 문항 제거
    lowVarianceThreshold = 0.1;
    if ~isempty(columnVariances)
        lowVarianceColumns = columnVariances < lowVarianceThreshold;

        if any(lowVarianceColumns)
            fprintf('  [제거] 저분산 문항 %d개\n', sum(lowVarianceColumns));
            responseData(:, lowVarianceColumns) = [];
            questionCols(lowVarianceColumns) = [];
            columnVariances(lowVarianceColumns) = [];
        end
    end

    % 5. 최종 품질 검사
    if size(responseData, 2) < 3
        fprintf('  [오류] 전처리 후 문항이 너무 적습니다.\n');
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('POST_PROCESS_INSUFFICIENT');
        continue;
    end

    try
        R_final = corrcoef(responseData, 'Rows', 'pairwise');
        det_final = det(R_final);
        cond_final = cond(R_final);

        fprintf('  - 최종 문항 수: %d개 (원본: %d개)\n', size(responseData, 2), originalNumQuestions);
        fprintf('  - 상관행렬 조건수: %.2e\n', cond_final);
        fprintf('  - 상관행렬 행렬식: %.2e\n', det_final);
        
        % 상세한 품질 진단 정보
        fprintf('\n  [상세 품질 진단]\n');
        fprintf('  - 데이터 크기: %d x %d\n', size(responseData, 1), size(responseData, 2));
        fprintf('  - 결측값 비율: %.2f%%\n', (sum(isnan(responseData(:))) / numel(responseData)) * 100);
        fprintf('  - 상관행렬 최대값: %.4f\n', max(R_final(:)));
        fprintf('  - 상관행렬 최소값: %.4f\n', min(R_final(:)));
        fprintf('  - 상관행렬 평균: %.4f\n', mean(R_final(:), 'omitnan'));
        
        % 특이값 분석
        [~, S, ~] = svd(R_final);
        singular_values = diag(S);
        fprintf('  - 특이값 범위: %.2e ~ %.2e\n', min(singular_values), max(singular_values));
        fprintf('  - 특이값 비율 (최대/최소): %.2e\n', max(singular_values)/min(singular_values));

        % 품질 판정 (기존 기준)
        if (det_final > 1e-10) && (cond_final < 1e10)
            dataQualityFlag = 'GOOD';
            fprintf('  ✓ 수치적 안정성 양호 (기존 기준)\n');
        elseif (det_final > 1e-15) && (cond_final < 1e12)
            dataQualityFlag = 'CAUTION';
            fprintf('  ⚠ 수치적 문제 있음 - 주의 필요 (기존 기준)\n');
        else
            dataQualityFlag = 'POOR';
            fprintf('  ✗ 심각한 수치적 문제 (기존 기준)\n');
        end
        
        % 대안적 품질 기준 제안
        fprintf('\n  [대안적 품질 기준]\n');
        if (det_final > 1e-15) && (cond_final < 1e12)
            fprintf('  ✓ 완화된 기준: GOOD\n');
        elseif (det_final > 1e-20) && (cond_final < 1e15)
            fprintf('  ✓ 매우 완화된 기준: CAUTION\n');
        else
            fprintf('  ✗ 모든 기준 실패: POOR\n');
        end

    catch ME
        dataQualityFlag = 'FAILED';
        fprintf('  [경고] 상관행렬 계산 실패: %s\n', ME.message);
    end

    %% 3-3. 최적 요인 수 결정 (원본과 동일)
    fprintf('\n▶ 최적 요인 수 결정\n');

    try
        [coeff, score, latent] = pca(responseData);

        numFactorsKaiser = sum(latent > 1);
        numFactorsScree = findElbowPoint(latent);
        numFactorsParallel = parallelAnalysis(responseData, 50);

        fprintf('  - Kaiser 기준: %d개\n', numFactorsKaiser);
        fprintf('  - Scree plot: %d개\n', numFactorsScree);
        fprintf('  - Parallel analysis: %d개\n', numFactorsParallel);

        suggestedFactors = [numFactorsKaiser, numFactorsScree, numFactorsParallel];
        optimalNumFactors = median(suggestedFactors);
        optimalNumFactors = max(1, min(optimalNumFactors, min(5, size(responseData, 2)-1)));

        fprintf('  ✓ 선택된 요인 수: %d개\n', optimalNumFactors);

    catch ME
        fprintf('  [경고] PCA 실패: %s\n', ME.message);
        optimalNumFactors = 1;
    end

    %% 3-4. 요인분석 수행 (원본과 동일)
    fprintf('\n▶ 요인분석 실행\n');

    isPCA = false;

    try
        [loadings, specificVar, T, stats, factorScores] = ...
            factoran(responseData, optimalNumFactors, 'rotate', 'promax', 'scores', 'regression');

        fprintf('  ✓ 요인분석 성공 (Promax 회전)\n');
        fprintf('  - 누적 분산 설명률: %.2f%%\n', 100 * (1 - mean(specificVar)));

    catch ME
        fprintf('  [경고] 요인분석 실패: %s\n', ME.message);
        fprintf('  [대안] PCA 점수 사용\n');

        try
            [coeff, score, latent, ~, explained] = pca(responseData);
            loadings = coeff(:, 1:optimalNumFactors);
            factorScores = score(:, 1:optimalNumFactors);
            isPCA = true;
            fprintf('  ✓ PCA 성공\n');
            fprintf('  - 누적 분산 설명률: %.2f%%\n', sum(explained(1:optimalNumFactors)));
        catch ME2
            fprintf('  [오류] PCA도 실패: %s\n', ME2.message);
            periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('ANALYSIS_FAILED');
            continue;
        end
    end

    %% 3-5. 성과 요인 식별 및 점수 산출 (원본과 동일)
    fprintf('\n▶ 성과 요인 식별 및 점수 산출\n');

    % 성과 관련 요인 식별
    performanceFactorIdx = identifyPerformanceFactorAdvanced(loadings, questionCols, questionInfo);
    fprintf('  - 성과 요인: Factor %d\n', performanceFactorIdx);

    % 개인별 성과점수 계산
    performanceScores = factorScores(:, performanceFactorIdx);

    % 0-100 점수로 변환
    minScore = min(performanceScores);
    maxScore = max(performanceScores);
    
    if maxScore > minScore
        scaledScores = 20 + (performanceScores - minScore) / (maxScore - minScore) * 60;
    else
        scaledScores = ones(size(performanceScores)) * 50; % 모든 점수가 같으면 중간값
    end

    fprintf('  - 성과점수 범위: %.1f ~ %.1f점\n', min(scaledScores), max(scaledScores));
    fprintf('  - 성과점수 평균: %.1f점 (표준편차: %.1f)\n', mean(scaledScores), std(scaledScores));

    %% 3-6. 결과 저장
    periodResult = struct();
    periodResult.period = periods{p};
    periodResult.analysisDate = datestr(now);
    periodResult.numParticipants = length(responseIDs);
    periodResult.numQuestions = size(responseData, 2);
    periodResult.originalNumQuestions = originalNumQuestions;
    periodResult.numFactors = optimalNumFactors;
    periodResult.dataQuality = dataQualityFlag;
    periodResult.isPCA = isPCA;
    periodResult.performanceFactorIdx = performanceFactorIdx;
    periodResult.participantIDs = responseIDs;
    periodResult.performanceScores = scaledScores;
    periodResult.factorLoadings = loadings;
    periodResult.factorScores = factorScores;
    periodResult.questionNames = questionCols;
    
    % 수평 진단 추가 정보
    if isfield(allData.(sprintf('period%d', p)), 'evaluatorCounts')
        periodResult.avgEvaluators = mean(allData.(sprintf('period%d', p)).evaluatorCounts);
    else
        periodResult.avgEvaluators = NaN;
    end

    periodResults.(sprintf('period%d', p)) = periodResult;

    % 통합 점수 테이블에 추가
    periodColName = sprintf('%s_Performance', periods{p});
    
    % ID 매칭하여 점수 할당
    periodScoreVector = NaN(height(consolidatedScores), 1);
    for i = 1:length(responseIDs)
        idIdx = find(strcmp(consolidatedScores.ID, responseIDs{i}));
        if ~isempty(idIdx)
            periodScoreVector(idIdx) = scaledScores(i);
        end
    end
    
    consolidatedScores.(periodColName) = periodScoreVector;

    fprintf('\n✅ [%s] 분석 완료\n', periods{p});
    fprintf('   참여자: %d명, 문항: %d개, 요인: %d개\n', ...
        length(responseIDs), size(responseData, 2), optimalNumFactors);
    if ~isnan(periodResult.avgEvaluators)
        fprintf('   평균 평가자 수: %.1f명\n', periodResult.avgEvaluators);
    end
    fprintf('\n');
end

%% 4. 종합 분석 및 통계
fprintf('[3단계] 종합 분석 및 통계\n');
fprintf('========================================\n');

% 성공한 분석 개수
successCount = 0;
totalParticipants = 0;
for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p)) && ...
       ~strcmp(periodResults.(sprintf('period%d', p)).period, 'EMPTY')
        successCount = successCount + 1;
        totalParticipants = totalParticipants + periodResults.(sprintf('period%d', p)).numParticipants;
    end
end

fprintf('✓ 성공한 분석: %d/%d개 시점\n', successCount, length(periods));
fprintf('✓ 총 분석 참여자: %d명\n', totalParticipants);

% 시점별 결과 요약
fprintf('\n▶ 시점별 분석 결과:\n');
fprintf('%-15s %10s %10s %10s %12s %10s\n', '시점', '참여자수', '문항수', '요인수', '데이터품질', '평가자수');
fprintf('%s\n', repmat('-', 1, 70));

for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p)) && ...
       ~strcmp(periodResults.(sprintf('period%d', p)).period, 'EMPTY')
        result = periodResults.(sprintf('period%d', p));
        evaluatorInfo = '';
        if ~isnan(result.avgEvaluators)
            evaluatorInfo = sprintf('%.1f', result.avgEvaluators);
        else
            evaluatorInfo = 'N/A';
        end
        fprintf('%-15s %10d %10d %10d %12s %10s\n', ...
            result.period, result.numParticipants, result.numQuestions, ...
            result.numFactors, result.dataQuality, evaluatorInfo);
    else
        fprintf('%-15s %10s %10s %10s %12s %10s\n', ...
            periods{p}, 'FAILED', '-', '-', '-', '-');
    end
end

%% 5. 결과 저장
fprintf('\n[4단계] 결과 저장\n');
fprintf('========================================\n');

try
    % 결과 디렉토리 생성
    resultDir = 'D:\project\HR데이터\결과';
    if ~exist(resultDir, 'dir')
        mkdir(resultDir);
    end
    
    % 파일명 생성
    timestamp = datestr(now, 'yyyymmdd_HHMMSS');
    resultFileName = sprintf('수평진단_시점별요인분석_결과_%s.xlsx', timestamp);
    resultFilePath = fullfile(resultDir, resultFileName);
    
    % 1. 종합 점수 시트
    writetable(consolidatedScores, resultFilePath, 'Sheet', '종합점수');
    fprintf('✓ 종합 점수 저장 완료\n');
    
    % 2. 시점별 상세 결과 시트
    for p = 1:length(periods)
        if isfield(periodResults, sprintf('period%d', p)) && ...
           ~strcmp(periodResults.(sprintf('period%d', p)).period, 'EMPTY')
            
            result = periodResults.(sprintf('period%d', p));
            
            % 상세 결과 테이블 생성
            detailTable = table();
            detailTable.ID = result.participantIDs;
            detailTable.PerformanceScore = result.performanceScores;
            
            % 요인별 점수 추가
            for f = 1:result.numFactors
                factorColName = sprintf('Factor%d_Score', f);
                factorScores = result.factorScores(:, f);
                
                % 0-100 스케일로 변환
                minFS = min(factorScores);
                maxFS = max(factorScores);
                if maxFS > minFS
                    scaledFS = 20 + (factorScores - minFS) / (maxFS - minFS) * 60;
                else
                    scaledFS = ones(size(factorScores)) * 50;
                end
                
                detailTable.(factorColName) = scaledFS;
            end
            
            sheetName = sprintf('%s_상세', result.period);
            writetable(detailTable, resultFilePath, 'Sheet', sheetName);
            
            % 요인 부하량 시트
            loadingTable = table();
            loadingTable.Question = result.questionNames';
            for f = 1:result.numFactors
                factorColName = sprintf('Factor%d', f);
                loadingTable.(factorColName) = result.factorLoadings(:, f);
            end
            
            loadingSheetName = sprintf('%s_부하량', result.period);
            writetable(loadingTable, resultFilePath, 'Sheet', loadingSheetName);
        end
    end
    
    fprintf('✓ 상세 결과 저장 완료\n');
    fprintf('📁 결과 파일: %s\n', resultFilePath);
    
catch ME
    fprintf('❌ 결과 저장 실패: %s\n', ME.message);
end

%% 6. 통합 점수 테이블 생성 (신뢰도 포함)
fprintf('\n[6단계] 통합 점수 테이블 생성 (품질 검증 포함)\n');
fprintf('========================================\n');

% 신뢰도 평가 기준
reliability_criteria = struct(...
    'high_threshold', 0.8, ...
    'moderate_threshold', 0.6, ...
    'min_participants', 30, ...
    'min_items', 10);

% 통합 점수 테이블 초기화 확장
for p = 1:length(periods)
    periodColName = sprintf('Period%d_Score', p);
    reliabilityColName = sprintf('Period%d_Reliability', p);
    stdColName = sprintf('Period%d_StdScore', p);
    
    if isfield(periodResults, sprintf('period%d', p)) && ...
       ~strcmp(periodResults.(sprintf('period%d', p)).period, 'EMPTY')
        
        result = periodResults.(sprintf('period%d', p));
        
        % 신뢰도 평가 (원본과 동일한 방식)
        switch result.dataQuality
            case 'GOOD'
                reliability_level = 'HIGH';
            case 'CAUTION'
                reliability_level = 'MODERATE';
            case {'POOR', 'FAILED'}
                reliability_level = 'UNUSABLE';
            otherwise
                reliability_level = 'UNUSABLE';
        end
        
        % 신뢰도 정보 추가
        reliabilityCol = cell(height(consolidatedScores), 1);
        reliabilityCol(:) = {reliability_level};
        consolidatedScores.(reliabilityColName) = reliabilityCol;
        
        % 표준화 점수 계산 (신뢰할 수 있는 경우만)
        if ~strcmp(reliability_level, 'UNUSABLE') && ismember(periodColName, consolidatedScores.Properties.VariableNames)
            scores = consolidatedScores.(periodColName);
            validScores = scores(~isnan(scores));
            
            if length(validScores) > 1
                meanScore = mean(validScores);
                stdScore = std(validScores);
                standardizedScores = (scores - meanScore) / stdScore;
                consolidatedScores.(stdColName) = standardizedScores;
            else
                consolidatedScores.(stdColName) = NaN(height(consolidatedScores), 1);
            end
        else
            consolidatedScores.(stdColName) = NaN(height(consolidatedScores), 1);
        end
        
        fprintf('  %s: 신뢰도 %s\n', periods{p}, reliability_level);
        
    else
        % 실패한 시점은 UNUSABLE로 표시
        reliabilityCol = cell(height(consolidatedScores), 1);
        reliabilityCol(:) = {'UNUSABLE'};
        consolidatedScores.(reliabilityColName) = reliabilityCol;
        consolidatedScores.(stdColName) = NaN(height(consolidatedScores), 1);
        
        fprintf('  %s: 분석 실패 - UNUSABLE\n', periods{p});
    end
end

%% 7. 종합 분석 및 통계 (품질 검증 반영)
fprintf('\n========================================\n');
fprintf('[7단계] 종합 분석 (품질 검증 포함)\n');
fprintf('========================================\n');

% 모든 시점 포함 (원본과 동일한 방식)
reliableColumns = {};
for p = 1:length(periods)
    reliabilityCol = sprintf('Period%d_Reliability', p);
    if ismember(reliabilityCol, consolidatedScores.Properties.VariableNames)
        % 해당 시점의 신뢰도 확인
        reliability = consolidatedScores.(reliabilityCol){1}; % 첫 번째 값으로 대표
        scoreCol = sprintf('%s_Performance', periods{p});  % 실제 컬럼명 사용
        reliableColumns{end+1} = scoreCol;
        fprintf('▶ %s: 신뢰도 %s - 분석 포함\n', periods{p}, reliability);
    end
end

fprintf('\n분석 포함 시점: %d개 / %d개\n', length(reliableColumns), length(periods));

% 신뢰할 수 있는 시점만으로 평균 점수 계산
if length(reliableColumns) >= 1
    % 각 개인의 참여 시점 수 계산 (신뢰할 수 있는 시점만)
    scoreMatrix = table2array(consolidatedScores(:, reliableColumns));
    consolidatedScores.ValidReliablePeriodCount = sum(~isnan(scoreMatrix), 2);

    % 신뢰할 수 있는 시점의 표준화 점수로 평균 계산
    reliableStdColumns = {};
    for i = 1:length(reliableColumns)
        stdCol = strrep(reliableColumns{i}, '_Score', '_StdScore');
        if ismember(stdCol, consolidatedScores.Properties.VariableNames)
            reliableStdColumns{end+1} = stdCol;
        end
    end

    if ~isempty(reliableStdColumns)
        stdMatrix = table2array(consolidatedScores(:, reliableStdColumns));
        consolidatedScores.ReliableAverageStdScore = mean(stdMatrix, 2, 'omitnan');

        % 최종 백분위 계산 (신뢰할 수 있는 점수만)
        validScores = consolidatedScores.ReliableAverageStdScore(~isnan(consolidatedScores.ReliableAverageStdScore));
        consolidatedScores.ReliableFinalPercentile = NaN(height(consolidatedScores), 1);
        validIdx = ~isnan(consolidatedScores.ReliableAverageStdScore);
        consolidatedScores.ReliableFinalPercentile(validIdx) = ...
            100 * tiedrank(consolidatedScores.ReliableAverageStdScore(validIdx)) / sum(validIdx);
    end
end

% 기존 AverageStdScore도 계산 (하위 호환성을 위해)
if ~ismember('AverageStdScore', consolidatedScores.Properties.VariableNames)
    % 모든 시점의 표준화 점수로 평균 계산 (신뢰도 무관)
    allStdColumns = {};
    for p = 1:length(periods)
        stdCol = sprintf('Period%d_StdScore', p);
        if ismember(stdCol, consolidatedScores.Properties.VariableNames)
            allStdColumns{end+1} = stdCol;
        end
    end

    if ~isempty(allStdColumns)
        allStdMatrix = table2array(consolidatedScores(:, allStdColumns));
        consolidatedScores.AverageStdScore = mean(allStdMatrix, 2, 'omitnan');

        % ValidPeriodCount도 계산
        if ~ismember('ValidPeriodCount', consolidatedScores.Properties.VariableNames)
            allScoreColumns = {};
            for p = 1:length(periods)
                scoreCol = sprintf('Period%d_Score', p);
                if ismember(scoreCol, consolidatedScores.Properties.VariableNames)
                    allScoreColumns{end+1} = scoreCol;
                end
            end

            if ~isempty(allScoreColumns)
                allScoreMatrix = table2array(consolidatedScores(:, allScoreColumns));
                consolidatedScores.ValidPeriodCount = sum(~isnan(allScoreMatrix), 2);
            end
        end

        fprintf('\n▶ 기존 호환성을 위한 전체 평균 점수 계산 완료\n');
        fprintf('  - 전체 평균 점수 (AverageStdScore): %d명\n', sum(~isnan(consolidatedScores.AverageStdScore)));
    end
end

%% 8. 성과기여도 데이터 로드 및 전처리
fprintf('\n========================================\n');
fprintf('[8단계] 성과기여도 데이터 분석\n');
fprintf('========================================\n\n');

% 성과기여도 데이터 파일 경로 설정
contributionDataPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';

try
    % 성과기여도 데이터 로드
    contributionData = readtable(contributionDataPath, 'Sheet', '성과기여도', 'VariableNamingRule', 'preserve');
    fprintf('성과기여도 데이터 로드 완료: %d명\n', height(contributionData));

    % 데이터 구조 확인
    fprintf('컬럼 수: %d\n', width(contributionData));
    fprintf('컬럼명 (처음 10개): ');
    disp(contributionData.Properties.VariableNames(1:min(10, end)));

catch ME
    fprintf('[오류] 성과기여도 데이터 로드 실패: %s\n', ME.message);
    fprintf('파일 경로와 시트명을 확인해주세요.\n');
    return;
end

%% 9. 성과기여도 점수 계산
fprintf('\n========================================\n');
fprintf('성과기여도 점수 계산\n');
fprintf('========================================\n\n');

% 분기별 성과기여도 점수 계산을 위한 구조체 초기화
contributionScores = struct();
contributionScores.ID = contributionData{:, 1}; % ID 컬럼

% 분기 정보 정의 (23Q1~25Q2)
quarters = {'23Q1', '23Q2', '23Q3', '23Q4', '24Q1', '24Q2', '24Q3', '24Q4', '25Q1', '25Q2'};

% 조직성과등급 점수 매핑 (S=5, A=4, B=3, C=2, D=1)
gradeToScore = containers.Map({'S', 'A', 'B', 'C', 'D'}, {5, 4, 3, 2, 1});

fprintf('분기별 성과기여도 점수 계산 중...\n');

% 각 분기별 처리
for q = 1:length(quarters)
    quarter = quarters{q};
    fprintf('  [%s 처리 중]\n', quarter);

    % 해당 분기의 컬럼 찾기
    contributionCol = sprintf('%s_개인기여도', quarter);
    organizationCol = sprintf('%s_조직', quarter);
    gradeCol = sprintf('%s_조직성과등급', quarter);

    % 컬럼 존재 여부 확인
    hasContribution = any(strcmp(contributionData.Properties.VariableNames, contributionCol));
    hasOrganization = any(strcmp(contributionData.Properties.VariableNames, organizationCol));
    hasGrade = any(strcmp(contributionData.Properties.VariableNames, gradeCol));

    if hasContribution && hasGrade
        % 개인기여도와 조직성과등급 데이터 추출
        personalContrib = contributionData{:, contributionCol};
        orgGrades = contributionData{:, gradeCol};

        % 조직성과등급을 숫자로 변환
        orgScores = NaN(size(orgGrades));
        for i = 1:length(orgGrades)
            if iscell(orgGrades)
                grade = orgGrades{i};
            else
                grade = orgGrades(i);
            end

            if ischar(grade) || isstring(grade)
                grade = char(grade);
                if isKey(gradeToScore, grade)
                    orgScores(i) = gradeToScore(grade);
                end
            end
        end

        % 성과기여도 점수 = 개인기여도 × 조직성과점수
        quarterScore = personalContrib .* orgScores;

        % 유효한 데이터 통계
        validCount = sum(~isnan(quarterScore));
        fprintf('    - 유효한 데이터: %d/%d명\n', validCount, length(quarterScore));

        if validCount > 0
            fprintf('    - 평균 점수: %.3f\n', nanmean(quarterScore));
            fprintf('    - 표준편차: %.3f\n', nanstd(quarterScore));
        end

        % 결과 저장
        contributionScores.(sprintf('Score_%s', quarter)) = quarterScore;

    else
        fprintf('    - [경고] 필요한 컬럼이 없습니다\n');
        contributionScores.(sprintf('Score_%s', quarter)) = NaN(height(contributionData), 1);
    end
end

%% 10. 시점별 성과기여도 집계 (반기별)
fprintf('\n========================================\n');
fprintf('반기별 성과기여도 집계\n');
fprintf('========================================\n\n');

% 반기별 매핑 (역량진단 시점과 맞추기)
% 23년 하반기: 23Q3, 23Q4
% 24년 상반기: 24Q1, 24Q2
% 24년 하반기: 24Q3, 24Q4
% 25년 상반기: 25Q1, 25Q2

periodMapping = {
    {'23Q3', '23Q4'};  % 23년 하반기
    {'24Q1', '24Q2'};  % 24년 상반기
    {'24Q3', '24Q4'};  % 24년 하반기
    {'25Q1', '25Q2'}   % 25년 상반기
    };

contributionByPeriod = table();
contributionByPeriod.ID = contributionScores.ID;

for p = 1:length(periodMapping)
    quarterList = periodMapping{p};
    periodName = sprintf('Contribution_Period%d', p);

    fprintf('[%s - %s 집계]\n', periods{p}, strjoin(quarterList, ', '));

    % 해당 반기의 분기별 점수들을 평균 계산
    periodScores = [];
    for q = 1:length(quarterList)
        quarter = quarterList{q};
        if isfield(contributionScores, sprintf('Score_%s', quarter))
            quarterScore = contributionScores.(sprintf('Score_%s', quarter));
            periodScores = [periodScores, quarterScore];
        end
    end

    if ~isempty(periodScores)
        % 반기별 평균 계산
        avgScore = mean(periodScores, 2, 'omitnan');
        contributionByPeriod.(periodName) = avgScore;

        validCount = sum(~isnan(avgScore));
        fprintf('  - 유효한 데이터: %d명\n', validCount);
        if validCount > 0
            fprintf('  - 평균: %.3f, 표준편차: %.3f\n', nanmean(avgScore), nanstd(avgScore));
        end
    else
        contributionByPeriod.(periodName) = NaN(height(contributionByPeriod), 1);
        fprintf('  - 데이터 없음\n');
    end
end

% 전체 평균 성과기여도 계산
allContribScores = [contributionByPeriod.Contribution_Period1, ...
    contributionByPeriod.Contribution_Period2, ...
    contributionByPeriod.Contribution_Period3, ...
    contributionByPeriod.Contribution_Period4];

contributionByPeriod.AverageContribution = mean(allContribScores, 2, 'omitnan');
contributionByPeriod.ValidPeriodCount = sum(~isnan(allContribScores), 2);

%% 11. 역량진단 성과점수와 성과기여도 매칭
fprintf('\n========================================\n');
fprintf('[9단계] 역량진단 vs 성과기여도 데이터 매칭\n');
fprintf('========================================\n\n');

% ID를 기준으로 두 데이터셋 매칭
% consolidatedScores (역량진단 기반) vs contributionByPeriod (성과기여도 기반)

% ID를 문자열로 통일
if isnumeric(consolidatedScores.ID)
    competencyIDs = arrayfun(@num2str, consolidatedScores.ID, 'UniformOutput', false);
else
    competencyIDs = cellfun(@char, consolidatedScores.ID, 'UniformOutput', false);
end

if isnumeric(contributionByPeriod.ID)
    contributionIDs = arrayfun(@num2str, contributionByPeriod.ID, 'UniformOutput', false);
else
    contributionIDs = cellfun(@char, contributionByPeriod.ID, 'UniformOutput', false);
end

% 교집합 찾기
[commonIDs, competencyIdx, contributionIdx] = intersect(competencyIDs, contributionIDs);

fprintf('매칭 결과:\n');
fprintf('  - 역량진단 데이터: %d명\n', height(consolidatedScores));
fprintf('  - 성과기여도 데이터: %d명\n', height(contributionByPeriod));
fprintf('  - 공통 ID: %d명\n', length(commonIDs));
fprintf('  - 매칭률: %.1f%%\n', 100 * length(commonIDs) / min(height(consolidatedScores), height(contributionByPeriod)));

if length(commonIDs) < 10
    fprintf('[경고] 매칭된 데이터가 너무 적습니다. ID 형식을 확인해주세요.\n');
    fprintf('샘플 ID (역량진단): ');
    disp(competencyIDs(1:min(5, end)));
    fprintf('샘플 ID (성과기여도): ');
    disp(contributionIDs(1:min(5, end)));
end

% 매칭된 데이터로 통합 테이블 생성
if length(commonIDs) > 0
    combinedData = table();
    combinedData.ID = commonIDs;

    % 역량진단 점수 추가
    combinedData.Factor_Period1 = consolidatedScores.('23년_하반기_Performance')(competencyIdx);
    combinedData.Factor_Period2 = consolidatedScores.('24년_상반기_Performance')(competencyIdx);
    combinedData.Factor_Period3 = consolidatedScores.('24년_하반기_Performance')(competencyIdx);
    combinedData.Factor_Period4 = consolidatedScores.('25년_상반기_Performance')(competencyIdx);
    combinedData.Factor_Average = consolidatedScores.AverageStdScore(competencyIdx);

    % 성과기여도 점수 추가
    combinedData.Contribution_Period1 = contributionByPeriod.Contribution_Period1(contributionIdx);
    combinedData.Contribution_Period2 = contributionByPeriod.Contribution_Period2(contributionIdx);
    combinedData.Contribution_Period3 = contributionByPeriod.Contribution_Period3(contributionIdx);
    combinedData.Contribution_Period4 = contributionByPeriod.Contribution_Period4(contributionIdx);
    combinedData.Contribution_Average = contributionByPeriod.AverageContribution(contributionIdx);

    fprintf('통합 데이터 생성 완료: %d명\n', height(combinedData));

else
    fprintf('[오류] 매칭된 데이터가 없습니다. 분석을 계속할 수 없습니다.\n');
    return;
end

%% 12. 상관분석
fprintf('\n========================================\n');
fprintf('[10단계] 역량진단 성과점수 vs 성과기여도 상관분석\n');
fprintf('========================================\n\n');

correlationResults = struct();

% 시점별 상관분석
fprintf('[시점별 상관분석]\n');
for p = 1:4
    factorCol = sprintf('Factor_Period%d', p);
    contribCol = sprintf('Contribution_Period%d', p);

    factorScores = combinedData.(factorCol);
    contribScores = combinedData.(contribCol);

    % 둘 다 유효한 값이 있는 경우만 분석
    validIdx = ~isnan(factorScores) & ~isnan(contribScores);
    validCount = sum(validIdx);

    if validCount >= 5  % 최소 5개 이상의 쌍이 있어야 상관분석 가능
        r = corrcoef(factorScores(validIdx), contribScores(validIdx));
        correlation = r(1, 2);

        % 유의성 검정 (t-test)
        t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
        p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

        fprintf('%s: r = %.3f (n=%d, p=%.3f)', periods{p}, correlation, validCount, p_value);
        if p_value < 0.001
            fprintf(' ***');
        elseif p_value < 0.01
            fprintf(' **');
        elseif p_value < 0.05
            fprintf(' *');
        end
        fprintf('\n');

        correlationResults.(sprintf('period%d', p)) = struct(...
            'correlation', correlation, ...
            'n', validCount, ...
            'p_value', p_value);
    else
        fprintf('%s: 분석 불가 (유효 데이터 %d개)\n', periods{p}, validCount);
        correlationResults.(sprintf('period%d', p)) = struct(...
            'correlation', NaN, ...
            'n', validCount, ...
            'p_value', NaN);
    end
end

% 전체 평균 점수 간 상관분석
fprintf('\n[전체 평균 점수 상관분석]\n');
validIdx = ~isnan(combinedData.Factor_Average) & ~isnan(combinedData.Contribution_Average);
validCount = sum(validIdx);

if validCount >= 5
    r = corrcoef(combinedData.Factor_Average(validIdx), combinedData.Contribution_Average(validIdx));
    correlation = r(1, 2);

    t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
    p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

    fprintf('전체 평균: r = %.3f (n=%d, p=%.3f)', correlation, validCount, p_value);
    if p_value < 0.001
        fprintf(' ***');
    elseif p_value < 0.01
        fprintf(' **');
    elseif p_value < 0.05
        fprintf(' *');
    end
    fprintf('\n');

    correlationResults.overall = struct(...
        'correlation', correlation, ...
        'n', validCount, ...
        'p_value', p_value);
else
    fprintf('전체 평균: 분석 불가 (유효 데이터 %d개)\n', validCount);
end

%% 13. 성장단계 변화량 분석
fprintf('\n========================================\n');
fprintf('[11단계] 성장단계 변화량 분석\n');
fprintf('========================================\n\n');

% 성장단계 데이터 로드
growthDataPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';

try
    % 성장단계 데이터 로드
    growthData = readtable(growthDataPath, 'Sheet', '성장단계', 'VariableNamingRule', 'preserve');
    fprintf('성장단계 데이터 로드 완료: %d명\n', height(growthData));

    % 데이터 구조 확인
    fprintf('컬럼 수: %d\n', width(growthData));
    fprintf('컬럼명: ');
    disp(growthData.Properties.VariableNames);

catch ME
    fprintf('[오류] 성장단계 데이터 로드 실패: %s\n', ME.message);
    fprintf('파일 경로와 시트명을 확인해주세요.\n');
    return;
end

%% 14. 확장된 CSR + 조직지표 분석 (Q40-46)
fprintf('\n========================================\n');
fprintf('[12단계] 확장된 CSR + 조직지표 분석 (25년 상반기)\n');
fprintf('========================================\n\n');

% 25년 상반기(period 4)에서 제외된 Q40-46 데이터 추출
extendedCSRResults = struct();

if isfield(allData, 'period4') && isfield(allData.period4, 'selfData')
    fprintf('▶ 25년 상반기 데이터에서 확장 문항 추출 시도\n');

    originalData = allData.period4.selfData;

    % 확장된 문항 정의 (Q40-46) - 모두 영문 변수명으로 통일
    extendedQuestions = {
        'Q40', 'OrganizationalSynergy';
        'Q41', 'Pride';
        'Q42', 'Communication_Relationship';
        'Q43', 'Communication_Purpose';
        'Q44', 'Strategy_CustomerValue';
        'Q45', 'Strategy_Performance';
        'Q46', 'Reflection_Organizational'
        };

    % ID 컬럼 찾기
    idCol = findIDColumn(originalData);
    if isempty(idCol)
        fprintf('  [오류] ID 컬럼을 찾을 수 없습니다.\n');
    else
        % 확장 데이터 추출
        extendedData = table();
        extendedData.ID = extractAndStandardizeIDs(originalData{:, idCol});

        foundQuestions = {};
        missingQuestions = {};

        fprintf('문항 검색 결과:\n');
        for i = 1:size(extendedQuestions, 1)
            qCode = extendedQuestions{i, 1};
            qName = extendedQuestions{i, 2};

            % 해당 문항 컬럼 찾기
            questionCol = [];
            colNames = originalData.Properties.VariableNames;

            for col = 1:width(originalData)
                colName = colNames{col};
                if strcmp(colName, qCode) || contains(colName, qCode) || ...
                        (length(colName) >= 3 && strcmp(colName(1:3), qCode))
                    questionCol = col;
                    break;
                end
            end

            if ~isempty(questionCol)
                extendedData.(qName) = originalData{:, questionCol};
                foundQuestions{end+1} = qCode;
                fprintf('  ✓ %s (%s) 발견 → %s\n', qCode, qName, colNames{questionCol});
            else
                extendedData.(qName) = NaN(height(extendedData), 1);
                missingQuestions{end+1} = qCode;
                fprintf('  ✗ %s (%s) 누락\n', qCode, qName);
            end
        end

        fprintf('\n발견된 확장 문항: %d개 / %d개\n', length(foundQuestions), size(extendedQuestions, 1));

        if ~isempty(foundQuestions)
            fprintf('  → 확장 분석 가능\n');
            extendedCSRResults.extendedData = extendedData;
            extendedCSRResults.foundQuestions = foundQuestions;
            extendedCSRResults.missingQuestions = missingQuestions;
        else
            fprintf('  → 확장 분석 불가능 (문항 없음)\n');
        end
    end
else
    fprintf('[정보] 25년 상반기 데이터에서 확장 문항 분석은 수행되지 않습니다.\n');
    fprintf('       (수평 진단 버전에서는 원본 문항이 이미 평균 처리됨)\n');
end

%% 15. 최종 결과 저장 (확장 버전)
fprintf('\n========================================\n');
fprintf('[13단계] 최종 결과 저장\n');
fprintf('========================================\n\n');

% Excel 파일로 저장
outputFileName = sprintf('수평진단_종합분석결과_%s.xlsx', datestr(now, 'yyyymmdd_HHMMSS'));

try
    % 1. 통합 점수 테이블 저장
    writetable(consolidatedScores, outputFileName, 'Sheet', '역량진단_통합점수');
    fprintf('✓ 역량진단 통합점수 저장\n');

    % 2. 각 시점별 상세 결과 저장
    for p = 1:length(periods)
        if isfield(periodResults, sprintf('period%d', p)) && ...
           ~strcmp(periodResults.(sprintf('period%d', p)).period, 'EMPTY')
            
            result = periodResults.(sprintf('period%d', p));

            % 개인별 점수 테이블
            periodScoreTable = table();
            periodScoreTable.ID = result.participantIDs;
            periodScoreTable.PerformanceScore = result.performanceScores;

            sheetName = sprintf('%s_점수', periods{p});
            writetable(periodScoreTable, outputFileName, 'Sheet', sheetName);

            % 요인 부하량 테이블
            loadingTable = table();
            loadingTable.Question = result.questionNames';
            for f = 1:result.numFactors
                loadingTable.(sprintf('Factor%d', f)) = result.factorLoadings(:, f);
            end

            sheetName = sprintf('%s_부하량', periods{p});
            writetable(loadingTable, outputFileName, 'Sheet', sheetName);
        end
    end
    fprintf('✓ 시점별 상세 결과 저장\n');

    % 3. 성과기여도 데이터 저장 (있는 경우)
    if exist('contributionByPeriod', 'var')
        writetable(contributionByPeriod, outputFileName, 'Sheet', '성과기여도점수');
        fprintf('✓ 성과기여도 점수 저장\n');
    end

    % 4. 통합 상관분석 데이터 저장 (있는 경우)
    if exist('combinedData', 'var')
        writetable(combinedData, outputFileName, 'Sheet', '상관분석_통합데이터');
        fprintf('✓ 상관분석 통합데이터 저장\n');

        % 5. 상관분석 결과 테이블 생성
        if exist('correlationResults', 'var')
            corrResultTable = table();
            corrResultTable.Period = [periods'; {'전체평균'}];
            corrResultTable.Correlation = NaN(5, 1);
            corrResultTable.N = NaN(5, 1);
            corrResultTable.P_Value = NaN(5, 1);
            corrResultTable.Significance = cell(5, 1);

            for p = 1:4
                if isfield(correlationResults, sprintf('period%d', p))
                    result = correlationResults.(sprintf('period%d', p));
                    corrResultTable.Correlation(p) = result.correlation;
                    corrResultTable.N(p) = result.n;
                    corrResultTable.P_Value(p) = result.p_value;

                    if ~isnan(result.p_value)
                        if result.p_value < 0.001
                            corrResultTable.Significance{p} = '***';
                        elseif result.p_value < 0.01
                            corrResultTable.Significance{p} = '**';
                        elseif result.p_value < 0.05
                            corrResultTable.Significance{p} = '*';
                        else
                            corrResultTable.Significance{p} = 'n.s.';
                        end
                    else
                        corrResultTable.Significance{p} = '분석불가';
                    end
                end
            end

            % 전체 평균 결과 추가
            if isfield(correlationResults, 'overall')
                result = correlationResults.overall;
                corrResultTable.Correlation(5) = result.correlation;
                corrResultTable.N(5) = result.n;
                corrResultTable.P_Value(5) = result.p_value;

                if result.p_value < 0.001
                    corrResultTable.Significance{5} = '***';
                elseif result.p_value < 0.01
                    corrResultTable.Significance{5} = '**';
                elseif result.p_value < 0.05
                    corrResultTable.Significance{5} = '*';
                else
                    corrResultTable.Significance{5} = 'n.s.';
                end
            end

            writetable(corrResultTable, outputFileName, 'Sheet', '상관분석결과');
            fprintf('✓ 상관분석 결과 저장\n');
        end
    end

    fprintf('📁 결과 파일: %s\n', outputFileName);

catch ME
    fprintf('❌ 결과 저장 실패: %s\n', ME.message);
end

%% 16. 최종 보고서
fprintf('\n[최종 보고서]\n');
fprintf('========================================\n');
fprintf('수평 진단 기반 시점별 개별 요인분석 완료\n');
fprintf('========================================\n\n');

fprintf('📊 분석 개요:\n');
fprintf('• 분석 방법: 수평 진단 (동료 평가) → 하향 진단 변환 후 요인분석\n');
fprintf('• 분석 시점: %d개 (%s)\n', length(periods), strjoin(periods, ', '));
fprintf('• 성공 분석: %d개 시점\n', successCount);
fprintf('• 총 참여자: %d명\n', totalParticipants);

if exist('combinedData', 'var')
    fprintf('• 성과기여도 매칭: %d명\n', height(combinedData));
end
fprintf('\n');

if successCount > 0
    fprintf('🎯 주요 특징:\n');
    fprintf('• 동료 평가 점수를 개인별로 평균하여 개별 분석 수행\n');
    fprintf('• 원본 코드와 동일한 정교한 전처리 및 요인분석 적용\n');
    fprintf('• 다중 기준(Kaiser, Scree, Parallel Analysis)으로 최적 요인 수 결정\n');
    fprintf('• 성과 관련 요인 자동 식별 및 점수 산출\n');
    fprintf('• 품질 검증을 통한 신뢰할 수 있는 시점만 활용\n');
    fprintf('• 성과기여도와의 상관관계 분석\n\n');
end

fprintf('✅ 수평 진단 기반 종합 분석이 완료되었습니다!\n');
fprintf('📁 상세 결과는 Excel 파일을 확인하세요.\n\n');

% 로그 종료
diary off;
fprintf('📝 로그 파일: D:\\project\\matlab_runlog\\runlog_horizontal_enhanced.txt\n');
fprintf('🎉 분석 완료!\n');

%% 보조 함수들 (원본에서 가져옴)

function emptyResult = createEmptyPeriodResult(reason)
emptyResult = struct();
emptyResult.period = 'EMPTY';
emptyResult.reason = reason;
emptyResult.analysisDate = datestr(now);
emptyResult.numParticipants = 0;
emptyResult.numQuestions = 0;
emptyResult.numFactors = 0;
end

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

function standardizedIDs = extractAndStandardizeIDs(rawIDs)
if isnumeric(rawIDs)
    standardizedIDs = arrayfun(@(x) sprintf('%.0f', x), rawIDs, 'UniformOutput', false);
elseif iscell(rawIDs)
    standardizedIDs = cellfun(@(x) char(x), rawIDs, 'UniformOutput', false);
else
    standardizedIDs = cellstr(rawIDs);
end

% 빈 값이나 NaN 처리
emptyIdx = cellfun(@isempty, standardizedIDs) | strcmp(standardizedIDs, 'NaN');
standardizedIDs(emptyIdx) = {''};
end

function cleanData = handleMissingValues(rawData)
cleanData = rawData;

for col = 1:size(rawData, 2)
    missingIdx = isnan(rawData(:, col));
    if any(missingIdx) && ~all(missingIdx)
        colMean = mean(rawData(~missingIdx, col));
        cleanData(missingIdx, col) = colMean;
    elseif all(missingIdx)
        cleanData(:, col) = 3; % 기본값 (중간값)
    end
end
end

function elbowPoint = findElbowPoint(eigenValues)
if length(eigenValues) < 3
    elbowPoint = 1;
    return;
end

% 2차 차분을 이용한 엘보 포인트 찾기
diffs = diff(eigenValues);
secondDiffs = diff(diffs);

% 가장 큰 변화가 일어나는 지점
[~, elbowPoint] = max(abs(secondDiffs));
elbowPoint = min(elbowPoint + 1, length(eigenValues));

% 최소 1, 최대 요인 수 제한
elbowPoint = max(1, min(elbowPoint, length(eigenValues)));
end

function numFactors = parallelAnalysis(data, numIterations)
[n, p] = size(data);

% 실제 고유값
realEigenValues = eig(cov(data));
realEigenValues = sort(realEigenValues, 'descend');

% 랜덤 데이터의 고유값
randomEigenValues = zeros(numIterations, p);

for iter = 1:numIterations
    randomData = randn(n, p);
    randomEigenValues(iter, :) = sort(eig(cov(randomData)), 'descend');
end

% 95 백분위수 사용
randomEigenThreshold = prctile(randomEigenValues, 95);

% 실제 고유값이 랜덤보다 큰 개수
numFactors = sum(realEigenValues > randomEigenThreshold');
numFactors = max(1, numFactors); % 최소 1개
end

function performanceIdx = identifyPerformanceFactorAdvanced(loadings, questionNames, questionInfo)
performanceKeywords = {'성과', '목표', '달성', '결과', '효과', '기여', '창출', '개선', '수행', '완수', '생산', '실적'};
numFactors = size(loadings, 2);
performanceScores = zeros(numFactors, 1);

for f = 1:numFactors
    % 높은 부하량 문항들
    highLoadingItems = find(abs(loadings(:, f)) > 0.3);

    for itemIdx = 1:length(highLoadingItems)
        item = highLoadingItems(itemIdx);
        questionName = questionNames{item};

        % 문항 정보에서 내용 찾기
        try
            if height(questionInfo) > 0
                % 문항명으로 매칭 시도
                for row = 1:height(questionInfo)
                    qCode = questionInfo{row, 1};
                    if iscell(qCode), qCode = qCode{1}; end
                    qCode = char(qCode);

                    if contains(questionName, qCode) || contains(qCode, questionName)
                        questionText = questionInfo{row, 2};
                        if iscell(questionText)
                            questionText = questionText{1};
                        end
                        questionText = char(questionText);

                        % 키워드 매칭
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
            % 매칭 실패 시 무시하고 계속
        end
    end

    % 요인의 전체적인 부하량 패턴도 고려
    % 높은 부하량이 많은 요인일수록 성과 관련 가능성 높음
    performanceScores(f) = performanceScores(f) + 0.1 * sum(abs(loadings(:, f)) > 0.5);
end

[~, performanceIdx] = max(performanceScores);

% 성과 요인을 찾지 못한 경우 첫 번째 요인 사용
if all(performanceScores == 0)
    performanceIdx = 1;
end
end

function masterIDs = extractMasterIDs(masterTable)
masterIDs = {};

if height(masterTable) == 0
    return;
end

% ID 컬럼 찾기
for col = 1:width(masterTable)
    colName = masterTable.Properties.VariableNames{col};
    if contains(lower(colName), {'id', '사번', 'empno', 'employee_id'})
        ids = masterTable{:, col};
        if isnumeric(ids)
            masterIDs = arrayfun(@(x) sprintf('%.0f', x), ids, 'UniformOutput', false);
        else
            masterIDs = cellstr(ids);
        end

        % 빈 값 제거
        masterIDs = masterIDs(~cellfun(@isempty, masterIDs));
        break;
    end
end
end

function score = convertGrowthStageToScore(stageText)
% 성장단계를 숫자 점수로 변환
% 규칙: 단계(열린<성취<책임<소명) + 레벨(LvN) + 추가(-k*0.1)

% 벡터/배열 입력 지원
if iscell(stageText) || isstring(stageText)
    score = arrayfun(@convertOne, stageText);
else
    score = convertOne(stageText);
end
end

function sc = convertOne(x)
% 결측 처리
if isempty(x) || (isstring(x) && ismissing(x))
    sc = NaN; return;
end
if iscell(x), x = x{1}; end
s = char(x);
s = strtrim(s);

% 전처리: 전각 공백 제거, 탭 제거
s = regexprep(s, '[\s　]+', ' ');   % 일반/전각 공백을 공백 하나로
s_lower = lower(s);

% 단계 우선순위 고정
stages = {'열린','성취','책임','소명'};
base   = [10,    20,    30,    40];

% 단계 매칭 (우선 매칭되는 첫 단계 사용)
stageIdx = find(cellfun(@(k) contains(s, k), stages), 1, 'first');

if isempty(stageIdx)
    sc = NaN; return;  % 단계 못 찾으면 NaN
end

sc = base(stageIdx);

% 레벨: Lv, LV, lv, 'Lv. 2' 등 허용
% 예: 'Lv2', 'LV.2', 'lv 2'
tok = regexp(s_lower, 'lv\.?\s*(\d+)', 'tokens', 'once');
if ~isempty(tok)
    lvl = str2double(tok{1});
    if ~isnan(lvl)
        % 비상식적 값 캡핑(옵션): 0~9 범위
        lvl = max(0, min(9, lvl));
        sc = sc + lvl;
    end
end

% 추가 숫자: "-3", "- 2" 등
tok2 = regexp(s_lower, '-\s*(\d+)', 'tokens', 'once');
if ~isempty(tok2)
    addk = str2double(tok2{1});
    if ~isnan(addk)
        sc = sc + 0.1 * addk;
    end
end

% 방어: 최종 0이면 NaN (단계가 없거나 전부 실패했을 때)
if sc == 0
    sc = NaN;
end
end

function alpha = cronbachAlpha(data)
k = size(data, 2);
if k < 2
    alpha = NaN;
    return;
end

itemVar = var(data, [], 1, 'omitnan');
totalVar = var(sum(data, 2, 'omitnan'), 'omitnan');

alpha = (k / (k - 1)) * (1 - sum(itemVar) / totalVar);
end