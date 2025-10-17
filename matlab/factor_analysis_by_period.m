1%% 시점별 개별 요인분석 기반 역량진단 성과점수 산출 및 성과기여도 상관분석
% 2023년 하반기 ~ 2025년 상반기 (4개 시점) 데이터 개별 분석
%
% 작성일: 2025년
% 목적: 각 시점별로 독립적인 요인분석 수행 후 개별 점수 산출 및 성과기여도와 상관분석
cd('D:\project\HR데이터\matlab')
clear; clc; close all;
diary('D:\project\matlab_runlog\runlog.txt');



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
fprintf('========================================\n\n');

%% 2. 데이터 로드 (면담용 문항 자동 제외 기능 추가)
fprintf('[1단계] 모든 시점 데이터 로드\n');
fprintf('----------------------------------------\n');

for p = 1:length(periods)
    fprintf('▶ %s 데이터 로드 중...\n', periods{p});
    fileName = fullfile(dataPath, fileNames{p});

    try
        % 기본 데이터 로드
        allData.(sprintf('period%d', p)).masterIDs = ...
            readtable(fileName, 'Sheet', '기준인원 검토', 'VariableNamingRule', 'preserve');
        allData.(sprintf('period%d', p)).selfData = ...
            readtable(fileName, 'Sheet', '하향 진단', 'VariableNamingRule', 'preserve');
        allData.(sprintf('period%d', p)).questionInfo = ...
            readtable(fileName, 'Sheet', '문항 정보_타인진단', 'VariableNamingRule', 'preserve');

        fprintf('  ✓ 마스터ID: %d명, 하향진단: %d명\n', ...
            height(allData.(sprintf('period%d', p)).masterIDs), ...
            height(allData.(sprintf('period%d', p)).selfData));

        % 면담용 문항 제외 처리 (모든 시점에 적용)
        fprintf('\n▶ %s: 면담용 문항 제외 처리\n', periods{p});

        selfData = allData.(sprintf('period%d', p)).selfData;
        questionInfo = allData.(sprintf('period%d', p)).questionInfo;

        % 먼저 Q로 시작하는 문항 컬럼 식별
        colNames = selfData.Properties.VariableNames;
        tempQuestionCols = {};
        for col = 1:width(selfData)
            colName = colNames{col};
            colData = selfData{:, col};

            if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
                tempQuestionCols{end+1} = colName;
            end
        end

        fprintf('  전체 문항 수: %d개\n', length(tempQuestionCols));

        % 면담용 문항 식별을 위한 키워드
        interviewKeywords = {'점수 산출에 활용되지 않으며', '면담을 위해 수집하는 정보입니다', ...
            '점수산출에 활용되지 않으며', '면담을 위해', '면담용', '인터뷰용'};

        % 25년 상반기 특별 처리: Q40-46도 함께 제외
        excludePatterns = {};
        if p == 4 % 25년 상반기
            excludePatterns = {'Q40','Q41','Q42', 'Q43', 'Q44', 'Q45', 'Q46', ...
                'q40','q41','q42', 'q43', 'q44', 'q45', 'q46'};
        end

        % 제외할 문항 찾기
        excludeIndices = [];
        excludedQuestions = {};
        excludeReasons = {}; % 제외 이유 저장

        % 1) 문항 정보에서 면담용 문항 찾기
        if height(questionInfo) > 0
            fprintf('  문항 정보에서 면담용 문항 검색 중...\n');

            for i = 1:length(tempQuestionCols)
                questionName = tempQuestionCols{i};

                % 문항 정보에서 해당 문항 찾기
                found = false;
                for row = 1:height(questionInfo)
                    try
                        % 첫 번째 컬럼에서 문항 코드 추출
                        qCode = questionInfo{row, 1};
                        if iscell(qCode)
                            qCode = qCode{1};
                        end
                        qCode = char(qCode);

                        % 문항 코드가 현재 문항과 매칭되는지 확인
                        if contains(questionName, qCode) || contains(qCode, questionName) || strcmp(questionName, qCode)

                            % 문항 설명에서 면담 관련 키워드 검색
                            questionText = '';
                            if width(questionInfo) >= 2
                                questionTextRaw = questionInfo{row, 2};
                                if iscell(questionTextRaw)
                                    questionTextRaw = questionTextRaw{1};
                                end
                                questionText = char(questionTextRaw);
                            end

                            % 추가 컬럼들도 검사 (설명이 다른 컬럼에 있을 수 있음)
                            for col = 3:min(width(questionInfo), 5)
                                if ~isempty(questionInfo{row, col})
                                    extraTextRaw = questionInfo{row, col};
                                    if iscell(extraTextRaw)
                                        extraTextRaw = extraTextRaw{1};
                                    end
                                    extraText = char(extraTextRaw);
                                    questionText = [char(questionText), ' ', extraText];
                                end
                            end

                            % 키워드 검사
                            for k = 1:length(interviewKeywords)
                                if contains(questionText, interviewKeywords{k})
                                    excludeIndices(end+1) = i;
                                    excludedQuestions{end+1} = questionName;
                                    excludeReasons{end+1} = sprintf('면담용 문항 (키워드: "%s")', interviewKeywords{k});
                                    found = true;
                                    break;
                                end
                            end

                            if found
                                break;
                            end
                        end
                    catch ME
                        % 매칭 실패 시 무시하고 계속
                        fprintf('    경고: 문항 %s 처리 중 오류 발생: %s\n', questionName, ME.message);
                        continue;
                    end
                end
            end
        end

        % 2) 25년 상반기인 경우 Q40-46 패턴 매칭 추가
        if p == 4 && ~isempty(excludePatterns)
            fprintf('  25년 상반기: Q40-46 패턴 매칭 중...\n');

            for i = 1:length(tempQuestionCols)
                questionName = tempQuestionCols{i};
                for j = 1:length(excludePatterns)
                    if strcmp(questionName, excludePatterns{j}) || ...
                            startsWith(questionName, excludePatterns{j}) || ...
                            contains(questionName, excludePatterns{j})

                        % 이미 제외 목록에 있는지 확인
                        if ~ismember(i, excludeIndices)
                            excludeIndices(end+1) = i;
                            excludedQuestions{end+1} = questionName;
                            excludeReasons{end+1} = '25년 상반기 Q40-46 제외';
                        end
                        break;
                    end
                end
            end
        end

        % 중복 제거
        [excludeIndices, uniqueIdx] = unique(excludeIndices);
        excludedQuestions = excludedQuestions(uniqueIdx);
        excludeReasons = excludeReasons(uniqueIdx);

        if ~isempty(excludeIndices)
            fprintf('  발견된 제외 대상 문항 (%d개):\n', length(excludeIndices));
            for i = 1:length(excludedQuestions)
                fprintf('    - %s (%s)\n', excludedQuestions{i}, excludeReasons{i});
            end

            % 제외할 컬럼들을 selfData에서 실제로 제거
            excludeColumnNames = tempQuestionCols(excludeIndices);

            % 원본 데이터 백업
            allData.(sprintf('period%d', p)).originalSelfData = selfData;

            % 해당 컬럼들을 테이블에서 제거
            selfData = removevars(selfData, excludeColumnNames);

            % 수정된 데이터를 다시 저장
            allData.(sprintf('period%d', p)).selfData = selfData;

            fprintf('  ✓ 면담용/제외 문항 제거 완료\n');
            fprintf('    - 제거 전: %d개 문항\n', length(tempQuestionCols));
            fprintf('    - 제거 후: %d개 문항\n', length(tempQuestionCols) - length(excludeIndices));

            % 제거된 문항 정보 저장 (나중에 참조용)
            if ~exist('excludedItemsInfo', 'var')
                excludedItemsInfo = struct();
            end
            excludedItemsInfo.(sprintf('period%d', p)) = struct(...
                'excludedQuestions', {excludedQuestions}, ...
                'excludedColumnNames', {excludeColumnNames}, ...
                'excludeReasons', {excludeReasons}, ...
                'originalNumQuestions', length(tempQuestionCols), ...
                'finalNumQuestions', length(tempQuestionCols) - length(excludeIndices));

        else
            fprintf('  → 제외 대상 문항을 찾을 수 없습니다.\n');
            if height(questionInfo) == 0
                fprintf('  → 문항 정보가 없어 면담용 문항을 식별할 수 없습니다.\n');
            end
            fprintf('  → 사용 가능한 문항들 (처음 10개):\n');
            for i = 1:min(10, length(tempQuestionCols))
                fprintf('      %s\n', tempQuestionCols{i});
            end
            if length(tempQuestionCols) > 10
                fprintf('      ... 외 %d개\n', length(tempQuestionCols) - 10);
            end
        end

    catch ME
        fprintf('  ✗ %s 데이터 로드 실패: %s\n', periods{p}, ME.message);
        return;
    end
end

% 전체 로드 완료 메시지
fprintf('\n모든 시점 데이터 로드 완료\n');
if exist('excludedItemsInfo', 'var')
    fprintf('▶ 면담용/제외 문항 처리 요약:\n');
    fields = fieldnames(excludedItemsInfo);
    totalExcluded = 0;

    for i = 1:length(fields)
        info = excludedItemsInfo.(fields{i});
        periodNum = str2double(fields{i}(end));
        numExcluded = length(info.excludedQuestions);
        totalExcluded = totalExcluded + numExcluded;

        fprintf('  %s: %d개 문항 제거 (%d→%d)\n', periods{periodNum}, ...
            numExcluded, info.originalNumQuestions, info.finalNumQuestions);

        % 제외 이유별 통계
        if ~isempty(info.excludeReasons)
            reasonTypes = unique(info.excludeReasons);
            for j = 1:length(reasonTypes)
                reasonCount = sum(strcmp(info.excludeReasons, reasonTypes{j}));
                fprintf('    └ %s: %d개\n', reasonTypes{j}, reasonCount);
            end
        end
    end

    fprintf('  전체 제외된 문항: %d개\n', totalExcluded);
else
    fprintf('▶ 제외된 문항 없음\n');
end
%% 3. 시점별 개별 요인분석 수행
fprintf('\n[2단계] 시점별 개별 요인분석 수행\n');
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

    selfData = allData.(sprintf('period%d', p)).selfData;
    questionInfo = allData.(sprintf('period%d', p)).questionInfo;

    %% 3-1. 문항 데이터 추출
    fprintf('▶ 문항 데이터 추출\n');

    % ID 컬럼 찾기
    idCol = findIDColumn(selfData);
    if isempty(idCol)
        fprintf('  [경고] ID 컬럼을 찾을 수 없습니다. 건너뜀.\n\n');
        % 빈 결과 저장 후 다음 시점으로
        periodResults.(sprintf('period%d', p)) = createEmptyPeriodResult('NO_ID_COLUMN');
        %continue;
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
        % continue;
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
        %continue;
    end

    %% 3-2. 데이터 품질 검사 및 전처리
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
        [row, col] = find(abs(R) > 0.95 & R ~= 1);

        if ~isempty(row)
            validPairs = row < col;
            row = row(validPairs);
            col = col(validPairs);

            removeIndices = [];
            for i = 1:length(row)
                if columnVariances(row(i)) < columnVariances(col(i))
                    removeIndices(end+1) = row(i);
                else
                    removeIndices(end+1) = col(i);
                end
            end

            removeIndices = unique(removeIndices);
            if ~isempty(removeIndices)
                fprintf('  [제거] 다중공선성 문항 %d개\n', length(removeIndices));
                responseData(:, removeIndices) = [];
                questionCols(removeIndices) = [];
                columnVariances(removeIndices) = [];
            end
        end
    end

    % 3. 극단값 처리
    outlierCount = 0;
    for col = 1:size(responseData, 2)
        colData = responseData(:, col);
        validData = colData(~isnan(colData));

        if length(validData) > 5
            meanVal = mean(validData);
            stdVal = std(validData);

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

    if outlierCount > 0
        fprintf('  [처리] 극단값 %d개 조정\n', outlierCount);
    end

    % 4. 저분산 문항 제거 추가
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
        %continue;
    end

    try
        R_final = corrcoef(responseData, 'Rows', 'pairwise');
        det_final = det(R_final);
        cond_final = cond(R_final);

        fprintf('  - 최종 문항 수: %d개 (원본: %d개)\n', size(responseData, 2), originalNumQuestions);
        fprintf('  - 상관행렬 조건수: %.2e\n', cond_final);

        % 품질 판정
        if (det_final > 1e-10) && (cond_final < 1e10)
            dataQualityFlag = 'GOOD';
            fprintf('  ✓ 수치적 안정성 양호\n');
        elseif (det_final > 1e-15) && (cond_final < 1e12)
            dataQualityFlag = 'CAUTION';
            fprintf('  ⚠ 수치적 문제 있음 - 주의 필요\n');
        else
            dataQualityFlag = 'POOR';
            fprintf('  ✗ 심각한 수치적 문제\n');
        end

    catch
        dataQualityFlag = 'FAILED';
        fprintf('  [경고] 상관행렬 계산 실패\n');
    end

    %% 3-3. 최적 요인 수 결정
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

    %% 3-4. 요인분석 수행
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

        numPCs = min(optimalNumFactors, size(score, 2));
        loadings = coeff(:, 1:numPCs);
        factorScores = score(:, 1:numPCs);
        optimalNumFactors = numPCs;
        isPCA = true;
    end

    %% 3-5. 성과 관련 요인 식별
    fprintf('\n▶ 성과 관련 요인 식별\n');

    try
        performanceFactorIdx = identifyPerformanceFactorAdvanced(loadings, questionCols, questionInfo);

        if isempty(performanceFactorIdx) || performanceFactorIdx < 1 || performanceFactorIdx > size(loadings, 2)
            performanceFactorIdx = 1;
            fprintf('  [경고] 성과 요인 식별 실패. 첫 번째 요인 사용\n');
        end

    catch ME
        fprintf('  [경고] 성과 요인 식별 오류: %s\n', ME.message);
        performanceFactorIdx = 1;
    end

    fprintf('  식별된 성과 요인: %d번째 요인\n', performanceFactorIdx);

    % 주요 구성 문항
    mainItems = find(abs(loadings(:, performanceFactorIdx)) > 0.4);
    if ~isempty(mainItems)
        fprintf('  주요 구성 문항 (부하량 > 0.4):\n');
        for i = 1:min(3, length(mainItems))
            loading = loadings(mainItems(i), performanceFactorIdx);
            questionName = questionCols{mainItems(i)};
            fprintf('    %s: %.3f\n', questionName, loading);
        end
    end


    % 주요 구성 문항 테이블 생성 및 저장
    if ~isempty(mainItems)
        competencyItemTable = table();
        competencyItemTable.ItemIndex = mainItems;
        competencyItemTable.QuestionName = questionCols(mainItems)';
        competencyItemTable.LoadingValue = loadings(mainItems, performanceFactorIdx);
    else
        % 빈 테이블 생성
        competencyItemTable = table();
        competencyItemTable.ItemIndex = [];
        competencyItemTable.QuestionName = {};
        competencyItemTable.LoadingValue = [];
    end


    %% 3-6. 개인별 성과점수 산출
    fprintf('\n▶ 개인별 성과점수 산출\n');

    if size(factorScores, 2) >= performanceFactorIdx
        performanceScores = factorScores(:, performanceFactorIdx);

        % 데이터 품질에 따른 신뢰도 설정
        switch dataQualityFlag
            case 'GOOD'
                standardizedScores = zscore(performanceScores);
                scoreReliability = 'HIGH';

            case 'CAUTION'
                standardizedScores = zscore(performanceScores);
                scoreReliability = 'MODERATE';
                fprintf('  [주의] 데이터 품질 문제로 신뢰도 낮음\n');

            case {'POOR', 'FAILED'}
                standardizedScores = NaN(size(performanceScores));
                scoreReliability = 'UNUSABLE';
                fprintf('  [경고] 데이터 품질 불량 - 점수 사용 불가\n');

            otherwise
                standardizedScores = zscore(performanceScores);
                scoreReliability = 'MODERATE';
        end

        % 백분위 계산
        if strcmp(scoreReliability, 'UNUSABLE')
            percentileRanks = NaN(size(performanceScores));
        else
            percentileRanks = 100 * tiedrank(performanceScores) / length(performanceScores);
        end

    else
        fprintf('  [오류] 요인 점수 추출 실패\n');
        performanceScores = NaN(length(responseIDs), 1);
        standardizedScores = NaN(length(responseIDs), 1);
        percentileRanks = NaN(length(responseIDs), 1);
        scoreReliability = 'UNUSABLE';
    end

    %% 3-7. 품질 평가
    fprintf('\n▶ 분석 품질 평가\n');

    % KMO 측도
    KMO = calculateKMO(responseData);
    fprintf('  - KMO 측도: %.3f', KMO);
    if KMO > 0.8
        fprintf(' (매우 적합)\n');
    elseif KMO > 0.7
        fprintf(' (적합)\n');
    else
        fprintf(' (보통 이하)\n');
    end

    % Bartlett 구형성 검정
    R = corrcoef(responseData, 'Rows', 'pairwise');
    [pBart, chi2Bart, dofBart] = bartlettSphericity(R, size(responseData, 1));

    if ~isnan(pBart) && pBart < 0.05
        fprintf('  - Bartlett 검정: 적합 (p=%.3f)\n', pBart);
    else
        fprintf('  - Bartlett 검정: 부적합\n');
    end

    % 크론바흐 알파
    highLoadingItems = abs(loadings(:, performanceFactorIdx)) > 0.4;
    if sum(highLoadingItems) >= 2
        alpha = cronbachAlpha(responseData(:, highLoadingItems));
        fprintf('  - 성과 요인 신뢰도 (α): %.3f\n', alpha);
    else
        alpha = NaN;
    end

    %% 3-8. 결과 저장
    periodResults.(sprintf('period%d', p)) = struct(...
        'loadings', loadings, ...
        'factorScores', factorScores, ...
        'performanceScores', performanceScores, ...
        'standardizedScores', standardizedScores, ...
        'percentileRanks', percentileRanks, ...
        'numFactors', optimalNumFactors, ...
        'performanceFactorIdx', performanceFactorIdx, ...
        'questionNames', {questionCols}, ...
        'responseIDs', {responseIDs}, ...
        'dataQualityFlag', dataQualityFlag, ...
        'competency_items', competencyItemTable, ...
        'scoreReliability', scoreReliability, ...
        'KMO', KMO, ...
        'bartlettP', pBart, ...
        'cronbachAlpha', alpha, ...
        'isPCA', isPCA);

    %% 3-9. 통합 테이블에 점수 추가
    colName = sprintf('Period%d_Score', p);
    colNameStd = sprintf('Period%d_StdScore', p);
    colNamePct = sprintf('Period%d_Percentile', p);
    colNameReliability = sprintf('Period%d_Reliability', p);

    consolidatedScores.(colName) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNameStd) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNamePct) = NaN(height(consolidatedScores), 1);
    consolidatedScores.(colNameReliability) = repmat({scoreReliability}, height(consolidatedScores), 1);

    % 신뢰할 수 있는 점수만 저장
    if ~strcmp(scoreReliability, 'UNUSABLE')
        for i = 1:length(responseIDs)
            idx = strcmp(consolidatedScores.ID, responseIDs{i});
            if any(idx)
                consolidatedScores.(colName)(idx) = performanceScores(i);
                consolidatedScores.(colNameStd)(idx) = standardizedScores(i);
                consolidatedScores.(colNamePct)(idx) = percentileRanks(i);
            end
        end

        fprintf('  ✓ %s 분석 완료 (신뢰도: %s)\n', periods{p}, scoreReliability);
    else
        fprintf('  ✗ %s 분석 결과 사용 불가\n', periods{p});
    end

    fprintf('\n');
end

%% 보조 함수: 빈 결과 생성
function emptyResult = createEmptyPeriodResult(reason)
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
    'dataQualityFlag', 'FAILED', ...
    'scoreReliability', 'UNUSABLE', ...
    'KMO', NaN, ...
    'bartlettP', NaN, ...
    'cronbachAlpha', NaN, ...
    'isPCA', false, ...
    'failureReason', reason);
end


%% 4. 종합 분석 및 통계 (품질 검증 반영)
fprintf('========================================\n');
fprintf('[3단계] 종합 분석 (품질 검증 포함)\n');
fprintf('========================================\n');

% 신뢰할 수 있는 시점만 선별
reliableColumns = {};
for p = 1:length(periods)
    reliabilityCol = sprintf('Period%d_Reliability', p);
    if ismember(reliabilityCol, consolidatedScores.Properties.VariableNames)
        % 해당 시점의 신뢰도 확인
        reliability = consolidatedScores.(reliabilityCol){1}; % 첫 번째 값으로 대표
        if ~strcmp(reliability, 'UNUSABLE')
            scoreCol = sprintf('Period%d_Score', p);
            reliableColumns{end+1} = scoreCol;
            fprintf('▶ %s: 신뢰도 %s - 분석 포함\n', periods{p}, reliability);
        else
            fprintf('▶ %s: 신뢰도 %s - 분석 제외\n', periods{p}, reliability);
        end
    end
end

fprintf('\n신뢰할 수 있는 시점: %d개 / %d개\n', length(reliableColumns), length(periods));

if isempty(reliableColumns)
    fprintf('[오류] 신뢰할 수 있는 시점이 없습니다. 종합 분석을 수행할 수 없습니다.\n');
    return;
end

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

% 분석 요약 통계 (품질 정보 포함)
fprintf('\n▶ 시점별 분석 요약 (품질 검증 포함)\n');
fprintf('----------------------------------------\n');
for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p))
        result = periodResults.(sprintf('period%d', p));
        fprintf('%s:\n', periods{p});
        fprintf('  - 분석 문항: %d개\n', length(result.questionNames));
        fprintf('  - 추출 요인: %d개\n', result.numFactors);
        fprintf('  - 성과 요인: %d번째\n', result.performanceFactorIdx);
        fprintf('  - 참여자: %d명\n', length(result.responseIDs));

        if isfield(result, 'KMO')
            fprintf('  - KMO: %.3f\n', result.KMO);
        end

        if isfield(result, 'dataQualityFlag')
            fprintf('  - 데이터 품질: %s\n', result.dataQualityFlag);
        end

        if isfield(result, 'scoreReliability')
            fprintf('  - 점수 신뢰도: %s\n', result.scoreReliability);
        end
        fprintf('\n');
    end
end

fprintf('▶ 통합 결과 (신뢰할 수 있는 시점 기준)\n');
fprintf('----------------------------------------\n');
fprintf('  - 전체 대상자: %d명\n', height(consolidatedScores));

if exist('consolidatedScores', 'var') && ismember('ValidReliablePeriodCount', consolidatedScores.Properties.VariableNames)
    fprintf('  - 1개 이상 신뢰할 수 있는 시점 참여: %d명 (%.1f%%)\n', ...
        sum(consolidatedScores.ValidReliablePeriodCount > 0), ...
        100*sum(consolidatedScores.ValidReliablePeriodCount > 0)/height(consolidatedScores));

    if length(reliableColumns) >= 2
        fprintf('  - 2개 이상 신뢰할 수 있는 시점 참여: %d명 (%.1f%%)\n', ...
            sum(consolidatedScores.ValidReliablePeriodCount >= 2), ...
            100*sum(consolidatedScores.ValidReliablePeriodCount >= 2)/height(consolidatedScores));
    end

    fprintf('  - 모든 신뢰할 수 있는 시점 참여: %d명 (%.1f%%)\n', ...
        sum(consolidatedScores.ValidReliablePeriodCount == length(reliableColumns)), ...
        100*sum(consolidatedScores.ValidReliablePeriodCount == length(reliableColumns))/height(consolidatedScores));
end

% 품질 문제 시점 별도 보고
fprintf('\n▶ 품질 문제 시점 보고\n');
fprintf('----------------------------------------\n');
problemPeriods = [];
for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p))
        result = periodResults.(sprintf('period%d', p));
        if isfield(result, 'scoreReliability') && strcmp(result.scoreReliability, 'UNUSABLE')
            problemPeriods{end+1} = periods{p};
            fprintf('  - %s: 데이터 품질 문제로 점수 사용 불가\n', periods{p});

            % 문제 원인 상세 진단
            if isfield(result, 'dataQualityFlag')
                switch result.dataQualityFlag
                    case 'POOR'
                        fprintf('    → 원인: 상관행렬 특이성 (행렬식 ≈ 0, 조건수 > 1e12)\n');
                        fprintf('    → 권장: 문항 수 축소, 표본 크기 확대, 또는 다른 분석 방법 고려\n');
                    case 'FAILED'
                        fprintf('    → 원인: 상관행렬 계산 실패\n');
                        fprintf('    → 권장: 데이터 전처리 재검토 또는 단순 기술통계 사용\n');
                    case 'INSUFFICIENT'
                        fprintf('    → 원인: 분석 가능한 문항 부족\n');
                        fprintf('    → 권장: 추가 문항 확보 또는 다른 시점 데이터 활용\n');
                end
            end
        end
    end
end

if isempty(problemPeriods)
    fprintf('  → 모든 시점의 데이터 품질이 양호합니다.\n');
else
    fprintf('\n  [중요] %d개 시점에서 품질 문제 발생\n', length(problemPeriods));
    fprintf('  이러한 시점의 점수는 후속 상관분석에서 자동으로 제외됩니다.\n');
end
%% 5. 성과기여도 데이터 로드 및 전처리
fprintf('\n========================================\n');
fprintf('[4단계] 성과기여도 데이터 분석\n');
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

%% 6. 성과기여도 점수 계산
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

%% 7. 시점별 성과기여도 집계 (반기별)
fprintf('\n========================================\n');
fprintf('반기별 성과기여도 집계\n');
fprintf('========================================\n\n');

% 반기별 매핑 (역량진단 시점과 맞추기)
% 23년 하반기: 23Q3, 23Q4
% 24년 상반기: 24Q1, 24Q2
% 24년 하반기: 24Q3, 24Q4
% 25년 상반기: 25Q1, 25Q2

periodMapping = {
    {'23Q3', '23Q4'},  % 23년 하반기
    {'24Q1', '24Q2'},  % 24년 상반기
    {'24Q3', '24Q4'},  % 24년 하반기
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

%% 8. 역량진단 성과점수와 성과기여도 매칭
fprintf('\n========================================\n');
fprintf('[5단계] 역량진단 vs 성과기여도 데이터 매칭\n');
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
    combinedData.Factor_Period1 = consolidatedScores.Period1_Score(competencyIdx);
    combinedData.Factor_Period2 = consolidatedScores.Period2_Score(competencyIdx);
    combinedData.Factor_Period3 = consolidatedScores.Period3_Score(competencyIdx);
    combinedData.Factor_Period4 = consolidatedScores.Period4_Score(competencyIdx);
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

%% 9. 상관분석
fprintf('\n========================================\n');
fprintf('[6단계] 역량진단 성과점수 vs 성과기여도 상관분석\n');
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

%% 10. 결과 저장
fprintf('\n========================================\n');
fprintf('[7단계] 결과 저장\n');
fprintf('========================================\n\n');

% Excel 파일로 저장
outputFileName = sprintf('competency_performance_correlation_results_%s.xlsx', datestr(now, 'yyyymmdd'));

% 통합 점수 테이블 저장
writetable(consolidatedScores, outputFileName, 'Sheet', '역량진단_통합점수');

% 각 시점별 상세 결과 저장
for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p))
        result = periodResults.(sprintf('period%d', p));

        % 개인별 점수 테이블
        periodScoreTable = table();
        periodScoreTable.ID = result.responseIDs;
        periodScoreTable.PerformanceScore = result.performanceScores;
        periodScoreTable.StandardizedScore = result.standardizedScores;
        periodScoreTable.PercentileRank = result.percentileRanks;

        sheetName = sprintf('%s_점수', periods{p});
        writetable(periodScoreTable, outputFileName, 'Sheet', sheetName);

        % 요인 부하량 테이블
        loadingTable = table();
        loadingTable.Question = result.questionNames';
        for f = 1:result.numFactors
            loadingTable.(sprintf('Factor%d', f)) = result.loadings(:, f);
        end

        sheetName = sprintf('%s_부하량', periods{p});
        writetable(loadingTable, outputFileName, 'Sheet', sheetName);
    end
end

% 성과기여도 데이터 저장
writetable(contributionByPeriod, outputFileName, 'Sheet', '성과기여도점수');

% 통합 상관분석 데이터 저장
writetable(combinedData, outputFileName, 'Sheet', '상관분석_통합데이터');

% 상관분석 결과 테이블 생성
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

% 분석 요약 테이블
summaryTable = table();
summaryTable.Period = [periods'; {'전체'}];
summaryTable.Participants = zeros(length(periods)+1, 1);
summaryTable.Questions = zeros(length(periods)+1, 1);
summaryTable.Factors = zeros(length(periods)+1, 1);
summaryTable.PerformanceFactor = zeros(length(periods)+1, 1);
summaryTable.KMO = zeros(length(periods)+1, 1);
summaryTable.CronbachAlpha = zeros(length(periods)+1, 1);

for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p))
        result = periodResults.(sprintf('period%d', p));
        summaryTable.Participants(p) = length(result.responseIDs);
        summaryTable.Questions(p) = length(result.questionNames);
        summaryTable.Factors(p) = result.numFactors;
        summaryTable.PerformanceFactor(p) = result.performanceFactorIdx;
        if isfield(result, 'KMO')
            summaryTable.KMO(p) = result.KMO;
        end
        if isfield(result, 'cronbachAlpha')
            summaryTable.CronbachAlpha(p) = result.cronbachAlpha;
        end
    end
end

% 전체 행 통계
summaryTable.Participants(end) = sum(consolidatedScores.ValidPeriodCount > 0);
summaryTable.Questions(end) = NaN;
summaryTable.Factors(end) = NaN;
summaryTable.PerformanceFactor(end) = NaN;
summaryTable.KMO(end) = mean(summaryTable.KMO(1:end-1), 'omitnan');
summaryTable.CronbachAlpha(end) = mean(summaryTable.CronbachAlpha(1:end-1), 'omitnan');

writetable(summaryTable, outputFileName, 'Sheet', '분석요약');

% MAT 파일 저장
matFileName = sprintf('competency_correlation_workspace_%s.mat', datestr(now, 'yyyymmdd'));
save(matFileName, 'allData', 'periodResults', 'consolidatedScores', 'contributionByPeriod', ...
    'combinedData', 'correlationResults', 'periods');

fprintf('▶ 결과 저장 완료\n');
fprintf('  - Excel 파일: %s\n', outputFileName);
fprintf('  - MAT 파일: %s\n', matFileName);

%% 11. 최종 보고
fprintf('\n========================================\n');
fprintf('분석 완료 요약\n');
fprintf('========================================\n');

fprintf('📊 데이터 현황\n');
fprintf('  • 분석 기간: %s ~ %s\n', periods{1}, periods{end});
fprintf('  • 총 대상자: %d명\n', height(consolidatedScores));
fprintf('  • 실제 참여자: %d명 (%.1f%%)\n', ...
    sum(consolidatedScores.ValidPeriodCount > 0), ...
    100*sum(consolidatedScores.ValidPeriodCount > 0)/height(consolidatedScores));

fprintf('\n🔍 시점별 요인분석 결과\n');
for p = 1:length(periods)
    if isfield(periodResults, sprintf('period%d', p))
        result = periodResults.(sprintf('period%d', p));
        fprintf('  • %s: %d개 요인, 성과요인=%d번째, KMO=%.3f\n', ...
            periods{p}, result.numFactors, result.performanceFactorIdx, ...
            result.KMO);
    end
end

fprintf('\n📈 성과점수 현황\n');
fprintf('  • 평균 표준화 점수 보유자: %d명 (%.1f%%)\n', ...
    sum(~isnan(consolidatedScores.AverageStdScore)), ...
    100*sum(~isnan(consolidatedScores.AverageStdScore))/height(consolidatedScores));

if sum(~isnan(consolidatedScores.AverageStdScore)) > 0
    fprintf('  • 평균 점수: %.3f (±%.3f)\n', ...
        mean(consolidatedScores.AverageStdScore, 'omitnan'), ...
        std(consolidatedScores.AverageStdScore, 'omitnan'));
end

fprintf('\n🔗 상관분석 결과\n');
fprintf('  • 매칭된 데이터: %d명\n', height(combinedData));
for p = 1:4
    if isfield(correlationResults, sprintf('period%d', p))
        result = correlationResults.(sprintf('period%d', p));
        if ~isnan(result.correlation)
            sig_str = '';
            if result.p_value < 0.001, sig_str = '***';
            elseif result.p_value < 0.01, sig_str = '**';
            elseif result.p_value < 0.05, sig_str = '*';
            end
            fprintf('  • %s: r=%.3f (n=%d) %s\n', periods{p}, result.correlation, result.n, sig_str);
        end
    end
end

if isfield(correlationResults, 'overall')
    result = correlationResults.overall;
    sig_str = '';
    if result.p_value < 0.001, sig_str = '***';
    elseif result.p_value < 0.01, sig_str = '**';
    elseif result.p_value < 0.05, sig_str = '*';
    end
    fprintf('  • 전체 평균: r=%.3f (n=%d) %s\n', result.correlation, result.n, sig_str);
end

fprintf('\n✅ 주요 특징:\n');
fprintf('  • 각 시점의 고유한 요인구조 파악\n');
fprintf('  • 시점별 문항 차이를 반영한 분석\n');
fprintf('  • 성과기여도와의 관계성 검증\n');
fprintf('  • 시간에 따른 변화 추적 가능\n');

fprintf('\n분석이 성공적으로 완료되었습니다!\n');


%% 10. 성장단계 변화량 분석
fprintf('\n========================================\n');
fprintf('[8단계] 성장단계 변화량 분석\n');
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

%% 10-1. 성장단계 점수 매핑 함수 정의

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

% function score = convertGrowthStageToScore(stageText)
%     % 성장단계를 숫자 점수로 변환
%     % 열린-성취-책임-소명 순으로 높은 등급
%     % 레벨이 높을수록 세부 등급이 높음
%
%     if isempty(stageText) || ismissing(stageText) || (iscell(stageText) && isempty(stageText{1}))
%         score = NaN;
%         return;
%     end
%
%     if iscell(stageText)
%         stageText = stageText{1};
%     end
%     stageText = char(stageText);
%
%     % 기본 단계별 점수 (소명 > 책임 > 성취 > 열린)
%     baseScores = containers.Map({'열린', '성취', '책임', '소명'}, {10, 20, 30, 40});
%
%     % 레벨 점수 추가 (Lv1=1, Lv2=2, Lv3=3, ...)
%     score = 0;
%
%     % 단계 확인
%     for stage = keys(baseScores)
%         if contains(stageText, stage{1})
%             score = score + baseScores(stage{1});
%             break;
%         end
%     end
%
%     % 레벨 확인
%     levelMatch = regexp(stageText, 'Lv(\d+)', 'tokens');
%     if ~isempty(levelMatch)
%         level = str2double(levelMatch{1}{1});
%         score = score + level;
%     end
%
%     % 추가 숫자 확인 (예: "열린(Lv2)-3"에서 -3 부분)
%     extraMatch = regexp(stageText, '-(\d+)', 'tokens');
%     if ~isempty(extraMatch)
%         extra = str2double(extraMatch{1}{1});
%         score = score + extra * 0.1; % 소수점으로 추가
%     end
%
%     if score == 0
%         score = NaN;
%     end

%% 10-2. 반기별 발현역량 점수 계산
fprintf('▶ 반기별 발현역량 점수 계산\n');

% 반기별 발현역량 컬럼 확인
competencyColumns = {'23H1 발현역량', '23H2 발현역량', '24H1 발현역량', '24H2 발현역량', '25H1 발현역량'};
growthCompetencyScores = table();
growthCompetencyScores.ID = growthData.ID;

% 각 반기별 점수 계산
for col = 1:length(competencyColumns)
    colName = competencyColumns{col};
    if any(strcmp(growthData.Properties.VariableNames, colName))
        competencyData = growthData.(colName);
        scores = zeros(height(growthData), 1);

        for i = 1:length(competencyData)
            scores(i) = convertGrowthStageToScore(competencyData(i));
        end

        scoreColName = sprintf('Competency_%s', strrep(colName, ' 발현역량', ''));
        scoreColName = strrep(scoreColName, 'H', '_H');
        growthCompetencyScores.(scoreColName) = scores;

        validCount = sum(~isnan(scores));
        fprintf('  %s: %d명 유효 (평균: %.2f)\n', colName, validCount, nanmean(scores));
    end
end

%% 10-3. 연도별 성장단계 점수 계산
fprintf('\n▶ 연도별 성장단계 점수 계산\n');

% 연도별 성장단계 컬럼 확인
stageColumns = {'24 성장단계', '25 성장단계'};

for col = 1:length(stageColumns)
    colName = stageColumns{col};
    if any(strcmp(growthData.Properties.VariableNames, colName))
        stageData = growthData.(colName);
        scores = zeros(height(growthData), 1);

        for i = 1:length(stageData)
            scores(i) = convertGrowthStageToScore(stageData(i));
        end

        scoreColName = sprintf('Stage_%s', strrep(colName, ' 성장단계', ''));
        growthCompetencyScores.(scoreColName) = scores;

        validCount = sum(~isnan(scores));
        fprintf('  %s: %d명 유효 (평균: %.2f)\n', colName, validCount, nanmean(scores));
    end
end

%% 10-4. 성장단계 변화량 계산
fprintf('\n▶ 성장단계 변화량 계산\n');

% 1) 반기별 발현역량 변화량 계산
competencyChangeTable = table();
competencyChangeTable.ID = growthCompetencyScores.ID;

% 연속된 반기 간 변화량
competencyPeriods = {'23_H1', '23_H2', '24_H1', '24_H2', '25_H1'};
for i = 1:length(competencyPeriods)-1
    currentCol = sprintf('Competency_%s', competencyPeriods{i});
    nextCol = sprintf('Competency_%s', competencyPeriods{i+1});

    if ismember(currentCol, growthCompetencyScores.Properties.VariableNames) && ...
            ismember(nextCol, growthCompetencyScores.Properties.VariableNames)

        currentScores = growthCompetencyScores.(currentCol);
        nextScores = growthCompetencyScores.(nextCol);

        % 변화량 = 다음 기간 점수 - 현재 기간 점수
        change = nextScores - currentScores;
        changeColName = sprintf('Change_%s_to_%s', competencyPeriods{i}, competencyPeriods{i+1});
        competencyChangeTable.(changeColName) = change;

        validCount = sum(~isnan(change));
        fprintf('  %s → %s: %d명 유효, 평균 변화량: %.3f\n', ...
            competencyPeriods{i}, competencyPeriods{i+1}, validCount, nanmean(change));
    end
end

% 전체 기간 변화량 (23H1 → 25H1)
if ismember('Competency_23_H1', growthCompetencyScores.Properties.VariableNames) && ...
        ismember('Competency_25_H1', growthCompetencyScores.Properties.VariableNames)

    overallChange = growthCompetencyScores.Competency_25_H1 - growthCompetencyScores.Competency_23_H1;
    competencyChangeTable.Overall_Change_23H1_to_25H1 = overallChange;

    validCount = sum(~isnan(overallChange));
    fprintf('  전체 변화량 (23H1→25H1): %d명 유효, 평균: %.3f\n', validCount, nanmean(overallChange));
end

% 2) 연도별 성장단계 변화량 계산
if ismember('Stage_24', growthCompetencyScores.Properties.VariableNames) && ...
        ismember('Stage_25', growthCompetencyScores.Properties.VariableNames)

    stageChange = growthCompetencyScores.Stage_25 - growthCompetencyScores.Stage_24;
    competencyChangeTable.Stage_Change_24_to_25 = stageChange;

    validCount = sum(~isnan(stageChange));
    fprintf('  성장단계 변화량 (24→25): %d명 유효, 평균: %.3f\n', validCount, nanmean(stageChange));
end

% 3) 변화 패턴 분석
fprintf('\n▶ 변화 패턴 분석\n');

% 전체 변화량 기준 분류
if ismember('Overall_Change_23H1_to_25H1', competencyChangeTable.Properties.VariableNames)
    overallChanges = competencyChangeTable.Overall_Change_23H1_to_25H1;
    validChanges = overallChanges(~isnan(overallChanges));

    if length(validChanges) > 0
        % 변화 방향별 분류
        improved = sum(validChanges > 0);
        maintained = sum(validChanges == 0);
        declined = sum(validChanges < 0);

        fprintf('  성장한 인원: %d명 (%.1f%%)\n', improved, 100*improved/length(validChanges));
        fprintf('  유지된 인원: %d명 (%.1f%%)\n', maintained, 100*maintained/length(validChanges));
        fprintf('  하락한 인원: %d명 (%.1f%%)\n', declined, 100*declined/length(validChanges));

        % 변화량 크기별 분류
        largeImprovement = sum(validChanges >= 1.0);
        smallImprovement = sum(validChanges > 0 & validChanges < 1.0);
        smallDecline = sum(validChanges < 0 & validChanges > -1.0);
        largeDecline = sum(validChanges <= -1.0);

        fprintf('  큰 폭 성장 (≥1.0): %d명 (%.1f%%)\n', largeImprovement, 100*largeImprovement/length(validChanges));
        fprintf('  소폭 성장 (0~1.0): %d명 (%.1f%%)\n', smallImprovement, 100*smallImprovement/length(validChanges));
        fprintf('  소폭 하락 (-1.0~0): %d명 (%.1f%%)\n', smallDecline, 100*smallDecline/length(validChanges));
        fprintf('  큰 폭 하락 (≤-1.0): %d명 (%.1f%%)\n', largeDecline, 100*largeDecline/length(validChanges));
    end
end

%% 10-5. 역량진단 성과점수와 성장단계 변화량 매칭
fprintf('\n========================================\n');
fprintf('[9단계] 성장단계 변화량 vs 역량진단 성과점수 상관분석\n');
fprintf('========================================\n\n');

% ID를 기준으로 데이터 매칭
% 기존 combinedData (역량진단+성과기여도)에 성장단계 변화량 추가

% ID를 문자열로 통일
if isnumeric(competencyChangeTable.ID)
    growthIDs = arrayfun(@num2str, competencyChangeTable.ID, 'UniformOutput', false);
else
    growthIDs = cellfun(@char, competencyChangeTable.ID, 'UniformOutput', false);
end

if isnumeric(combinedData.ID)
    combinedIDs = arrayfun(@num2str, combinedData.ID, 'UniformOutput', false);
else
    combinedIDs = cellfun(@char, combinedData.ID, 'UniformOutput', false);
end

% 교집합 찾기
[commonGrowthIDs, combinedGrowthIdx, growthIdx] = intersect(combinedIDs, growthIDs);

fprintf('성장단계 데이터 매칭 결과:\n');
fprintf('  - 역량진단+성과기여도 데이터: %d명\n', height(combinedData));
fprintf('  - 성장단계 변화량 데이터: %d명\n', height(competencyChangeTable));
fprintf('  - 공통 ID: %d명\n', length(commonGrowthIDs));
fprintf('  - 매칭률: %.1f%%\n', 100 * length(commonGrowthIDs) / min(height(combinedData), height(competencyChangeTable)));

% 매칭된 데이터로 확장된 통합 테이블 생성
if length(commonGrowthIDs) > 0
    % 기존 combinedData에 성장단계 변화량 컬럼 추가
    extendedCombinedData = combinedData;

    % 성장단계 변화량 컬럼들을 NaN으로 초기화
    changeColumns = competencyChangeTable.Properties.VariableNames(2:end); % ID 제외
    for col = 1:length(changeColumns)
        extendedCombinedData.(changeColumns{col}) = NaN(height(combinedData), 1);
    end

    % 매칭된 데이터만 성장단계 변화량 값 할당
    for col = 1:length(changeColumns)
        colName = changeColumns{col};
        if ismember(colName, competencyChangeTable.Properties.VariableNames)
            extendedCombinedData.(colName)(combinedGrowthIdx) = ...
                competencyChangeTable.(colName)(growthIdx);
        end
    end

    fprintf('확장된 통합 데이터 생성 완료: %d명\n', height(extendedCombinedData));

    % 주요 변화량 변수들
    keyChangeVars = {};
    if ismember('Overall_Change_23H1_to_25H1', changeColumns)
        keyChangeVars{end+1} = 'Overall_Change_23H1_to_25H1';
    end
    if ismember('Stage_Change_24_to_25', changeColumns)
        keyChangeVars{end+1} = 'Stage_Change_24_to_25';
    end

    % 최근 변화량도 추가 (가장 최근 반기 변화)
    recentChangeVars = {};
    for col = 1:length(changeColumns)
        if contains(changeColumns{col}, '24_H2_to_25_H1') || contains(changeColumns{col}, '24H2_to_25H1')
            recentChangeVars{end+1} = changeColumns{col};
        end
    end
    keyChangeVars = [keyChangeVars, recentChangeVars];

else
    fprintf('[오류] 매칭된 데이터가 없습니다. 성장단계 분석을 계속할 수 없습니다.\n');
    return;
end

%% 10-6. 성장단계 변화량과 역량진단 성과점수 상관분석
fprintf('\n========================================\n');
fprintf('성장단계 변화량 vs 역량진단 성과점수 상관분석\n');
fprintf('========================================\n\n');

growthCorrelationResults = struct();

% 1) 전체 평균 성과점수와 주요 변화량들 간 상관분석
fprintf('[주요 변화량과 전체 평균 성과점수 상관분석]\n');

for i = 1:length(keyChangeVars)
    changeVar = keyChangeVars{i};

    if ismember(changeVar, extendedCombinedData.Properties.VariableNames)
        factorScores = extendedCombinedData.Factor_Average;
        changeScores = extendedCombinedData.(changeVar);

        % 둘 다 유효한 값이 있는 경우만 분석
        validIdx = ~isnan(factorScores) & ~isnan(changeScores);
        validCount = sum(validIdx);

        if validCount >= 5
            r = corrcoef(factorScores(validIdx), changeScores(validIdx));
            correlation = r(1, 2);

            % 유의성 검정
            t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
            p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

            fprintf('%s: r = %.3f (n=%d, p=%.3f)', changeVar, correlation, validCount, p_value);
            if p_value < 0.001
                fprintf(' ***');
            elseif p_value < 0.01
                fprintf(' **');
            elseif p_value < 0.05
                fprintf(' *');
            end
            fprintf('\n');

            growthCorrelationResults.(changeVar) = struct(...
                'correlation', correlation, ...
                'n', validCount, ...
                'p_value', p_value, ...
                'factor_type', 'overall');

        else
            fprintf('%s: 분석 불가 (유효 데이터 %d개)\n', changeVar, validCount);
        end
    end
end

% 2) 시점별 성과점수와 변화량 간 상관분석
fprintf('\n[시점별 성과점수와 변화량 상관분석]\n');

% 전체 변화량과 각 시점별 점수 상관
if ismember('Overall_Change_23H1_to_25H1', keyChangeVars)
    changeVar = 'Overall_Change_23H1_to_25H1';
    changeScores = extendedCombinedData.(changeVar);

    for p = 1:4
        factorCol = sprintf('Factor_Period%d', p);
        if ismember(factorCol, extendedCombinedData.Properties.VariableNames)
            factorScores = extendedCombinedData.(factorCol);

            validIdx = ~isnan(factorScores) & ~isnan(changeScores);
            validCount = sum(validIdx);

            if validCount >= 5
                r = corrcoef(factorScores(validIdx), changeScores(validIdx));
                correlation = r(1, 2);

                t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
                p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

                fprintf('%s vs %s: r = %.3f (n=%d, p=%.3f)', ...
                    periods{p}, changeVar, correlation, validCount, p_value);
                if p_value < 0.001
                    fprintf(' ***');
                elseif p_value < 0.01
                    fprintf(' **');
                elseif p_value < 0.05
                    fprintf(' *');
                end
                fprintf('\n');

                growthCorrelationResults.(sprintf('%s_vs_%s', factorCol, changeVar)) = struct(...
                    'correlation', correlation, ...
                    'n', validCount, ...
                    'p_value', p_value, ...
                    'factor_type', sprintf('period_%d', p));
            end
        end
    end
end

% 3) 성과기여도와 성장단계 변화량 상관분석
fprintf('\n[성과기여도와 성장단계 변화량 상관분석]\n');

for i = 1:length(keyChangeVars)
    changeVar = keyChangeVars{i};

    if ismember(changeVar, extendedCombinedData.Properties.VariableNames)
        contribScores = extendedCombinedData.Contribution_Average;
        changeScores = extendedCombinedData.(changeVar);

        validIdx = ~isnan(contribScores) & ~isnan(changeScores);
        validCount = sum(validIdx);

        if validCount >= 5
            r = corrcoef(contribScores(validIdx), changeScores(validIdx));
            correlation = r(1, 2);

            t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
            p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

            fprintf('성과기여도 vs %s: r = %.3f (n=%d, p=%.3f)', ...
                changeVar, correlation, validCount, p_value);
            if p_value < 0.001
                fprintf(' ***');
            elseif p_value < 0.01
                fprintf(' **');
            elseif p_value < 0.05
                fprintf(' *');
            end
            fprintf('\n');

            growthCorrelationResults.(sprintf('contribution_vs_%s', changeVar)) = struct(...
                'correlation', correlation, ...
                'n', validCount, ...
                'p_value', p_value, ...
                'factor_type', 'contribution');
        end
    end
end

%% 10-7. 결과 저장 (성장단계 관련)
fprintf('\n========================================\n');
fprintf('성장단계 분석 결과 저장\n');
fprintf('========================================\n\n');

% 기존 Excel 파일에 시트 추가
if exist('outputFileName', 'var')
    % 성장단계 점수 데이터 저장
    writetable(growthCompetencyScores, outputFileName, 'Sheet', '성장단계점수');

    % 성장단계 변화량 데이터 저장
    writetable(competencyChangeTable, outputFileName, 'Sheet', '성장단계변화량');

    % 확장된 통합 데이터 저장
    writetable(extendedCombinedData, outputFileName, 'Sheet', '통합데이터_성장단계포함');

    % 성장단계 상관분석 결과 테이블 생성
    growthCorrFields = fieldnames(growthCorrelationResults);
    if ~isempty(growthCorrFields)
        growthCorrTable = table();
        growthCorrTable.Analysis_Type = growthCorrFields;
        growthCorrTable.Correlation = NaN(length(growthCorrFields), 1);
        growthCorrTable.N = NaN(length(growthCorrFields), 1);
        growthCorrTable.P_Value = NaN(length(growthCorrFields), 1);
        growthCorrTable.Significance = cell(length(growthCorrFields), 1);

        for i = 1:length(growthCorrFields)
            field = growthCorrFields{i};
            result = growthCorrelationResults.(field);

            growthCorrTable.Correlation(i) = result.correlation;
            growthCorrTable.N(i) = result.n;
            growthCorrTable.P_Value(i) = result.p_value;

            if result.p_value < 0.001
                growthCorrTable.Significance{i} = '***';
            elseif result.p_value < 0.01
                growthCorrTable.Significance{i} = '**';
            elseif result.p_value < 0.05
                growthCorrTable.Significance{i} = '*';
            else
                growthCorrTable.Significance{i} = 'n.s.';
            end
        end

        writetable(growthCorrTable, outputFileName, 'Sheet', '성장단계_상관분석결과');
    end

    fprintf('▶ 성장단계 분석 결과 저장 완료\n');
    fprintf('  - 성장단계점수 시트\n');
    fprintf('  - 성장단계변화량 시트\n');
    fprintf('  - 통합데이터_성장단계포함 시트\n');
    fprintf('  - 성장단계_상관분석결과 시트\n');
end

%% 10-8. 성장단계 분석 최종 요약
fprintf('\n========================================\n');
fprintf('성장단계 분석 최종 요약\n');
fprintf('========================================\n');

fprintf('📈 성장단계 변화 현황\n');
if ismember('Overall_Change_23H1_to_25H1', competencyChangeTable.Properties.VariableNames)
    overallChanges = competencyChangeTable.Overall_Change_23H1_to_25H1;
    validChanges = overallChanges(~isnan(overallChanges));

    if length(validChanges) > 0
        fprintf('  • 전체 변화량 분석 대상: %d명\n', length(validChanges));
        fprintf('  • 평균 변화량: %.3f (±%.3f)\n', mean(validChanges), std(validChanges));
        fprintf('  • 성장한 인원: %d명 (%.1f%%)\n', sum(validChanges > 0), ...
            100*sum(validChanges > 0)/length(validChanges));
        fprintf('  • 하락한 인원: %d명 (%.1f%%)\n', sum(validChanges < 0), ...
            100*sum(validChanges < 0)/length(validChanges));
    end
end

fprintf('\n🔗 주요 상관분석 결과\n');
fprintf('  • 성장단계-역량진단 매칭: %d명\n', length(commonGrowthIDs));

% 가장 높은 상관계수 찾기
maxCorr = -1;
maxCorrVar = '';
for i = 1:length(growthCorrFields)
    field = growthCorrFields{i};
    corr = abs(growthCorrelationResults.(field).correlation);
    if corr > maxCorr && ~isnan(corr)
        maxCorr = corr;
        maxCorrVar = field;
    end
end

if ~isempty(maxCorrVar)
    result = growthCorrelationResults.(maxCorrVar);
    sig_str = '';
    if result.p_value < 0.001, sig_str = '***';
    elseif result.p_value < 0.01, sig_str = '**';
    elseif result.p_value < 0.05, sig_str = '*';
    end
    fprintf('  • 최고 상관계수: %s (r=%.3f) %s\n', maxCorrVar, result.correlation, sig_str);
end

fprintf('\n✅ 성장단계 분석 완료!\n');
fprintf('  • 반기별 발현역량 점수 산출\n');
fprintf('  • 연도별 성장단계 점수 산출\n');
fprintf('  • 다양한 변화량 지표 계산\n');
fprintf('  • 역량진단-성장단계 상관분석\n');
fprintf('  • 성과기여도-성장단계 상관분석\n');

...
    %% 11. 확장된 CSR + 조직지표 분석 (Q40-46)
%clc
fprintf('\n========================================\n');
fprintf('[11단계] 확장된 CSR + 조직지표 분석 (25년 상반기)\n');
fprintf('========================================\n\n');

% 25년 상반기(period 4)에서 제외된 Q40-46 데이터 추출
extendedCSRResults = struct();

if isfield(allData, 'period4') && isfield(allData.period4, 'originalSelfData')
    fprintf('▶ 25년 상반기 원본 데이터에서 확장 문항 추출\n');

    originalData = allData.period4.originalSelfData;

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
        return;
    end

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
        %% 영역별 점수 계산
        fprintf('\n▶ 영역별 점수 계산\n');

        % 1. 조직 기반 지표 (Q40, Q41)
        orgCols = {'OrganizationalSynergy', 'Pride'};
        validOrgCols = {};
        for col = 1:length(orgCols)
            if ismember(orgCols{col}, extendedData.Properties.VariableNames)
                validOrgCols{end+1} = orgCols{col};
            end
        end

        if ~isempty(validOrgCols)
            orgMatrix = table2array(extendedData(:, validOrgCols));
            extendedData.Organizational_Score = mean(orgMatrix, 2, 'omitnan');
            validOrg = sum(~isnan(extendedData.Organizational_Score));
            fprintf('  조직 기반 점수 (Q40,41): %d명 (평균: %.3f)\n', validOrg, nanmean(extendedData.Organizational_Score));

            % 개별 지표도 보고
            if ismember('OrganizationalSynergy', validOrgCols)
                validSynergy = sum(~isnan(extendedData.OrganizationalSynergy));
                fprintf('    └ 조직시너지 (Q40): %d명 (평균: %.3f)\n', validSynergy, nanmean(extendedData.OrganizationalSynergy));
            end
            if ismember('Pride', validOrgCols)
                validPride = sum(~isnan(extendedData.Pride));
                fprintf('    └ 자부심 (Q41): %d명 (평균: %.3f)\n', validPride, nanmean(extendedData.Pride));
            end
        end

        % 2. Communication 점수 (Q42, Q43)
        commCols = {'Communication_Relationship', 'Communication_Purpose'};
        validCommCols = {};
        for col = 1:length(commCols)
            if ismember(commCols{col}, extendedData.Properties.VariableNames)
                validCommCols{end+1} = commCols{col};
            end
        end

        if ~isempty(validCommCols)
            commMatrix = table2array(extendedData(:, validCommCols));
            extendedData.Communication_Score = mean(commMatrix, 2, 'omitnan');
            validComm = sum(~isnan(extendedData.Communication_Score));
            fprintf('  Communication 점수 (Q42,43): %d명 (평균: %.3f)\n', validComm, nanmean(extendedData.Communication_Score));
        end

        % 3. Strategy 점수 (Q44, Q45)
        stratCols = {'Strategy_CustomerValue', 'Strategy_Performance'};
        validStratCols = {};
        for col = 1:length(stratCols)
            if ismember(stratCols{col}, extendedData.Properties.VariableNames)
                validStratCols{end+1} = stratCols{col};
            end
        end

        if ~isempty(validStratCols)
            stratMatrix = table2array(extendedData(:, validStratCols));
            extendedData.Strategy_Score = mean(stratMatrix, 2, 'omitnan');
            validStrat = sum(~isnan(extendedData.Strategy_Score));
            fprintf('  Strategy 점수 (Q44,45): %d명 (평균: %.3f)\n', validStrat, nanmean(extendedData.Strategy_Score));
        end

        % 4. Reflection 점수 (Q46)
        if ismember('Reflection_Organizational', extendedData.Properties.VariableNames)
            extendedData.Reflection_Score = extendedData.Reflection_Organizational;
            validRefl = sum(~isnan(extendedData.Reflection_Score));
            fprintf('  Reflection 점수 (Q46): %d명 (평균: %.3f)\n', validRefl, nanmean(extendedData.Reflection_Score));
        end

        % 5. 전체 CSR 점수 (Q42-46, Communication+Strategy+Reflection)
        csrScoreCols = {'Communication_Score', 'Strategy_Score', 'Reflection_Score'};
        validCSRCols = {};
        for col = 1:length(csrScoreCols)
            if ismember(csrScoreCols{col}, extendedData.Properties.VariableNames)
                validCSRCols{end+1} = csrScoreCols{col};
            end
        end

        if ~isempty(validCSRCols)
            csrMatrix = table2array(extendedData(:, validCSRCols));
            extendedData.Total_CSR_Score = mean(csrMatrix, 2, 'omitnan');
            validTotal = sum(~isnan(extendedData.Total_CSR_Score));
            fprintf('  전체 CSR 점수 (Q42-46): %d명 (평균: %.3f)\n', validTotal, nanmean(extendedData.Total_CSR_Score));
        end

        % 6. 통합 리더십 점수 (Q40-46 전체)
        allScoreCols = {'Organizational_Score', 'Communication_Score', 'Strategy_Score', 'Reflection_Score'};
        validAllCols = {};
        for col = 1:length(allScoreCols)
            if ismember(allScoreCols{col}, extendedData.Properties.VariableNames)
                validAllCols{end+1} = allScoreCols{col};
            end
        end

        if ~isempty(validAllCols)
            allMatrix = table2array(extendedData(:, validAllCols));
            extendedData.Total_Leadership_Score = mean(allMatrix, 2, 'omitnan');
            validAll = sum(~isnan(extendedData.Total_Leadership_Score));
            fprintf('  통합 리더십 점수 (Q40-46): %d명 (평균: %.3f)\n', validAll, nanmean(extendedData.Total_Leadership_Score));
        end

        % 결과 저장
        extendedCSRResults.extendedData = extendedData;
        extendedCSRResults.foundQuestions = foundQuestions;
        extendedCSRResults.missingQuestions = missingQuestions;

    else
        fprintf('[경고] 확장 문항을 찾을 수 없습니다.\n');
        return;
    end

else
    fprintf('[경고] 25년 상반기 원본 데이터를 찾을 수 없습니다.\n');
    return;
end

%% 12. 확장된 상관분석 (조직지표 + CSR vs 성과)
fprintf('\n========================================\n');
fprintf('[12단계] 확장된 상관분석 (조직지표+CSR vs 성과)\n');
fprintf('========================================\n\n');

% 기존 통합 데이터와 확장 데이터 매칭
if exist('combinedData', 'var') && istable(combinedData) && exist('extendedData', 'var') && istable(extendedData)

    % ID를 문자열로 통일
    if isnumeric(extendedData.ID)
        extendedIDs = arrayfun(@num2str, extendedData.ID, 'UniformOutput', false);
    else
        extendedIDs = cellfun(@char, extendedData.ID, 'UniformOutput', false);
    end

    if isnumeric(combinedData.ID)
        combinedIDs = arrayfun(@num2str, combinedData.ID, 'UniformOutput', false);
    else
        combinedIDs = cellfun(@char, combinedData.ID, 'UniformOutput', false);
    end

    % 교집합 찾기
    [commonExtendedIDs, combinedExtendedIdx, extendedIdx] = intersect(combinedIDs, extendedIDs);

    fprintf('확장 데이터 매칭 결과:\n');
    fprintf('  - 역량진단+성과기여도 데이터: %d명\n', height(combinedData));
    fprintf('  - 확장 문항 데이터 (Q40-46): %d명\n', height(extendedData));
    fprintf('  - 공통 ID: %d명\n', length(commonExtendedIDs));
    fprintf('  - 매칭률: %.1f%%\n', 100 * length(commonExtendedIDs) / min(height(combinedData), height(extendedData)));

    if length(commonExtendedIDs) >= 5

        %% 통합 데이터 생성
        fprintf('\n▶ 확장 통합 데이터 생성\n');

        extendedCombinedData = combinedData;  % 전체 combinedData 복사

        % 확장 점수 컬럼들 정의 - 모두 영문으로 통일
        extendedScoreColumns = {
            'OrganizationalSynergy', 'Pride', 'Organizational_Score', ...
            'Communication_Relationship', 'Communication_Purpose', 'Communication_Score', ...
            'Strategy_CustomerValue', 'Strategy_Performance', 'Strategy_Score', ...
            'Reflection_Organizational', 'Reflection_Score', ...
            'Total_CSR_Score', 'Total_Leadership_Score'
            };

        % 실제로 존재하는 확장 점수 컬럼만 선택
        existingExtendedCols = {};
        for i = 1:length(extendedScoreColumns)
            colName = extendedScoreColumns{i};
            if ismember(colName, extendedData.Properties.VariableNames)
                existingExtendedCols{end+1} = colName;
                extendedCombinedData.(colName) = NaN(height(combinedData), 1);  % NaN으로 초기화
            end
        end

        fprintf('  발견된 확장 점수 컬럼: %d개\n', length(existingExtendedCols));
        for i = 1:length(existingExtendedCols)
            fprintf('    - %s\n', existingExtendedCols{i});
        end

        % 매칭된 데이터만 확장 점수 할당
        for i = 1:length(existingExtendedCols)
            colName = existingExtendedCols{i};
            extendedCombinedData.(colName)(combinedExtendedIdx) = extendedData.(colName)(extendedIdx);
        end

        fprintf('\n확장 통합 데이터 생성 완료: %d명\n', height(extendedCombinedData));

        % 통합 결과 검증
        for i = 1:length(existingExtendedCols)
            colName = existingExtendedCols{i};
            validCount = sum(~isnan(extendedCombinedData.(colName)));
            if validCount > 0
                meanScore = nanmean(extendedCombinedData.(colName));
                fprintf('  %s: %d명 유효 데이터 (평균: %.3f)\n', colName, validCount, meanScore);
            end
        end

        %% 상관분석 실행
        fprintf('\n▶ 확장된 상관분석 실행\n');

        extendedCorrelationResults = struct();

        % 분석할 성과 변수들
        performanceVars = {};
        if ismember('Factor_Average', extendedCombinedData.Properties.VariableNames)
            performanceVars{end+1} = 'Factor_Average';
        end
        if ismember('Factor_Period2', extendedCombinedData.Properties.VariableNames)
            performanceVars{end+1} = 'Factor_Period2';
        end
        if ismember('Factor_Period4', extendedCombinedData.Properties.VariableNames)
            performanceVars{end+1} = 'Factor_Period4';
        end

        % 성과기여도 변수들
        contributionVars = {};
        if ismember('Contribution_Average', extendedCombinedData.Properties.VariableNames)
            contributionVars{end+1} = 'Contribution_Average';
        end
        if ismember('Contribution_Period2', extendedCombinedData.Properties.VariableNames)
            contributionVars{end+1} = 'Contribution_Period2';
        end
        if ismember('Contribution_Period4', extendedCombinedData.Properties.VariableNames)
            contributionVars{end+1} = 'Contribution_Period4';
        end

        allPerformanceVars = [performanceVars, contributionVars];

        fprintf('분석할 성과 변수: %d개\n', length(allPerformanceVars));

        % 주요 확장 지표별 상관분석
        keyExtendedVars = {
            'OrganizationalSynergy', 'Pride', 'Organizational_Score', ...
            'Communication_Score', 'Strategy_Score', 'Reflection_Score', ...
            'Total_CSR_Score', 'Total_Leadership_Score'
            };
        analysisKeyVars = {};
        for i = 1:length(keyExtendedVars)
            if ismember(keyExtendedVars{i}, existingExtendedCols)
                analysisKeyVars{end+1} = keyExtendedVars{i};
            end
        end

        fprintf('주요 분석 지표: %d개\n', length(analysisKeyVars));

        for perfIdx = 1:length(allPerformanceVars)
            perfVar = allPerformanceVars{perfIdx};

            fprintf('\n[%s와 확장 지표 상관분석]\n', perfVar);

            for extIdx = 1:length(analysisKeyVars)
                extVar = analysisKeyVars{extIdx};

                perfScores = extendedCombinedData.(perfVar);
                extScores = extendedCombinedData.(extVar);

                % 유효한 데이터만 분석
                validIdx = ~isnan(perfScores) & ~isnan(extScores);
                validCount = sum(validIdx);

                if validCount >= 5
                    r = corrcoef(perfScores(validIdx), extScores(validIdx));
                    correlation = r(1, 2);

                    % 유의성 검정
                    t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
                    p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

                    fprintf('  %s: r = %.3f (n=%d, p=%.3f)', extVar, correlation, validCount, p_value);
                    if p_value < 0.001
                        fprintf(' ***');
                    elseif p_value < 0.01
                        fprintf(' **');
                    elseif p_value < 0.05
                        fprintf(' *');
                    end
                    fprintf('\n');

                    % 결과 저장
                    extendedCorrelationResults.(sprintf('%s_vs_%s', perfVar, extVar)) = struct(...
                        'correlation', correlation, ...
                        'n', validCount, ...
                        'p_value', p_value);
                else
                    fprintf('  %s: 분석 불가 (유효 데이터 %d개)\n', extVar, validCount);
                end
            end
        end

        %% 결과 저장
        fprintf('\n▶ 확장 분석 결과 저장\n');

        if exist('outputFileName', 'var')
            % 확장 원본 데이터 저장
            writetable(extendedData, outputFileName, 'Sheet', '확장문항_원본데이터');

            % 확장 통합 데이터 저장
            writetable(extendedCombinedData, outputFileName, 'Sheet', '확장문항_통합데이터');

            % 확장 상관분석 결과 테이블 생성
            extCorrFields = fieldnames(extendedCorrelationResults);
            if ~isempty(extCorrFields)
                extCorrTable = table();
                extCorrTable.Analysis_Type = extCorrFields;
                extCorrTable.Correlation = NaN(length(extCorrFields), 1);
                extCorrTable.N = NaN(length(extCorrFields), 1);
                extCorrTable.P_Value = NaN(length(extCorrFields), 1);
                extCorrTable.Significance = cell(length(extCorrFields), 1);

                for i = 1:length(extCorrFields)
                    field = extCorrFields{i};
                    result = extendedCorrelationResults.(field);

                    extCorrTable.Correlation(i) = result.correlation;
                    extCorrTable.N(i) = result.n;
                    extCorrTable.P_Value(i) = result.p_value;

                    if result.p_value < 0.001
                        extCorrTable.Significance{i} = '***';
                    elseif result.p_value < 0.01
                        extCorrTable.Significance{i} = '**';
                    elseif result.p_value < 0.05
                        extCorrTable.Significance{i} = '*';
                    else
                        extCorrTable.Significance{i} = 'n.s.';
                    end
                end

                writetable(extCorrTable, outputFileName, 'Sheet', '확장문항_상관분석결과');
            end

            fprintf('  ✓ 확장 분석 결과 저장 완료\n');
        end

        %% 최종 요약
        fprintf('\n========================================\n');
        fprintf('확장된 조직지표+CSR 분석 최종 요약\n');
        fprintf('========================================\n');

        fprintf('📊 확장 데이터 현황\n');
        fprintf('  • 분석 대상: %d명\n', length(commonExtendedIDs));

        % 각 영역별 요약
        if ismember('Organizational_Score', analysisKeyVars)
            orgScores = extendedCombinedData.Organizational_Score(~isnan(extendedCombinedData.Organizational_Score));
            if ~isempty(orgScores)
                fprintf('  • 조직 기반 점수 (Q40,41): %.3f (±%.3f, n=%d)\n', mean(orgScores), std(orgScores), length(orgScores));
            end
        end

        if ismember('Total_CSR_Score', analysisKeyVars)
            csrScores = extendedCombinedData.Total_CSR_Score(~isnan(extendedCombinedData.Total_CSR_Score));
            if ~isempty(csrScores)
                fprintf('  • CSR 점수 (Q42-46): %.3f (±%.3f, n=%d)\n', mean(csrScores), std(csrScores), length(csrScores));
            end
        end

        if ismember('Total_Leadership_Score', analysisKeyVars)
            leadScores = extendedCombinedData.Total_Leadership_Score(~isnan(extendedCombinedData.Total_Leadership_Score));
            if ~isempty(leadScores)
                fprintf('  • 통합 리더십 점수 (Q40-46): %.3f (±%.3f, n=%d)\n', mean(leadScores), std(leadScores), length(leadScores));
            end
        end

        fprintf('\n🔗 주요 상관분석 결과\n');

        % 가장 높은 상관계수들 찾기 (Top 3)
        if exist('extCorrFields', 'var') && ~isempty(extCorrFields)
            corrValues = [];
            corrNames = {};
            for i = 1:length(extCorrFields)
                field = extCorrFields{i};
                corr = abs(extendedCorrelationResults.(field).correlation);
                if ~isnan(corr)
                    corrValues(end+1) = corr;
                    corrNames{end+1} = field;
                end
            end

            if ~isempty(corrValues)
                [sortedCorr, sortIdx] = sort(corrValues, 'descend');
                topN = min(3, length(sortedCorr));

                for i = 1:topN
                    fieldName = corrNames{sortIdx(i)};
                    result = extendedCorrelationResults.(fieldName);
                    sig_str = '';
                    if result.p_value < 0.001, sig_str = '***';
                    elseif result.p_value < 0.01, sig_str = '**';
                    elseif result.p_value < 0.05, sig_str = '*';
                    end
                    fprintf('  %d. %s: r=%.3f %s (n=%d)\n', i, fieldName, result.correlation, sig_str, result.n);
                end
            end
        end

        fprintf('\n✅ 확장된 조직지표+CSR 분석 완료!\n');

    else
        fprintf('[경고] 매칭된 데이터가 부족합니다 (최소 5명 필요).\n');
    end

else
    fprintf('[경고] 필요한 데이터를 찾을 수 없습니다.\n');
    if ~exist('combinedData', 'var')
        fprintf('  - combinedData 변수 없음\n');
    elseif ~istable(combinedData)
        fprintf('  - combinedData가 테이블이 아님\n');
    end
    if ~exist('extendedData', 'var')
        fprintf('  - extendedData 변수 없음\n');
    elseif ~istable(extendedData)
        fprintf('  - extendedData가 테이블이 아님\n');
    end
end
%% 13.25년 상반기 면담용 문항 개별 분석 (Q28, Q29, Q30)

fprintf('\n========================================\n');
fprintf('25년 상반기 면담용 문항 개별 분석\n');
fprintf('========================================\n\n');

%% 13-1. 25년 상반기 데이터에서 면담용 문항 추출
period_index = 4; % 25년 상반기
period_name = '25년_상반기';

if isfield(allData, sprintf('period%d', period_index)) && ...
        isfield(allData.(sprintf('period%d', period_index)), 'originalSelfData')

    fprintf('▶ %s 원본 데이터에서 면담용 문항 추출\n', period_name);

    originalData = allData.(sprintf('period%d', period_index)).originalSelfData;

    % ID 컬럼 찾기
    idCol = findIDColumn(originalData);
    if isempty(idCol)
        fprintf('[오류] ID 컬럼을 찾을 수 없습니다.\n');
        return;
    end

    % ID 추출
    responseIDs = extractAndStandardizeIDs(originalData{:, idCol});

    % 면담용 문항 데이터 테이블 생성
    interviewItemData = table();
    interviewItemData.ID = responseIDs;

    % Q28, Q29, Q30 개별 추출
    interviewQuestions = {'Q28', 'Q29', 'Q30'};
    foundQuestions = {};

    for q = 1:length(interviewQuestions)
        qCode = interviewQuestions{q};

        % 해당 문항 컬럼 찾기
        questionCol = [];
        colNames = originalData.Properties.VariableNames;

        for col = 1:width(originalData)
            colName = colNames{col};
            if strcmp(colName, qCode)
                questionCol = col;
                break;
            end
        end

        if ~isempty(questionCol)
            colData = originalData{:, questionCol};

            % 데이터 타입 확인
            if isnumeric(colData)
                numericCount = sum(~isnan(colData));
                fprintf('  ✓ %s 발견: 숫자형 %d개 유효 응답\n', qCode, numericCount);

                % 테이블에 추가
                interviewItemData.(qCode) = colData;
                foundQuestions{end+1} = qCode;

                % 기초 통계
                validData = colData(~isnan(colData));
                if ~isempty(validData)
                    fprintf('    평균: %.2f, 표준편차: %.2f, 범위: %.0f-%.0f\n', ...
                        mean(validData), std(validData), min(validData), max(validData));
                end
            else
                fprintf('  ✗ %s 발견: 질적 데이터 (상관분석 제외)\n', qCode);
            end
        else
            fprintf('  ✗ %s 누락\n', qCode);
        end
    end

    fprintf('\n발견된 숫자형 면담용 문항: %d개 (%s)\n', ...
        length(foundQuestions), strjoin(foundQuestions, ', '));
    fprintf('총 응답자: %d명\n', height(interviewItemData));

else
    fprintf('[오류] %s 원본 데이터를 찾을 수 없습니다.\n', period_name);
    return;
end

%% 13-2. 성과 데이터와 매칭
fprintf('\n▶ 성과 데이터와 매칭\n');

if exist('consolidatedScores', 'var') && istable(consolidatedScores)

    % ID를 문자열로 통일
    if isnumeric(interviewItemData.ID)
        interviewIDs = arrayfun(@num2str, interviewItemData.ID, 'UniformOutput', false);
    else
        interviewIDs = cellfun(@char, interviewItemData.ID, 'UniformOutput', false);
    end

    if isnumeric(consolidatedScores.ID)
        performanceIDs = arrayfun(@num2str, consolidatedScores.ID, 'UniformOutput', false);
    else
        performanceIDs = cellfun(@char, consolidatedScores.ID, 'UniformOutput', false);
    end

    % 교집합 찾기
    [commonIDs, interviewIdx, performanceIdx] = intersect(interviewIDs, performanceIDs);

    fprintf('매칭 결과:\n');
    fprintf('  - 면담용 문항 데이터: %d명\n', height(interviewItemData));
    fprintf('  - 성과 데이터: %d명\n', height(consolidatedScores));
    fprintf('  - 공통 ID: %d명\n', length(commonIDs));
    fprintf('  - 매칭률: %.1f%%\n', 100 * length(commonIDs) / min(height(interviewItemData), height(consolidatedScores)));

    if length(commonIDs) >= 5
        % 매칭된 데이터 생성
        matchedData = table();
        matchedData.ID = commonIDs;

        % 면담용 문항 데이터 추가
        for q = 1:length(foundQuestions)
            qCode = foundQuestions{q};
            if ismember(qCode, interviewItemData.Properties.VariableNames)
                matchedData.(qCode) = interviewItemData.(qCode)(interviewIdx);
            end
        end

        % 성과 데이터 추가
        matchedData.Performance_Average = consolidatedScores.AverageStdScore(performanceIdx);
        matchedData.Performance_Period4 = consolidatedScores.Period4_Score(performanceIdx);

        fprintf('매칭된 데이터 생성 완료: %d명\n', height(matchedData));

    else
        fprintf('[경고] 매칭된 데이터가 부족합니다 (최소 5명 필요).\n');
        return;
    end

else
    fprintf('[경고] 성과 데이터(consolidatedScores)를 찾을 수 없습니다.\n');
    return;
end

%% 13-3. 개별 면담용 문항 vs 성과점수 상관분석
fprintf('\n========================================\n');
fprintf('개별 면담용 문항 vs 성과점수 상관분석\n');
fprintf('========================================\n\n');

correlationResults = struct();

% 분석할 성과 변수들
performanceVars = {'Performance_Average', 'Performance_Period4'};
performanceNames = {'전체 평균 성과점수', '25년 상반기 성과점수'};

for perfIdx = 1:length(performanceVars)
    perfVar = performanceVars{perfIdx};
    perfName = performanceNames{perfIdx};

    if ismember(perfVar, matchedData.Properties.VariableNames)
        fprintf('\n▶ %s와의 상관분석\n', perfName);

        perfScores = matchedData.(perfVar);
        validPerfIdx = ~isnan(perfScores);

        fprintf('유효한 성과점수: %d명\n', sum(validPerfIdx));

        if sum(validPerfIdx) >= 5

            % 각 면담용 문항별 상관분석
            for q = 1:length(foundQuestions)
                qCode = foundQuestions{q};

                if ismember(qCode, matchedData.Properties.VariableNames)
                    qScores = matchedData.(qCode);

                    % 양쪽 모두 유효한 데이터만 선택
                    bothValidIdx = validPerfIdx & ~isnan(qScores);
                    validCount = sum(bothValidIdx);

                    if validCount >= 5
                        % 상관계수 계산
                        r = corrcoef(perfScores(bothValidIdx), qScores(bothValidIdx));
                        correlation = r(1, 2);

                        % 유의성 검정
                        t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
                        p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));

                        % 결과 출력
                        fprintf('  %s: r = %.3f (n=%d, p=%.3f)', qCode, correlation, validCount, p_value);
                        if p_value < 0.001
                            fprintf(' ***');
                        elseif p_value < 0.01
                            fprintf(' **');
                        elseif p_value < 0.05
                            fprintf(' *');
                        end
                        fprintf('\n');

                        % 결과 저장
                        correlationResults.(sprintf('%s_vs_%s', qCode, perfVar)) = struct(...
                            'question', qCode, ...
                            'performance_var', perfVar, ...
                            'correlation', correlation, ...
                            'n', validCount, ...
                            'p_value', p_value);

                    else
                        fprintf('  %s: 분석 불가 (유효 데이터 %d개)\n', qCode, validCount);
                    end
                end
            end
        else
            fprintf('유효한 성과점수가 부족합니다.\n');
        end
    end
end

%% 13-4. 면담용 문항 간 상관분석 (다중공선성 확인)
if length(foundQuestions) >= 2
    fprintf('\n▶ 면담용 문항 간 상관분석 (다중공선성 확인)\n');

    % 면담용 문항들만 추출
    interviewMatrix = [];
    interviewColNames = {};

    for q = 1:length(foundQuestions)
        qCode = foundQuestions{q};
        if ismember(qCode, matchedData.Properties.VariableNames)
            qScores = matchedData.(qCode);
            if ~all(isnan(qScores))
                interviewMatrix = [interviewMatrix, qScores];
                interviewColNames{end+1} = qCode;
            end
        end
    end

    if size(interviewMatrix, 2) >= 2
        % 상관행렬 계산
        R = corrcoef(interviewMatrix, 'Rows', 'pairwise');

        fprintf('면담용 문항 간 상관행렬:\n');
        fprintf('        ');
        for i = 1:length(interviewColNames)
            fprintf('%8s', interviewColNames{i});
        end
        fprintf('\n');

        for i = 1:length(interviewColNames)
            fprintf('%8s', interviewColNames{i});
            for j = 1:length(interviewColNames)
                if i == j
                    fprintf('%8s', '1.000');
                elseif i < j
                    fprintf('%8.3f', R(i,j));
                else
                    fprintf('%8s', '');
                end
            end
            fprintf('\n');
        end

        % 높은 상관 찾기
        fprintf('\n높은 상관 (|r| > 0.7):\n');
        highCorr = false;
        for i = 1:length(interviewColNames)
            for j = i+1:length(interviewColNames)
                if abs(R(i,j)) > 0.7
                    fprintf('  %s - %s: r = %.3f\n', interviewColNames{i}, interviewColNames{j}, R(i,j));
                    highCorr = true;
                end
            end
        end
        if ~highCorr
            fprintf('  높은 상관을 보이는 문항 쌍이 없습니다.\n');
        end
    end
end

%% 13-5. 결과 요약 및 저장
fprintf('\n========================================\n');
fprintf('분석 결과 요약\n');
fprintf('========================================\n');

fprintf('📊 데이터 현황\n');
fprintf('  • 분석 시점: %s\n', period_name);
fprintf('  • 발견된 면담용 문항: %d개 (%s)\n', length(foundQuestions), strjoin(foundQuestions, ', '));
fprintf('  • 분석 대상자: %d명\n', length(commonIDs));

fprintf('\n🔍 상관분석 결과\n');
corrFields = fieldnames(correlationResults);
if ~isempty(corrFields)

    % 가장 높은 상관 찾기
    maxAbsCorr = 0;
    maxCorrField = '';
    significantCount = 0;

    for i = 1:length(corrFields)
        result = correlationResults.(corrFields{i});
        absCorr = abs(result.correlation);

        if absCorr > maxAbsCorr
            maxAbsCorr = absCorr;
            maxCorrField = corrFields{i};
        end

        if result.p_value < 0.05
            significantCount = significantCount + 1;
        end
    end

    fprintf('  • 총 상관분석: %d개\n', length(corrFields));
    fprintf('  • 유의한 상관: %d개\n', significantCount);

    if ~isempty(maxCorrField)
        maxResult = correlationResults.(maxCorrField);
        sig_str = '';
        if maxResult.p_value < 0.001, sig_str = '***';
        elseif maxResult.p_value < 0.01, sig_str = '**';
        elseif maxResult.p_value < 0.05, sig_str = '*';
        end
        fprintf('  • 최고 상관: %s (r=%.3f) %s\n', maxCorrField, maxResult.correlation, sig_str);
    end
else
    fprintf('  • 상관분석 결과 없음\n');
end

%% 13-6. Excel 파일 저장
if exist('outputFileName', 'var')
    fprintf('\n▶ 결과 저장\n');

    % 매칭된 데이터 저장
    writetable(matchedData, outputFileName, 'Sheet', '면담문항_25년상반기_매칭데이터');

    % 상관분석 결과 저장
    if ~isempty(corrFields)
        corrTable = table();
        corrTable.Analysis_Type = corrFields;
        corrTable.Question = cell(length(corrFields), 1);
        corrTable.Performance_Variable = cell(length(corrFields), 1);
        corrTable.Correlation = NaN(length(corrFields), 1);
        corrTable.N = NaN(length(corrFields), 1);
        corrTable.P_Value = NaN(length(corrFields), 1);
        corrTable.Significance = cell(length(corrFields), 1);

        for i = 1:length(corrFields)
            field = corrFields{i};
            result = correlationResults.(field);

            corrTable.Question{i} = result.question;
            corrTable.Performance_Variable{i} = result.performance_var;
            corrTable.Correlation(i) = result.correlation;
            corrTable.N(i) = result.n;
            corrTable.P_Value(i) = result.p_value;

            if result.p_value < 0.001
                corrTable.Significance{i} = '***';
            elseif result.p_value < 0.01
                corrTable.Significance{i} = '**';
            elseif result.p_value < 0.05
                corrTable.Significance{i} = '*';
            else
                corrTable.Significance{i} = 'n.s.';
            end
        end

        writetable(corrTable, outputFileName, 'Sheet', '면담문항_25년상반기_상관분석');
    end

    fprintf('  ✓ Excel 파일 저장 완료\n');
    fprintf('    - 면담문항_25년상반기_매칭데이터 시트\n');
    fprintf('    - 면담문항_25년상반기_상관분석 시트\n');
end

fprintf('\n✅ 25년 상반기 면담용 문항 개별 분석 완료!\n');fprintf('  • 성과 데이터와 교차검증\n');
%% ============================================================================
%% 보조 함수들
%% ============================================================================

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

function [p, chi2stat, dof] = bartlettSphericity(R, n)
% Bartlett 구형성 검정 (Bartlett, 1951)
% H0: R = I (상관행렬이 단위행렬)
% H1: R ≠ I (상관행렬이 단위행렬이 아님)

% 입력 검증
if ~ismatrix(R) || size(R,1) ~= size(R,2)
    error('R은 정방행렬이어야 합니다.');
end

if ~isscalar(n) || n <= 0
    error('n은 양수 스칼라여야 합니다.');
end

pvars = size(R, 1);

% 변수가 너무 적은 경우
if pvars < 2
    p = NaN;
    chi2stat = NaN;
    dof = 0;
    warning('변수가 2개 미만이므로 Bartlett 검정을 수행할 수 없습니다.');
    return;
end

% 행렬의 조건수 확인 (특이행렬 방지)
if cond(R) > 1e12
    warning('상관행렬의 조건수가 너무 높습니다. 결과를 신중히 해석하세요.');
end

try
    % 행렬식 계산 (로그값 사용으로 수치적 안정성 확보)
    detR = det(R);

    % 행렬식이 0이거나 음수인 경우 처리
    if detR <= 0
        warning('상관행렬의 행렬식이 0이거나 음수입니다. 특이행렬일 가능성이 있습니다.');
        p = NaN;
        chi2stat = NaN;
        dof = pvars * (pvars - 1) / 2;
        return;
    end

    lnDet = log(detR);

    % 카이제곱 검정통계량 계산 (스칼라 확보)
    chi2stat = -(n - 1 - (2*pvars + 5)/6) * lnDet;

    % 스칼라인지 확인
    if ~isscalar(chi2stat)
        warning('chi2stat이 스칼라가 아닙니다. 첫 번째 값을 사용합니다.');
        chi2stat = chi2stat(1);
    end

    % 자유도 (스칼라)
    dof = pvars * (pvars - 1) / 2;

    % chi2stat이 음수이거나 복소수인 경우 처리
    if ~isreal(chi2stat) || chi2stat < 0
        warning('검정통계량이 유효하지 않습니다 (음수 또는 복소수). NaN을 반환합니다.');
        p = NaN;
        chi2stat = NaN;
        return;
    end

    % p-값 계산 (스칼라 입력 보장)
    p = 1 - chi2cdf(chi2stat, dof);

    % p값이 유효한지 확인
    if ~isreal(p) || p < 0 || p > 1
        warning('p값이 유효하지 않습니다. NaN을 반환합니다.');
        p = NaN;
    end

catch ME
    warning('Bartlett 검정 계산 중 오류 발생: %s', ME.message);
    p = NaN;
    chi2stat = NaN;
    dof = pvars * (pvars - 1) / 2;
end
end


function kmo = calculateKMO(data)
% Kaiser-Meyer-Olkin 측도 계산
R = corrcoef(data);

% 역행렬이 존재하지 않는 경우 처리
try
    R_inv = inv(R);
catch
    kmo = 0;
    return;
end

% 편상관계수 행렬
A = zeros(size(R));
for i = 1:size(R, 1)
    for j = 1:size(R, 2)
        if i ~= j
            A(i,j) = -R_inv(i,j) / sqrt(R_inv(i,i) * R_inv(j,j));
        end
    end
end

% KMO 계산
R2 = R.^2;
A2 = A.^2;

% 대각선 제외
R2_sum = sum(R2(:)) - sum(diag(R2));
A2_sum = sum(A2(:));

if (R2_sum + A2_sum) == 0
    kmo = 0;
else
    kmo = R2_sum / (R2_sum + A2_sum);
end
end





function alpha = cronbachAlpha(data)
% 크론바흐 알파 계산
k = size(data, 2);
if k < 2
    alpha = NaN;
    return;
end

itemVar = var(data, 0, 1);
totalVar = var(sum(data, 2));

if totalVar == 0
    alpha = NaN;
else
    alpha = (k / (k - 1)) * (1 - sum(itemVar) / totalVar);
end
end
