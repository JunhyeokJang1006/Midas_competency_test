%% 역량진단 데이터 요인분석 및 성과점수 산출
% 2023년 하반기 ~ 2025년 상반기 (4개 시점) 데이터 분석
% 
% 작성일: 2025년
% 목적: 요인 분석을 통한 성과 관련 문항 묶음 및 점수화

clear; clc; close all;

%% 1. 데이터 파일 경로 설정
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터'; % 데이터 파일이 있는 경로
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

% 결과 저장을 위한 구조체 초기화
allData = struct();
factorResults = struct();

%% 2. 각 시점별 데이터 로드 및 전처리
fprintf('========================================\n');
fprintf('데이터 로드 및 전처리 시작\n');
fprintf('========================================\n\n');

for p = 1:length(periods)
    fprintf('[%s 데이터 처리 중...]\n', periods{p});
    
    % 파일 읽기
    fileName = fullfile(dataPath, fileNames{p});
    
    % 2.1 마스터 ID 리스트 로드
    masterIDs = readtable(fileName, 'Sheet', '기준인원 검토','VariableNamingRule','preserve');
    
    % 2.2 자가진단 데이터 로드
    selfData = readtable(fileName, 'Sheet', '자가 진단','VariableNamingRule','preserve' );
    
    % 2.3 수평진단 데이터 로드 및 집계
    peerData = readtable(fileName, 'Sheet', '수평 진단','VariableNamingRule','preserve');
    
    % 2.4 하향진단 데이터 로드
    downData = readtable(fileName, 'Sheet', '하향 진단','VariableNamingRule','preserve');
    
    % 2.5 문항 정보 로드
    questionInfo_self = readtable(fileName, 'Sheet', '문항 정보_자가진단','VariableNamingRule','preserve');
    questionInfo_other = readtable(fileName, 'Sheet', '문항 정보_타인진단','VariableNamingRule','preserve');
    
    
    % 데이터 정리
    allData.(sprintf('period%d', p)).masterIDs = masterIDs;
    allData.(sprintf('period%d', p)).selfData = selfData;
    allData.(sprintf('period%d', p)).peerData = peerData;
    allData.(sprintf('period%d', p)).downData = downData;
    allData.(sprintf('period%d', p)).questionInfo_self = questionInfo_self;
    allData.(sprintf('period%d', p)).questionInfo_other = questionInfo_other;
    
    fprintf('  - 마스터 ID 수: %d\n', height(masterIDs));
    fprintf('  - 자가진단 응답자: %d\n', height(selfData));
    fprintf('  - 수평진단 응답 수: %d\n', height(peerData));
    fprintf('  - 하향진단 응답 수: %d\n', height(downData));
    fprintf('\n');
end

%% 3. 자가진단 데이터 요인 분석 (수정된 버전)
fprintf('========================================\n');
fprintf('자가진단 데이터 요인 분석\n');
fprintf('========================================\n\n');

for p = 1:length(periods)
    fprintf('[%s 요인 분석]\n', periods{p});
    
    % 자가진단 데이터 추출
    selfData = allData.(sprintf('period%d', p)).selfData;
    
    % 데이터 구조 확인
    fprintf('  - 테이블 크기: %dx%d\n', height(selfData), width(selfData));
    fprintf('  - 컬럼명: ');
    disp(selfData.Properties.VariableNames);
    
    % 숫자 컬럼만 찾기 (Q로 시작하는 문항들)
    numericCols = [];
    colNames = selfData.Properties.VariableNames;
    
    for col = 1:width(selfData)
        colName = colNames{col};
        colData = selfData{:, col};
        
        % 숫자 데이터인지 확인 (Q로 시작하거나 숫자 데이터)
        if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
            numericCols = [numericCols, col];
        end
    end
    
    fprintf('  - 발견된 숫자 컬럼 수: %d\n', length(numericCols));
    
    % 숫자 컬럼이 있는 경우에만 분석 진행
    if length(numericCols) >= 3  % 최소 3개 문항은 있어야 요인분석 가능
        % 숫자 컬럼만 추출
        responseData = table2array(selfData(:, numericCols));
        
        fprintf('  - 분석 데이터 크기: %dx%d\n', size(responseData, 1), size(responseData, 2));
        
        % 결측치 확인
        missingCount = sum(isnan(responseData(:)));
        fprintf('  - 결측치 개수: %d\n', missingCount);
        
        % 결측치 처리 (평균으로 대체)
        for col = 1:size(responseData, 2)
            missingIdx = isnan(responseData(:, col));
            if any(missingIdx)
                colMean = mean(responseData(~missingIdx, col));
                if ~isnan(colMean)  % 컬럼 전체가 결측치가 아닌 경우만
                    responseData(missingIdx, col) = colMean;
                else
                    % 전체가 결측치인 경우 중간값(예: 3)으로 대체
                    responseData(missingIdx, col) = 3;
                end
            end
        end
        
        % 데이터 유효성 검사
        if size(responseData, 1) < 3 || size(responseData, 2) < 3
            fprintf('  [경고] 데이터가 부족합니다 (최소 3x3 필요)\n');
            continue;
        end
        
        % 상관관계 매트릭스 확인 (특이행렬 방지)
        corrMatrix = corrcoef(responseData);
        if det(corrMatrix) < 1e-10
            fprintf('  [경고] 상관관계 매트릭스가 특이행렬에 가깝습니다\n');
        end
        
        % 요인 분석 수행
        try
            % 요인 개수 결정 (Kaiser 기준: eigenvalue > 1, 최대 5개)
            [coeff, score, latent, ~, explained] = pca(responseData);
            numFactors = min(sum(latent > 1), 5);  % 최대 5개로 제한
            
            if numFactors < 1
                numFactors = 1;  % 최소 1개 요인
            end
            
            fprintf('  - 고유값 > 1인 요인 수: %d\n', numFactors);
            
            % 요인 분석 (Varimax 회전)
            [loadings, specificVar, T, stats, F] = ...
                factoran(responseData, numFactors, 'rotate', 'varimax');
            
            % 결과 저장
            factorResults.(sprintf('period%d', p)).loadings = loadings;
            factorResults.(sprintf('period%d', p)).scores = F;
            factorResults.(sprintf('period%d', p)).explained = stats.dfe;
            factorResults.(sprintf('period%d', p)).numFactors = numFactors;
            factorResults.(sprintf('period%d', p)).numericCols = numericCols;  % 사용된 컬럼 인덱스 저장
            
            fprintf('  - 추출된 요인 수: %d\n', numFactors);
            fprintf('  - 요인분석 성공\n');
            
            % 성과 관련 요인 식별
            questionInfo_self = allData.(sprintf('period%d', p)).questionInfo_self;
            [performanceFactorIdx, performanceItems] = identifyPerformanceFactor(loadings, ...
                questionInfo_self, numericCols);
            
            factorResults.(sprintf('period%d', p)).performanceFactorIdx = performanceFactorIdx;
            factorResults.(sprintf('period%d', p)).performanceItems = performanceItems;
            
        catch ME
            fprintf('  [경고] 요인 분석 실패: %s\n', ME.message);
            fprintf('  [대안] PCA 점수 사용\n');
            
            % 실패 시 PCA 점수 사용
            [coeff, score, latent] = pca(responseData);
            numFactorsPC = min(3, size(score, 2));
            
            factorResults.(sprintf('period%d', p)).scores = score(:, 1:numFactorsPC);
            factorResults.(sprintf('period%d', p)).numFactors = numFactorsPC;
            factorResults.(sprintf('period%d', p)).numericCols = numericCols;
            factorResults.(sprintf('period%d', p)).loadings = coeff(:, 1:numFactorsPC);  % PCA 로딩
        end
        
    else
        fprintf('  [경고] 숫자형 문항이 부족합니다 (최소 3개 필요, 현재 %d개)\n', length(numericCols));
        fprintf('  사용 가능한 컬럼들:\n');
        for col = 1:width(selfData)
            colName = colNames{col};
            colData = selfData{:, col};
            fprintf('    %s: %s\n', colName, class(colData));
        end
    end
    
    fprintf('\n');
end

%% 4. 수평진단 데이터 집계 및 요인 분석 (수정된 버전)
fprintf('========================================\n');
fprintf('수평진단 데이터 집계 및 분석\n');
fprintf('========================================\n\n');

for p = 1:length(periods)
    fprintf('[%s 수평진단 집계]\n', periods{p});
    
    peerData = allData.(sprintf('period%d', p)).peerData;
    
    % 대상자별 평균 점수 계산
    if height(peerData) > 0
        % 데이터 구조 확인
        fprintf('  - 수평진단 데이터 크기: %dx%d\n', height(peerData), width(peerData));
        fprintf('  - 컬럼명: ');
        disp(peerData.Properties.VariableNames);
        
        % 두 번째 컬럼의 데이터 타입 확인
        targetIDColumn = peerData{:, 2};
        fprintf('  - 대상자 ID 컬럼 타입: %s\n', class(targetIDColumn));
        
        % 대상자 ID 추출 및 타입별 처리
        if iscell(targetIDColumn)
            % Cell 타입인 경우
            targetIDs = unique(targetIDColumn);
            % 빈 셀 제거
            targetIDs = targetIDs(~cellfun(@isempty, targetIDs));
            
            fprintf('  - 대상자 ID 샘플: ');
            if length(targetIDs) > 0
                disp(targetIDs(1:min(3, length(targetIDs))));
            end
            
        elseif isnumeric(targetIDColumn)
            % 숫자 타입인 경우
            targetIDs = unique(targetIDColumn);
            targetIDs = targetIDs(~isnan(targetIDs));
            
        elseif isstring(targetIDColumn) || ischar(targetIDColumn)
            % 문자열 타입인 경우
            targetIDs = unique(targetIDColumn);
            
        else
            fprintf('  [경고] 지원되지 않는 ID 데이터 타입: %s\n', class(targetIDColumn));
            continue;
        end
        
        fprintf('  - 고유 대상자 수: %d\n', length(targetIDs));
        
        % 숫자 컬럼 찾기 (3번째 컬럼부터)
        numericCols = [];
        colNames = peerData.Properties.VariableNames;
        
        for col = 3:width(peerData)  % 3번째 컬럼부터 확인
            colData = peerData{:, col};
            if isnumeric(colData)
                numericCols = [numericCols, col];
            end
        end
        
        fprintf('  - 발견된 숫자 컬럼 수: %d\n', length(numericCols));
        
        if length(numericCols) < 3
            fprintf('  [경고] 분석할 숫자 컬럼이 부족합니다\n');
            continue;
        end
        
        % 대상자별 데이터 집계
        aggregatedPeerData = [];
        validTargetIDs = {};
        aggregatedCount = 0;
        
        for i = 1:length(targetIDs)
            % 타입별 매칭 방법
            if iscell(targetIDColumn)
                % Cell 타입 매칭
                if iscell(targetIDs)
                    targetRows = strcmp(targetIDColumn, targetIDs{i});
                else
                    targetRows = strcmp(targetIDColumn, targetIDs(i));
                end
            elseif isnumeric(targetIDColumn)
                % 숫자 타입 매칭
                targetRows = (targetIDColumn == targetIDs(i));
            elseif isstring(targetIDColumn) || ischar(targetIDColumn)
                % 문자열 타입 매칭
                targetRows = strcmp(targetIDColumn, targetIDs(i));
            end
            
            % 해당 대상자의 응답 추출
            if sum(targetRows) > 0
                try
                    % 숫자 컬럼만 추출
                    targetResponses = table2array(peerData(targetRows, numericCols));
                    
                    % 평균 계산
                    meanResponses = mean(targetResponses, 1, 'omitnan');
                    
                    % 유효한 데이터인지 확인 (모든 값이 NaN이 아닌 경우)
                    if ~all(isnan(meanResponses))
                        aggregatedCount = aggregatedCount + 1;
                        aggregatedPeerData(aggregatedCount, :) = meanResponses;
                        validTargetIDs{aggregatedCount} = targetIDs(i);
                    end
                    
                catch ME
                    fprintf('  [경고] 대상자 %d 처리 중 오류: %s\n', i, ME.message);
                end
            end
        end
        
        fprintf('  - 성공적으로 집계된 대상자 수: %d\n', aggregatedCount);
        
        % 요인 분석 수행
        if size(aggregatedPeerData, 1) > 10 && size(aggregatedPeerData, 2) >= 3
            try
                % 결측치 처리
                for col = 1:size(aggregatedPeerData, 2)
                    missingIdx = isnan(aggregatedPeerData(:, col));
                    if any(missingIdx)
                        colMean = mean(aggregatedPeerData(~missingIdx, col));
                        if ~isnan(colMean)
                            aggregatedPeerData(missingIdx, col) = colMean;
                        else
                            aggregatedPeerData(missingIdx, col) = 3; % 기본값
                        end
                    end
                end
                
                numFactors = min(3, floor(size(aggregatedPeerData, 2)/3));
                if numFactors < 1
                    numFactors = 1;
                end
                
                [loadings_peer, ~, ~, ~, F_peer] = ...
                    factoran(aggregatedPeerData, numFactors, 'rotate', 'varimax');
                
                factorResults.(sprintf('period%d', p)).peer_scores = F_peer;
                factorResults.(sprintf('period%d', p)).peer_loadings = loadings_peer;
                factorResults.(sprintf('period%d', p)).peer_targetIDs = validTargetIDs;
                
                fprintf('  - 집계된 평가 대상자 수: %d\n', aggregatedCount);
                fprintf('  - 추출된 요인 수: %d\n', numFactors);
                
            catch ME
                fprintf('  [경고] 수평진단 요인 분석 실패: %s\n', ME.message);
                
                % 대안: PCA 사용
                try
                    [coeff_peer, score_peer] = pca(aggregatedPeerData);
                    numFactorsPC = min(3, size(score_peer, 2));
                    
                    factorResults.(sprintf('period%d', p)).peer_scores = score_peer(:, 1:numFactorsPC);
                    factorResults.(sprintf('period%d', p)).peer_loadings = coeff_peer(:, 1:numFactorsPC);
                    factorResults.(sprintf('period%d', p)).peer_targetIDs = validTargetIDs;
                    
                    fprintf('  - PCA로 대체 분석 완료 (요인 수: %d)\n', numFactorsPC);
                catch
                    fprintf('  [경고] PCA 분석도 실패\n');
                end
            end
        else
            fprintf('  [경고] 요인 분석을 위한 데이터 부족 (샘플: %d, 변수: %d)\n', ...
                size(aggregatedPeerData, 1), size(aggregatedPeerData, 2));
        end
    else
        fprintf('  - 수평진단 데이터가 없습니다\n');
    end
    
    fprintf('\n');
end

%% 5. 종합 성과 점수 산출 (개선된 버전)
fprintf('========================================\n');
fprintf('종합 성과 점수 산출 (NaN 처리 개선)\n');
fprintf('========================================\n\n');

% 모든 시점의 마스터 ID 통합
allMasterIDs = [];
for p = 1:length(periods)
    masterTable = allData.(sprintf('period%d', p)).masterIDs;
    if height(masterTable) > 0
        idCol = [];
        % ID 컬럼 찾기
        for col = 1:width(masterTable)
            colName = masterTable.Properties.VariableNames{col};
            if strcmpi(colName, 'ID') || strcmpi(colName, 'idhr') || contains(colName, 'id', 'IgnoreCase', true)
                idCol = col;
                break;
            end
        end
        
        if ~isempty(idCol)
            ids = masterTable{:, idCol};
            if isnumeric(ids)
                ids = arrayfun(@num2str, ids, 'UniformOutput', false);
            end
            allMasterIDs = [allMasterIDs; ids];
        end
    end
end

allMasterIDs = unique(allMasterIDs);
fprintf('총 마스터 ID 수: %d\n', length(allMasterIDs));

% 성과 점수 테이블 초기화
performanceScores = table();
performanceScores.ID = allMasterIDs;
performanceScores.Score_Period1 = NaN(length(allMasterIDs), 1);
performanceScores.Score_Period2 = NaN(length(allMasterIDs), 1);
performanceScores.Score_Period3 = NaN(length(allMasterIDs), 1);
performanceScores.Score_Period4 = NaN(length(allMasterIDs), 1);

% 각 시점별 유효한 데이터 카운트
validDataCount = zeros(4, 1);
totalAttempts = zeros(4, 1);

% 각 시점별 점수 계산
for p = 1:length(periods)
    fprintf('\n[%s 성과 점수 계산]\n', periods{p});
    
    % 요인분석 결과 확인
    if ~isfield(factorResults, sprintf('period%d', p)) || ...
       ~isfield(factorResults.(sprintf('period%d', p)), 'scores')
        fprintf('  - 요인분석 결과 없음. 건너뜀.\n');
        continue;
    end
    
    selfData = allData.(sprintf('period%d', p)).selfData;
    scores = factorResults.(sprintf('period%d', p)).scores;
    
    if isempty(scores)
        fprintf('  - 점수 데이터 없음. 건너뜀.\n');
        continue;
    end
    
    fprintf('  - 분석된 응답자 수: %d\n', size(scores, 1));
    
    % 성과 요인 점수 추출
    if isfield(factorResults.(sprintf('period%d', p)), 'performanceFactorIdx')
        perfFactorIdx = factorResults.(sprintf('period%d', p)).performanceFactorIdx;
        if ~isempty(perfFactorIdx) && perfFactorIdx <= size(scores, 2)
            perfScores = scores(:, perfFactorIdx);
        else
            perfScores = scores(:, 1);  % 첫 번째 요인 사용
        end
    else
        perfScores = scores(:, 1);
    end
    
    % ID 컬럼 찾기 (개선된 방법)
    idCol = [];
    for col = 1:width(selfData)
        colName = selfData.Properties.VariableNames{col};
        % ID 관련 컬럼명 패턴 확장
        if contains(lower(colName), {'id', 'idhr', '사번', 'empno', 'employee'})
            colData = selfData{:, col};
            % 데이터 타입 확인
            if (isnumeric(colData) && ~all(isnan(colData))) || ...
               (iscell(colData) && ~all(cellfun(@isempty, colData))) || ...
               (isstring(colData) && ~all(ismissing(colData)))
                idCol = col;
                break;
            end
        end
    end
    
    if isempty(idCol)
        fprintf('  [경고] ID 컬럼을 찾을 수 없음\n');
        fprintf('  사용 가능한 컬럼: ');
        disp(selfData.Properties.VariableNames);
        continue;
    end
    
    % ID 데이터 추출 및 표준화
    selfIDs = selfData{:, idCol};
    if isnumeric(selfIDs)
        selfIDs = arrayfun(@(x) sprintf('%.0f', x), selfIDs, 'UniformOutput', false);
    elseif iscell(selfIDs)
        % 이미 cell이므로 그대로 사용, 단 빈 값 처리
        selfIDs = cellfun(@(x) char(x), selfIDs, 'UniformOutput', false);
    end
    
    totalAttempts(p) = length(selfIDs);
    
    % ID 매칭하여 점수 할당
    successCount = 0;
    for i = 1:length(selfIDs)
        if ~isempty(selfIDs{i}) && ~strcmp(selfIDs{i}, 'NaN')
            % 마스터 ID와 매칭
            idx = find(strcmp(performanceScores.ID, selfIDs{i}));
            if ~isempty(idx)
                performanceScores.(sprintf('Score_Period%d', p))(idx) = perfScores(i);
                successCount = successCount + 1;
            else
                % 마스터 ID에 없는 경우 - 추가할지 확인
                fprintf('  [참고] 마스터에 없는 ID: %s\n', selfIDs{i});
            end
        end
    end
    
    validDataCount(p) = successCount;
    fprintf('  - 성공적으로 매칭된 점수: %d/%d\n', successCount, totalAttempts(p));
end

% 통계 요약
fprintf('\n[시점별 데이터 현황]\n');
for p = 1:4
    validCount = sum(~isnan(performanceScores.(sprintf('Score_Period%d', p))));
    fprintf('%s: %d명 (전체 %d명 중 %.1f%%)\n', ...
        periods{p}, validCount, length(allMasterIDs), 100*validCount/length(allMasterIDs));
end

% 개선된 평균 성과 점수 계산
scoreMatrix = [performanceScores.Score_Period1, performanceScores.Score_Period2, ...
               performanceScores.Score_Period3, performanceScores.Score_Period4];

% 각 개인별 유효한 점수 개수 확인
validScoresPerPerson = sum(~isnan(scoreMatrix), 2);

fprintf('\n[개인별 유효 점수 분포]\n');
for i = 1:4
    count = sum(validScoresPerPerson == i);
    fprintf('%d개 시점 데이터 보유: %d명 (%.1f%%)\n', i, count, 100*count/length(allMasterIDs));
end

% 평균 계산 (최소 1개 이상의 유효한 점수가 있는 경우만)
performanceScores.AverageScore = NaN(height(performanceScores), 1);
performanceScores.ValidPeriodCount = validScoresPerPerson;

for i = 1:height(performanceScores)
    if validScoresPerPerson(i) > 0
        validScores = scoreMatrix(i, ~isnan(scoreMatrix(i, :)));
        performanceScores.AverageScore(i) = mean(validScores);
    end
end

% 유효한 평균 점수가 있는 경우만 표준화 및 백분위 계산
validAvgIdx = ~isnan(performanceScores.AverageScore);
validAvgScores = performanceScores.AverageScore(validAvgIdx);

if length(validAvgScores) > 1
    % 표준화 (Z-score)
    performanceScores.StandardizedScore = NaN(height(performanceScores), 1);
    performanceScores.StandardizedScore(validAvgIdx) = zscore(validAvgScores);
    
    % 백분위 점수
    performanceScores.PercentileRank = NaN(height(performanceScores), 1);
    performanceScores.PercentileRank(validAvgIdx) = ...
        100 * tiedrank(validAvgScores) / length(validAvgScores);
else
    performanceScores.StandardizedScore = NaN(height(performanceScores), 1);
    performanceScores.PercentileRank = NaN(height(performanceScores), 1);
end

fprintf('\n[최종 결과 요약]\n');
fprintf('전체 대상자: %d명\n', height(performanceScores));
fprintf('유효한 평균 점수 보유자: %d명 (%.1f%%)\n', ...
    sum(validAvgIdx), 100*sum(validAvgIdx)/height(performanceScores));
fprintf('완전한 NaN 행: %d명 (%.1f%%)\n', ...
    sum(validScoresPerPerson == 0), 100*sum(validScoresPerPerson == 0)/height(performanceScores));

% NaN 원인 분석 추가
fprintf('\n[NaN 원인 분석]\n');
fprintf('1. 마스터 ID는 있지만 실제 진단 미참여: %d명\n', sum(validScoresPerPerson == 0));
fprintf('2. 시점별 부분 참여 (1-3개 시점): %d명\n', sum(validScoresPerPerson > 0 & validScoresPerPerson < 4));
fprintf('3. 전체 시점 참여: %d명\n', sum(validScoresPerPerson == 4));

% 데이터 품질 개선 제안
fprintf('\n[데이터 품질 개선 제안]\n');
fprintf('- 마스터 ID 리스트를 실제 참여자 기준으로 정리\n');
fprintf('- 시점별 참여 현황 사전 확인\n');
fprintf('- 최소 참여 시점 기준 설정 (예: 2개 이상)\n');
%% 6. 결과 시각화
fprintf('========================================\n');
fprintf('결과 시각화\n');
fprintf('========================================\n\n');

% 6.1 시점별 성과 점수 분포
figure('Position', [100, 100, 1200, 600]);

subplot(2, 3, 1);
histogram(performanceScores.Score_Period1, 20);
title('23년 하반기 성과 점수 분포');
xlabel('성과 점수'); ylabel('빈도');

subplot(2, 3, 2);
histogram(performanceScores.Score_Period2, 20);
title('24년 상반기 성과 점수 분포');
xlabel('성과 점수'); ylabel('빈도');

subplot(2, 3, 3);
histogram(performanceScores.Score_Period3, 20);
title('24년 하반기 성과 점수 분포');
xlabel('성과 점수'); ylabel('빈도');

subplot(2, 3, 4);
histogram(performanceScores.Score_Period4, 20);
title('25년 상반기 성과 점수 분포');
xlabel('성과 점수'); ylabel('빈도');

subplot(2, 3, 5);
histogram(performanceScores.AverageScore, 20);
title('평균 성과 점수 분포');
xlabel('평균 성과 점수'); ylabel('빈도');

subplot(2, 3, 6);
boxplot([performanceScores.Score_Period1, performanceScores.Score_Period2, ...
         performanceScores.Score_Period3, performanceScores.Score_Period4], ...
         'Labels', periods);
title('시점별 성과 점수 비교');
ylabel('성과 점수');
xtickangle(45);

% 6.2 요인 부하량 히트맵 (첫 번째 시점 예시)
if isfield(factorResults.period1, 'loadings')
    figure('Position', [100, 100, 800, 600]);
    imagesc(factorResults.period1.loadings');
    colorbar;
    title('23년 하반기 요인 부하량 매트릭스');
    xlabel('문항 번호');
    ylabel('요인');
    colormap('jet');
end

%% 7. 결과 저장
fprintf('========================================\n');
fprintf('결과 저장\n');
fprintf('========================================\n\n');

% Excel 파일로 저장
writetable(performanceScores, 'performance_scores_output.xlsx', 'Sheet', '종합성과점수');

% 요약 통계 생성
summaryStats = table();
summaryStats.Period = periods';
summaryStats.N = [sum(~isnan(performanceScores.Score_Period1)); ...
                  sum(~isnan(performanceScores.Score_Period2)); ...
                  sum(~isnan(performanceScores.Score_Period3)); ...
                  sum(~isnan(performanceScores.Score_Period4))];
summaryStats.Mean = [nanmean(performanceScores.Score_Period1); ...
                     nanmean(performanceScores.Score_Period2); ...
                     nanmean(performanceScores.Score_Period3); ...
                     nanmean(performanceScores.Score_Period4)];
summaryStats.Std = [nanstd(performanceScores.Score_Period1); ...
                    nanstd(performanceScores.Score_Period2); ...
                    nanstd(performanceScores.Score_Period3); ...
                    nanstd(performanceScores.Score_Period4)];

writetable(summaryStats, 'performance_scores_output.xlsx', 'Sheet', '요약통계');

% MAT 파일로 전체 결과 저장
save('factor_analysis_results.mat', 'allData', 'factorResults', 'performanceScores');

fprintf('분석 완료!\n');
fprintf('결과 파일:\n');
fprintf('  - performance_scores_output.xlsx: 종합 성과 점수\n');
fprintf('  - factor_analysis_results.mat: 전체 분석 결과\n');

%% 6. 성과기여도 데이터 로드 및 전처리
fprintf('========================================\n');
fprintf('성과기여도 데이터 분석\n');
fprintf('========================================\n\n');

% 성과기여도 데이터 파일 경로 설정
contributionDataPath = 'D:\project\HR데이터\데이터\최근 3년 입사자_인적정보.xlsx';

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

%% 7. 성과기여도 점수 계산
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

%% 8. 시점별 성과기여도 집계 (반기별)
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

%% 9. 역량진단 성과점수와 성과기여도 매칭
fprintf('\n========================================\n');
fprintf('역량진단 vs 성과기여도 데이터 매칭\n');
fprintf('========================================\n\n');

% ID를 기준으로 두 데이터셋 매칭
% performanceScores (역량진단 기반) vs contributionByPeriod (성과기여도 기반)

% ID를 문자열로 통일
performanceIDs = cellfun(@char, num2cell(performanceScores.ID), 'UniformOutput', false);
if isnumeric(contributionByPeriod.ID)
    contributionIDs = arrayfun(@num2str, contributionByPeriod.ID, 'UniformOutput', false);
else
    contributionIDs = cellfun(@char, contributionByPeriod.ID, 'UniformOutput', false);
end

% 교집합 찾기
[commonIDs, performanceIdx, contributionIdx] = intersect(performanceIDs, contributionIDs);

fprintf('매칭 결과:\n');
fprintf('  - 역량진단 데이터: %d명\n', height(performanceScores));
fprintf('  - 성과기여도 데이터: %d명\n', height(contributionByPeriod));
fprintf('  - 공통 ID: %d명\n', length(commonIDs));
fprintf('  - 매칭률: %.1f%%\n', 100 * length(commonIDs) / min(height(performanceScores), height(contributionByPeriod)));

if length(commonIDs) < 10
    fprintf('[경고] 매칭된 데이터가 너무 적습니다. ID 형식을 확인해주세요.\n');
    fprintf('샘플 ID (역량진단): ');
    disp(performanceIDs(1:min(5, end)));
    fprintf('샘플 ID (성과기여도): ');
    disp(contributionIDs(1:min(5, end)));
end

% 매칭된 데이터로 통합 테이블 생성
if length(commonIDs) > 0
    combinedData = table();
    combinedData.ID = commonIDs;
    
    % 역량진단 점수 추가
    combinedData.Factor_Period1 = performanceScores.Score_Period1(performanceIdx);
    combinedData.Factor_Period2 = performanceScores.Score_Period2(performanceIdx);
    combinedData.Factor_Period3 = performanceScores.Score_Period3(performanceIdx);
    combinedData.Factor_Period4 = performanceScores.Score_Period4(performanceIdx);
    combinedData.Factor_Average = performanceScores.AverageScore(performanceIdx);
    
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

%% 10. 상관분석
fprintf('\n========================================\n');
fprintf('역량진단 성과점수 vs 성과기여도 상관분석\n');
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

%% 11. 결과 시각화
fprintf('\n========================================\n');
fprintf('상관분석 시각화\n');
fprintf('========================================\n\n');

% 시점별 상관계수 그래프
figure('Position', [100, 100, 1200, 800]);

% 서브플롯 1: 시점별 상관계수
subplot(2, 3, 1);
correlations = [];
valid_periods = [];
for p = 1:4
    corr_data = correlationResults.(sprintf('period%d', p));
    if ~isnan(corr_data.correlation)
        correlations = [correlations; corr_data.correlation];
        valid_periods = [valid_periods; p];
    end
end

if ~isempty(correlations)
    bar(valid_periods, correlations);
    title('시점별 상관계수');
    xlabel('시점');
    ylabel('상관계수');
    ylim([-1, 1]);
    grid on;
    
    % 유의수준 표시
    for i = 1:length(valid_periods)
        p_val = correlationResults.(sprintf('period%d', valid_periods(i))).p_value;
        if p_val < 0.05
            text(valid_periods(i), correlations(i) + 0.05, '*', ...
                'HorizontalAlignment', 'center', 'FontSize', 14);
        end
    end
end

% 서브플롯 2-5: 시점별 산점도
for p = 1:4
    subplot(2, 3, p + 1);
    
    factorCol = sprintf('Factor_Period%d', p);
    contribCol = sprintf('Contribution_Period%d', p);
    
    factorScores = combinedData.(factorCol);
    contribScores = combinedData.(contribCol);
    
    validIdx = ~isnan(factorScores) & ~isnan(contribScores);
    
    if sum(validIdx) >= 5
        scatter(factorScores(validIdx), contribScores(validIdx), 'filled');
        
        % 회귀선 추가
        p_fit = polyfit(factorScores(validIdx), contribScores(validIdx), 1);
        x_fit = linspace(min(factorScores(validIdx)), max(factorScores(validIdx)), 100);
        y_fit = polyval(p_fit, x_fit);
        hold on;
        plot(x_fit, y_fit, 'r-', 'LineWidth', 2);
        
        title(sprintf('%s (r=%.3f)', periods{p}, correlationResults.(sprintf('period%d', p)).correlation));
        xlabel('역량진단 성과점수');
        ylabel('성과기여도 점수');
        grid on;
    else
        title(sprintf('%s (데이터 부족)', periods{p}));
    end
end

% 전체 평균 산점도
if isfield(correlationResults, 'overall') && ~isnan(correlationResults.overall.correlation)
    figure('Position', [200, 200, 600, 500]);
    
    validIdx = ~isnan(combinedData.Factor_Average) & ~isnan(combinedData.Contribution_Average);
    
    scatter(combinedData.Factor_Average(validIdx), combinedData.Contribution_Average(validIdx), ...
        'filled', 'SizeData', 100);
    
    % 회귀선 추가
    p_fit = polyfit(combinedData.Factor_Average(validIdx), combinedData.Contribution_Average(validIdx), 1);
    x_fit = linspace(min(combinedData.Factor_Average(validIdx)), max(combinedData.Factor_Average(validIdx)), 100);
    y_fit = polyval(p_fit, x_fit);
    hold on;
    plot(x_fit, y_fit, 'r-', 'LineWidth', 2);
    
    title(sprintf('전체 평균 성과점수 상관관계 (r=%.3f, p=%.3f)', ...
        correlationResults.overall.correlation, correlationResults.overall.p_value));
    xlabel('역량진단 성과점수 (평균)');
    ylabel('성과기여도 점수 (평균)');
    grid on;
    
    % 상관계수와 유의수준 텍스트 추가
    text(0.05, 0.95, sprintf('r = %.3f\nn = %d\np = %.3f', ...
        correlationResults.overall.correlation, ...
        correlationResults.overall.n, ...
        correlationResults.overall.p_value), ...
        'Units', 'normalized', 'VerticalAlignment', 'top', ...
        'BackgroundColor', 'white', 'EdgeColor', 'black');
end

%% 12. 결과 저장 및 요약
fprintf('\n========================================\n');
fprintf('결과 저장\n');
fprintf('========================================\n\n');

% 통합 데이터 저장
writetable(combinedData, 'performance_correlation_analysis.xlsx', 'Sheet', '통합데이터');

% 성과기여도 데이터 저장
writetable(contributionByPeriod, 'performance_correlation_analysis.xlsx', 'Sheet', '성과기여도점수');

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
        
        if result.p_value < 0.001
            corrResultTable.Significance{p} = '***';
        elseif result.p_value < 0.01
            corrResultTable.Significance{p} = '**';
        elseif result.p_value < 0.05
            corrResultTable.Significance{p} = '*';
        else
            corrResultTable.Significance{p} = '';
        end
    end
end

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
        corrResultTable.Significance{5} = '';
    end
end

writetable(corrResultTable, 'performance_correlation_analysis.xlsx', 'Sheet', '상관분석결과');

% MAT 파일로 전체 결과 저장
save('performance_contribution_analysis.mat', 'contributionScores', 'contributionByPeriod', ...
     'combinedData', 'correlationResults');

fprintf('분석 완료!\n');
fprintf('결과 파일:\n');
fprintf('  - performance_correlation_analysis.xlsx: 상관분석 결과\n');
fprintf('  - performance_contribution_analysis.mat: 전체 분석 결과\n\n');

fprintf('========================================\n');
fprintf('분석 요약\n');
fprintf('========================================\n');
fprintf('1. 매칭된 데이터: %d명\n', height(combinedData));
fprintf('2. 유의한 상관관계가 발견된 시점:\n');
for p = 1:4
    if isfield(correlationResults, sprintf('period%d', p))
        result = correlationResults.(sprintf('period%d', p));
        if ~isnan(result.p_value) && result.p_value < 0.05
            fprintf('   - %s: r=%.3f (p<0.05)\n', periods{p}, result.correlation);
        end
    end
end
if isfield(correlationResults, 'overall') && correlationResults.overall.p_value < 0.05
    fprintf('   - 전체 평균: r=%.3f (p<0.05)\n', correlationResults.overall.correlation);
end


%% 보조 함수 수정
function [perfFactorIdx, perfItems] = identifyPerformanceFactor(loadings, questionInfo, numericCols)
    % 성과 관련 키워드
    performanceKeywords = {'성과', '목표', '달성', '결과', '효과', '기여', '창출', '개선'};
    
    % 각 요인별 성과 관련성 점수 계산
    numFactors = size(loadings, 2);
    performanceRelevance = zeros(numFactors, 1);
    
    for f = 1:numFactors
        % 높은 부하량을 가진 문항들 찾기 (절대값 0.4 이상)
        highLoadingItems = find(abs(loadings(:, f)) > 0.4);
        
        % 해당 문항들의 내용 확인
        relevanceScore = 0;
        for itemIdx = 1:length(highLoadingItems)
            item = highLoadingItems(itemIdx);
            
            % numericCols를 고려한 실제 문항 번호 계산
            if item <= length(numericCols) && length(numericCols) <= height(questionInfo)
                actualItemNum = numericCols(item);
                if actualItemNum <= height(questionInfo)
                    try
                        questionText = questionInfo{actualItemNum, 2};  % 문항 내용이 2번째 컬럼이라고 가정
                        if iscell(questionText)
                            questionText = questionText{1};
                        end
                        questionText = char(questionText);
                        
                        % 키워드 매칭
                        for k = 1:length(performanceKeywords)
                            if contains(questionText, performanceKeywords{k})
                                relevanceScore = relevanceScore + abs(loadings(item, f));
                            end
                        end
                    catch
                        % 문항 정보 접근 실패 시 무시
                    end
                end
            end
        end
        
        performanceRelevance(f) = relevanceScore;
    end
    
    % 가장 성과 관련성이 높은 요인 선택
    [maxRelevance, perfFactorIdx] = max(performanceRelevance);
    
    % 해당 요인의 주요 문항들
    perfItems = find(abs(loadings(:, perfFactorIdx)) > 0.4);
    
    % 성과 요인을 찾지 못한 경우 첫 번째 요인 사용
    if maxRelevance == 0
        perfFactorIdx = 1;
        perfItems = find(abs(loadings(:, 1)) > 0.4);
    end
end