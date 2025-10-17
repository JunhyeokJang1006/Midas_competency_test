%% 각 Period별 문항과 역량검사 종합점수 간 상관 매트릭스 생성 (개선된 버전)
% 
% 목적: 각 시점별로 수집된 문항들과 역량검사 종합점수 간의 
%       전체 상관 매트릭스를 생성하고 분석
%
% 주요 개선사항:
% 1. Winsorization을 통한 이상치 처리
% 2. 상수항(분산이 0인 변수) 자동 제거
% 3. Min-Max 스케일링 결과를 0-100점으로 변환
%
% 작성일: 2025년

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('Period별 문항-종합점수 상관 매트릭스 생성 (개선된 버전)\n');
fprintf('========================================\n\n');

%% 1. 기존 분석 결과 로드
fprintf('[1단계] 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

% 기존 역량진단 데이터 로드
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

try
    competencyTestData = readtable(competencyTestPath, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('✓ 역량검사 데이터 로드 완료: %d명\n', height(competencyTestData));
catch ME
    fprintf('✗ 역량검사 데이터 로드 실패: %s\n', ME.message);
    return;
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
        allData.period1.selfData = readtable(fileName_23_1st, 'Sheet', '하향진단', 'VariableNamingRule', 'preserve');
        try
            allData.period1.questionInfo = readtable(fileName_23_1st, 'Sheet', '문항 정보', 'VariableNamingRule', 'preserve');
        catch
            allData.period1.questionInfo = table();
        end
        
        % 기존 데이터를 뒤로 밀기
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

%% 3. 각 시점별 성과 관련 문항 정의
performanceQuestions = struct();
performanceQuestions.period1 = {'Q3', 'Q4', 'Q5', 'Q22', 'Q23', 'Q45', 'Q46', 'Q51'};  % 23년 상반기
performanceQuestions.period2 = {'Q4', 'Q21', 'Q23', 'Q25', 'Q32', 'Q33', 'Q34'}; % 23년 하반기
performanceQuestions.period3 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'}; % 24년 상반기
performanceQuestions.period4 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'}; % 24년 하반기
performanceQuestions.period5 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'}; % 25년 상반기

%% 4. 종합성과 히스토그램 생성 (모든 기간 데이터 포함)
fprintf('\n[2단계] 종합성과 히스토그램 생성\n');
fprintf('----------------------------------------\n');

allPerformanceData = table();

for p = 1:length(periods)
    fprintf('\n▶ %s 종합성과점수 계산 중...\n', periods{p});
    
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
    
    % 성과 관련 문항 추출
    perfQuestions = performanceQuestions.(sprintf('period%d', p));
    availableQuestions = intersect(perfQuestions, questionCols);
    
    if length(availableQuestions) < 3
        fprintf('  [경고] 성과 관련 문항이 부족합니다 (%d개)\n', length(availableQuestions));
        continue;
    end
    
    % 성과 관련 문항의 인덱스 찾기
    perfIndices = [];
    for q = 1:length(availableQuestions)
        qIdx = find(strcmp(questionCols, availableQuestions{q}));
        if ~isempty(qIdx)
            perfIndices(end+1) = qIdx;
        end
    end
    
    % 성과 종합점수 계산
    performanceData = questionData(:, perfIndices);
    
    % Winsorization 적용
    performanceData = winsorizeData(performanceData, 0.05, 0.95);
    
    % 리커트 척도 표준화 적용 (0-100점)
    standardizedPerformanceData = standardizeQuestionScalesToHundred(performanceData, availableQuestions, p);
    
    % 각 응답자별 성과점수 계산
    performanceScores = nanmean(standardizedPerformanceData, 2);
    
    % 결측치가 너무 많은 응답자 제외
    validPerformanceRows = sum(~isnan(performanceData), 2) >= (length(perfIndices) * 0.5);
    cleanPerformanceScores = performanceScores(validPerformanceRows);
    cleanResponseIDs = responseIDs(validPerformanceRows);
    
    fprintf('  - 종합성과점수 계산 완료: %d명 (평균: %.2f, 표준편차: %.2f)\n', ...
        length(cleanPerformanceScores), mean(cleanPerformanceScores), std(cleanPerformanceScores));
    
    % 해당 시점의 데이터를 테이블로 구성
    tempTable = table();
    tempTable.ID = cleanResponseIDs;
    tempTable.PerformanceScore = cleanPerformanceScores;
    tempTable.Period = repmat({periods{p}}, length(cleanResponseIDs), 1);
    
    allPerformanceData = [allPerformanceData; tempTable];
end

% 개인별로 성과점수 평균 계산
fprintf('\n▶ 개인별 종합성과점수 통합 중...\n');
uniqueIDs = unique(allPerformanceData.ID);
integratedPerformanceScores = [];
integratedIDs = {};

for i = 1:length(uniqueIDs)
    personID = uniqueIDs{i};
    personData = allPerformanceData(strcmp(allPerformanceData.ID, personID), :);
    
    avgPerformanceScore = nanmean(personData.PerformanceScore);
    numPeriods = height(personData);
    
    integratedPerformanceScores(end+1) = avgPerformanceScore;
    integratedIDs{end+1} = personID;
    
    if mod(i, 20) == 0
        fprintf('  진행: %d/%d명 처리 완료\n', i, length(uniqueIDs));
    end
end

fprintf('  - 고유한 개인 수: %d명\n', length(uniqueIDs));
fprintf('  - 평균 참여 시점: %.1f개\n', height(allPerformanceData) / length(uniqueIDs));

% 종합성과 히스토그램 생성
if ~isempty(integratedPerformanceScores)
    figure('Name', '전체 종합성과점수 분포', 'Position', [500, 500, 800, 600]);
    
    validScores = integratedPerformanceScores(~isnan(integratedPerformanceScores));
    
    histogram(validScores, 30);
    title('전체 종합성과점수 분포 (고유한 개인별, 100점 만점)', 'FontSize', 14, 'FontWeight', 'bold');
    xlabel('종합성과점수 (100점 만점)', 'FontSize', 12);
    ylabel('빈도', 'FontSize', 12);
    grid on;
    
    meanScore = mean(validScores);
    stdScore = std(validScores);
    avgParticipation = height(allPerformanceData) / length(uniqueIDs);
    
    textStr = sprintf('평균: %.1f점\n표준편차: %.1f\nN: %d명 (고유한 개인)\n평균 참여: %.1f회\n범위: %.1f ~ %.1f점', ...
                     meanScore, stdScore, length(validScores), avgParticipation, ...
                     min(validScores), max(validScores));
    text(0.02, 0.98, textStr, 'Units', 'normalized', 'FontSize', 11, ...
         'BackgroundColor', 'white', 'EdgeColor', 'black', ...
         'VerticalAlignment', 'top', 'HorizontalAlignment', 'left');
    
    fprintf('✓ 전체 종합성과점수 히스토그램 생성 완료: %d명\n', length(validScores));
end

%% 5. 각 Period별 상관 분석
fprintf('\n[3단계] 각 Period별 문항 데이터 분석\n');
fprintf('----------------------------------------\n');

correlationMatrices = struct();

for p = 1:length(periods)
    fprintf('\n▶ %s 처리 중...\n', periods{p});
    
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
    
    % Winsorization 적용
    questionData = winsorizeData(questionData, 0.05, 0.95);
    
    % 상수항 제거
    [questionData, questionCols, removedCols] = removeConstantVariables(questionData, questionCols);
    if ~isempty(removedCols)
        fprintf('  - 상수항으로 제거된 문항: %s\n', strjoin(removedCols, ', '));
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
        correlationMatrices.(sprintf('period%d', p)) = struct(...
            'correlationMatrix', correlationMatrix, ...
            'pValues', pValues, ...
            'variableNames', {variableNames}, ...
            'questionNames', {questionCols}, ...
            'sampleSize', size(cleanData, 1), ...
            'cleanData', cleanData, ...
            'cleanIDs', {matchedIDs}, ...
            'removedConstantCols', {removedCols});
        
        fprintf('  ✓ 상관 매트릭스 계산 완료 (%dx%d)\n', ...
            size(correlationMatrix, 1), size(correlationMatrix, 2));
        
        displayTopCorrelations(correlationMatrix, pValues, questionCols);
    end
end

%% 6. 결과 저장
fprintf('\n[4단계] 결과 저장\n');
fprintf('----------------------------------------\n');

dateStr = datestr(now, 'yyyymmdd_HHMM');
outputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\correlation_matrices_improved_%s.xlsx', dateStr);

savedSheets = {};
periodFields = fieldnames(correlationMatrices);

for i = 1:length(periodFields)
    fieldName = periodFields{i};
    periodNum = str2double(fieldName(end));
    result = correlationMatrices.(fieldName);
    
    corrTable = array2table(result.correlationMatrix, ...
        'VariableNames', result.variableNames, ...
        'RowNames', result.variableNames);
    
    pTable = array2table(result.pValues, ...
        'VariableNames', result.variableNames, ...
        'RowNames', result.variableNames);
    
    corrSheetName = sprintf('%s_상관계수', periods{periodNum});
    pSheetName = sprintf('%s_p값', periods{periodNum});
    
    try
        writetable(corrTable, outputFileName, 'Sheet', corrSheetName, 'WriteRowNames', true);
        writetable(pTable, outputFileName, 'Sheet', pSheetName, 'WriteRowNames', true);
        
        savedSheets{end+1} = corrSheetName;
        savedSheets{end+1} = pSheetName;
        
        fprintf('✓ %s 매트릭스 저장 완료\n', periods{periodNum});
        
    catch ME
        fprintf('✗ %s 매트릭스 저장 실패: %s\n', periods{periodNum}, ME.message);
    end
end

%% 7. 요약 테이블 생성 및 저장
summaryTable = createSummaryTable(correlationMatrices, periods);

try
    writetable(summaryTable, outputFileName, 'Sheet', '분석요약');
    savedSheets{end+1} = '분석요약';
    fprintf('✓ 분석 요약 저장 완료\n');
catch ME
    fprintf('✗ 분석 요약 저장 실패: %s\n', ME.message);
end

matFileName = sprintf('D:\\project\\correlation_matrices_improved_%s.mat', dateStr);
save(matFileName, 'correlationMatrices', 'periods', 'summaryTable', 'allData');

%% 8. 성과 관련 문항 종합점수 분석
fprintf('\n[5단계] 성과 관련 문항 종합점수 분석 (역량검사와의 상관분석)\n');
fprintf('========================================\n');

performanceResults = struct();

for p = 1:length(periods)
    fprintf('\n▶ %s 성과점수 분석 중...\n', periods{p});
    
    if ~isfield(correlationMatrices, sprintf('period%d', p))
        fprintf('  [경고] %s 상관 매트릭스를 찾을 수 없습니다.\n', periods{p});
        continue;
    end
    
    result = correlationMatrices.(sprintf('period%d', p));
    perfQuestions = performanceQuestions.(sprintf('period%d', p));
    
    availableQuestions = intersect(perfQuestions, result.questionNames);
    missingQuestions = setdiff(perfQuestions, result.questionNames);
    
    fprintf('  - 정의된 성과문항: %d개 (%s)\n', length(perfQuestions), strjoin(perfQuestions, ', '));
    fprintf('  - 실제 사용 가능: %d개 (%s)\n', length(availableQuestions), strjoin(availableQuestions, ', '));
    
    if ~isempty(missingQuestions)
        fprintf('  - 누락된 문항: %s\n', strjoin(missingQuestions, ', '));
    end
    
    if length(availableQuestions) < 3
        fprintf('  [경고] 성과 관련 문항이 부족합니다 (%d개)\n', length(availableQuestions));
        continue;
    end
    
    perfIndices = [];
    for q = 1:length(availableQuestions)
        qIdx = find(strcmp(result.questionNames, availableQuestions{q}));
        if ~isempty(qIdx)
            perfIndices(end+1) = qIdx;
        end
    end
    
    performanceData = result.cleanData(:, perfIndices);
    
    % 100점 만점으로 표준화
    standardizedPerformanceData = standardizeQuestionScalesToHundred(performanceData, availableQuestions, p);
    
    performanceScores = nanmean(standardizedPerformanceData, 2);
    
    validPerformanceRows = sum(~isnan(performanceData), 2) >= (length(perfIndices) * 0.5);
    cleanPerformanceScores = performanceScores(validPerformanceRows);
    cleanAllData = result.cleanData(validPerformanceRows, :);
    
    fprintf('  - 성과점수 계산 완료: %d명 (평균: %.1f점, 표준편차: %.1f)\n', ...
        length(cleanPerformanceScores), mean(cleanPerformanceScores), std(cleanPerformanceScores));
    
    rawCompetencyTestScores = cleanAllData(:, end);
    
    % 역량검사 점수도 100점 만점으로 표준화
    competencyTestScores = standardizeCompetencyScoresToHundred(rawCompetencyTestScores);
    
    performanceCorrelationData = [competencyTestScores, cleanPerformanceScores];
    
    validRows = ~any(isnan(performanceCorrelationData), 2);
    cleanPerformanceCorrelationData = performanceCorrelationData(validRows, :);
    
    if size(cleanPerformanceCorrelationData, 1) < 3
        fprintf('  [경고] 역량검사-성과점수 상관분석을 위한 데이터가 부족합니다 (%d명)\n', ...
            size(cleanPerformanceCorrelationData, 1));
        continue;
    end
    
    try
        [corrCoeff, pValue] = corr(cleanPerformanceCorrelationData(:,1), cleanPerformanceCorrelationData(:,2));
        
        validPerformanceIndices = find(validPerformanceRows);
        finalValidIndices = validPerformanceIndices(validRows);
        cleanIDs = result.cleanIDs(finalValidIndices);
        
        performanceResults.(sprintf('period%d', p)) = struct(...
            'competencyTestScores', cleanPerformanceCorrelationData(:,1), ...
            'performanceScores', cleanPerformanceCorrelationData(:,2), ...
            'cleanIDs', {cleanIDs}, ...
            'correlation', corrCoeff, ...
            'pValue', pValue, ...
            'sampleSize', size(cleanPerformanceCorrelationData, 1), ...
            'performanceQuestions', {availableQuestions}, ...
            'performanceMean', mean(cleanPerformanceCorrelationData(:,2)), ...
            'performanceStd', std(cleanPerformanceCorrelationData(:,2)), ...
            'competencyMean', mean(cleanPerformanceCorrelationData(:,1)), ...
            'competencyStd', std(cleanPerformanceCorrelationData(:,1)));
        
        sig_str = '';
        if pValue < 0.001, sig_str = '***';
        elseif pValue < 0.01, sig_str = '**';
        elseif pValue < 0.05, sig_str = '*';
        end
        
        fprintf('  ✓ 역량검사점수 vs 성과점수 상관분석 완료\n');
        fprintf('    → r = %.3f (p = %.3f) %s (N = %d)\n', ...
            corrCoeff, pValue, sig_str, size(cleanPerformanceCorrelationData, 1));
        fprintf('    → 역량검사점수: 평균 %.1f점 (SD %.1f)\n', ...
            mean(cleanPerformanceCorrelationData(:,1)), std(cleanPerformanceCorrelationData(:,1)));
        fprintf('    → 성과점수: 평균 %.1f점 (SD %.1f)\n', ...
            mean(cleanPerformanceCorrelationData(:,2)), std(cleanPerformanceCorrelationData(:,2)));
        
    catch ME
        fprintf('  [오류] 성과점수 상관 분석 실패: %s\n', ME.message);
    end
end

%% 9. 성과점수 분석 결과 저장
if ~isempty(fieldnames(performanceResults))
    fprintf('\n[6단계] 역량검사-성과점수 상관분석 결과 저장\n');
    fprintf('----------------------------------------\n');
    
    perfSummaryTable = createPerformanceSummaryTable(performanceResults, periods);
    
    try
        writetable(perfSummaryTable, outputFileName, 'Sheet', '역량검사_성과점수_상관분석');
        fprintf('✓ 역량검사-성과점수 상관분석 결과 저장 완료\n');
    catch ME
        fprintf('✗ 역량검사-성과점수 상관분석 저장 실패: %s\n', ME.message);
    end
    
    save(matFileName, 'correlationMatrices', 'periods', 'summaryTable', 'allData', ...
         'performanceResults', 'performanceQuestions', 'perfSummaryTable', '-append');
    
    fprintf('\n📊 역량검사-성과점수 상관분석 완료 - %d개 시점 처리됨\n', length(fieldnames(performanceResults)));
end

%% 10. 종합 성과점수 분석
fprintf('\n[7단계] 종합 성과점수 분석 (5개 시점 통합)\n');
fprintf('========================================\n');

if ~isempty(fieldnames(performanceResults))
    [integratedPerformanceData, overallCorrelation] = integratePerformanceScores(performanceResults, competencyTestData, periods, correlationMatrices);
    
    if ~isempty(integratedPerformanceData)
        fprintf('✅ 종합 성과점수 분석 완료\n');
        fprintf('   → 전체 상관계수: r = %.3f (p = %.3f) %s (N = %d)\n', ...
            overallCorrelation.correlation, overallCorrelation.pValue, ...
            overallCorrelation.significance, overallCorrelation.sampleSize);
        
        try
            writetable(integratedPerformanceData, outputFileName, 'Sheet', '종합성과점수분석');
            fprintf('✓ 종합 성과점수 분석 결과 저장 완료\n');
        catch ME
            fprintf('✗ 종합 성과점수 분석 저장 실패: %s\n', ME.message);
        end
        
        save(matFileName, 'integratedPerformanceData', 'overallCorrelation', '-append');
    else
        fprintf('⚠️  종합 성과점수 계산을 위한 데이터가 부족합니다.\n');
    end
else
    fprintf('⚠️  성과점수 분석 결과가 없어서 종합 분석을 수행할 수 없습니다.\n');
    integratedPerformanceData = [];
    overallCorrelation = struct();
end

%% 11. 분포 기반 시각화
fprintf('\n[8단계] 분포 기반 시각화\n');
fprintf('========================================\n');

if length(fieldnames(correlationMatrices)) > 0
    periodFields = fieldnames(correlationMatrices);
    createDistributionVisualizations(correlationMatrices, periods, periodFields, performanceResults, integratedPerformanceData, overallCorrelation, competencyTestData);
end

fprintf('\n========================================\n');
fprintf('✅ 모든 분석이 완료되었습니다!\n');
fprintf('========================================\n');

%% ===== 보조 함수들 =====

function winsorizedData = winsorizeData(data, lowerPercentile, upperPercentile)
    % Winsorization을 적용하여 이상치를 제거하는 함수
    winsorizedData = data;
    
    for col = 1:size(data, 2)
        colData = data(:, col);
        validData = colData(~isnan(colData));
        
        if ~isempty(validData)
            lowerBound = prctile(validData, lowerPercentile * 100);
            upperBound = prctile(validData, upperPercentile * 100);
            
            colData(colData < lowerBound) = lowerBound;
            colData(colData > upperBound) = upperBound;
            
            winsorizedData(:, col) = colData;
        end
    end
end

function [cleanData, cleanCols, removedCols] = removeConstantVariables(data, colNames)
    % 상수항(분산이 0인 변수)을 제거하는 함수
    variances = var(data, 'omitnan');
    validCols = variances > 1e-10;
    
    cleanData = data(:, validCols);
    cleanCols = colNames(validCols);
    removedCols = colNames(~validCols);
end

function standardizedData = standardizeQuestionScalesToHundred(questionData, questionNames, periodNum)
    % 문항별 리커트 척도를 100점 만점으로 표준화하는 함수
    standardizedData = questionData;
    
    for i = 1:size(questionData, 2)
        columnData = questionData(:, i);
        validData = columnData(~isnan(columnData));
        
        if isempty(validData)
            continue;
        end
        
        actualMin = min(validData);
        actualMax = max(validData);
        
        % Min-Max 스케일링을 100점 만점으로
        if actualMax > actualMin
            standardizedData(:, i) = ((columnData - actualMin) / (actualMax - actualMin)) * 100;
        else
            standardizedData(:, i) = 50 * ones(size(columnData));
            standardizedData(isnan(columnData), i) = NaN;
        end
    end
end

function standardizedScores = standardizeCompetencyScoresToHundred(rawScores)
    % 역량검사 점수를 100점 만점으로 표준화하는 함수
    validScores = rawScores(~isnan(rawScores));
    
    if isempty(validScores)
        standardizedScores = rawScores;
        return;
    end
    
    minScore = min(validScores);
    maxScore = max(validScores);
    
    if maxScore > minScore
        standardizedScores = ((rawScores - minScore) / (maxScore - minScore)) * 100;
    else
        standardizedScores = 50 * ones(size(rawScores));
        standardizedScores(isnan(rawScores)) = NaN;
    end
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

function standardizedIDs = extractAndStandardizeIDs(rawIDs)
    if isnumeric(rawIDs)
        standardizedIDs = arrayfun(@(x) sprintf('%.0f', x), rawIDs, 'UniformOutput', false);
    elseif iscell(rawIDs)
        standardizedIDs = cellfun(@(x) char(x), rawIDs, 'UniformOutput', false);
    else
        standardizedIDs = cellstr(rawIDs);
    end
    
    emptyIdx = cellfun(@isempty, standardizedIDs) | strcmp(standardizedIDs, 'NaN');
    standardizedIDs(emptyIdx) = {''};
end

function [competencyScore, usedColumnName] = findCompetencyScoreColumn(competencyTestData, testIdx)
    competencyScore = [];
    usedColumnName = '';
    colNames = competencyTestData.Properties.VariableNames;
    
    exactMatches = {'Average_Competency_Score', 'CompetencyScore', 'Competency_Score'};
    for i = 1:length(exactMatches)
        if ismember(exactMatches{i}, colNames)
            competencyScore = competencyTestData.(exactMatches{i})(testIdx);
            usedColumnName = exactMatches{i};
            return;
        end
    end
    
    scoreKeywords = {'총점', '종합점수', '평균점수', '총합', '합계', 'total', 'average', 'score', '점수'};
    
    for col = 1:width(competencyTestData)
        colName = colNames{col};
        colNameLower = lower(colName);
        
        if contains(colNameLower, {'id', '사번', 'empno'})
            continue;
        end
        
        for k = 1:length(scoreKeywords)
            if contains(colNameLower, lower(scoreKeywords{k}))
                colData = competencyTestData{:, col};
                if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
                    competencyScore = colData(testIdx);
                    usedColumnName = colName;
                    return;
                end
            end
        end
    end
    
    candidateColumns = {};
    candidateScores = [];
    
    for col = 2:width(competencyTestData)
        colName = colNames{col};
        colNameLower = lower(colName);
        colData = competencyTestData{:, col};
        
        if contains(colNameLower, {'id', '사번', 'empno', 'employee'})
            continue;
        end
        
        if isnumeric(colData) && ~all(isnan(colData))
            colVariance = var(colData, 'omitnan');
            if colVariance > 0
                candidateColumns{end+1} = colName;
                
                score = 0;
                colMean = mean(colData, 'omitnan');
                if colMean >= 1 && colMean <= 100
                    score = score + 3;
                elseif colMean >= 0.1 && colMean <= 10
                    score = score + 2;
                end
                
                if colVariance > 0.1 && colVariance < 1000
                    score = score + 2;
                end
                
                missingRate = sum(isnan(colData)) / length(colData);
                if missingRate < 0.1
                    score = score + 1;
                end
                
                candidateScores(end+1) = score;
            end
        end
    end
    
    if ~isempty(candidateColumns)
        [~, bestIdx] = max(candidateScores);
        bestColumn = candidateColumns{bestIdx};
        competencyScore = competencyTestData.(bestColumn)(testIdx);
        usedColumnName = bestColumn;
    end
end

function [matchedData, matchedIDs, sampleSize] = matchWithCompetencyTest(questionData, responseIDs, competencyTestData)
    if isnumeric(competencyTestData.ID)
        testIDs = arrayfun(@num2str, competencyTestData.ID, 'UniformOutput', false);
    else
        testIDs = cellfun(@char, competencyTestData.ID, 'UniformOutput', false);
    end
    
    [commonIDs, responseIdx, testIdx] = intersect(responseIDs, testIDs);
    
    if length(commonIDs) >= 5
        matchedQuestionData = questionData(responseIdx, :);
        
        [competencyScore, usedColumnName] = findCompetencyScoreColumn(competencyTestData, testIdx);
        
        if ~isempty(competencyScore)
            matchedData = [matchedQuestionData, competencyScore];
            matchedIDs = commonIDs;
            sampleSize = length(commonIDs);
        else
            matchedData = [];
            matchedIDs = {};
            sampleSize = 0;
        end
    else
        matchedData = [];
        matchedIDs = {};
        sampleSize = 0;
    end
end

function [correlationMatrix, pValues, cleanData, variableNames] = performCorrelationAnalysis(matchedData, questionCols)
    if isempty(matchedData)
        correlationMatrix = [];
        pValues = [];
        cleanData = [];
        variableNames = {};
        return;
    end
    
    variableNames = [questionCols, {'CompetencyTest_Total'}];
    
    validRows = sum(isnan(matchedData), 2) < (size(matchedData, 2) * 0.5);
    cleanData = matchedData(validRows, :);
    
    variances = var(cleanData, 'omitnan');
    validCols = ~isnan(variances) & variances > 1e-10;
    
    if sum(validCols) < 2
        correlationMatrix = [];
        pValues = [];
        cleanData = [];
        variableNames = {};
        return;
    end
    
    cleanData = cleanData(:, validCols);
    variableNames = variableNames(validCols);
    
    try
        correlationMatrix = corrcoef(cleanData, 'Rows', 'pairwise');
        
        n = size(cleanData, 1);
        tStat = correlationMatrix .* sqrt((n-2) ./ (1 - correlationMatrix.^2));
        pValues = 2 * (1 - tcdf(abs(tStat), n-2));
        
        pValues(logical(eye(size(pValues)))) = 0;
        
    catch
        correlationMatrix = [];
        pValues = [];
        cleanData = [];
        variableNames = {};
    end
end

function displayTopCorrelations(correlationMatrix, pValues, questionCols)
    if size(correlationMatrix, 2) < 2
        return;
    end
    
    lastColIdx = size(correlationMatrix, 2);
    questionCorrs = correlationMatrix(1:end-1, lastColIdx);
    questionPvals = pValues(1:end-1, lastColIdx);
    
    [~, sortIdx] = sort(abs(questionCorrs), 'descend');
    fprintf('  상위 5개 문항의 종합점수와의 상관:\n');
    
    for i = 1:min(5, length(sortIdx))
        idx = sortIdx(i);
        if idx <= length(questionCols)
            qName = questionCols{idx};
            corr = questionCorrs(idx);
            pval = questionPvals(idx);
            
            sig_str = '';
            if pval < 0.001, sig_str = '***';
            elseif pval < 0.01, sig_str = '**';
            elseif pval < 0.05, sig_str = '*';
            end
            
            fprintf('    %s: r=%.3f (p=%.3f) %s\n', qName, corr, pval, sig_str);
        end
    end
end

function summaryTable = createSummaryTable(correlationMatrices, periods)
    summaryTable = table();
    periodFields = fieldnames(correlationMatrices);
    
    for i = 1:length(periodFields)
        fieldName = periodFields{i};
        periodNum = str2double(fieldName(end));
        result = correlationMatrices.(fieldName);
        
        newRow = table();
        newRow.Period = {periods{periodNum}};
        newRow.SampleSize = result.sampleSize;
        newRow.NumQuestions = length(result.questionNames);
        
        if size(result.correlationMatrix, 2) >= 2
            lastColIdx = size(result.correlationMatrix, 2);
            questionCorrs = result.correlationMatrix(1:end-1, lastColIdx);
            
            newRow.MaxCorrelation = max(abs(questionCorrs));
            newRow.MinCorrelation = min(abs(questionCorrs));
            newRow.MeanCorrelation = mean(abs(questionCorrs));
            newRow.SignificantCorrs = sum(result.pValues(1:end-1, lastColIdx) < 0.05);
        else
            newRow.MaxCorrelation = NaN;
            newRow.MinCorrelation = NaN;
            newRow.MeanCorrelation = NaN;
            newRow.SignificantCorrs = 0;
        end
        
        if isfield(result, 'removedConstantCols') && ~isempty(result.removedConstantCols)
            newRow.RemovedConstants = {strjoin(result.removedConstantCols, ', ')};
        else
            newRow.RemovedConstants = {''};
        end
        
        summaryTable = [summaryTable; newRow];
    end
end

function perfSummaryTable = createPerformanceSummaryTable(performanceResults, periods)
    perfSummaryTable = table();
    perfFields = fieldnames(performanceResults);
    
    for i = 1:length(perfFields)
        fieldName = perfFields{i};
        periodNum = str2double(fieldName(end));
        result = performanceResults.(fieldName);
        
        newRow = table();
        newRow.Period = {periods{periodNum}};
        newRow.SampleSize = result.sampleSize;
        newRow.NumPerformanceQuestions = length(result.performanceQuestions);
        newRow.PerformanceQuestions = {strjoin(result.performanceQuestions, ', ')};
        
        newRow.CompetencyMean = result.competencyMean;
        newRow.CompetencyStd = result.competencyStd;
        
        newRow.PerformanceMean = result.performanceMean;
        newRow.PerformanceStd = result.performanceStd;
        
        newRow.Correlation = result.correlation;
        newRow.PValue = result.pValue;
        
        if result.pValue < 0.001
            newRow.Significance = {'***'};
        elseif result.pValue < 0.01
            newRow.Significance = {'**'};
        elseif result.pValue < 0.05
            newRow.Significance = {'*'};
        else
            newRow.Significance = {'ns'};
        end
        
        perfSummaryTable = [perfSummaryTable; newRow];
    end
end

function [integratedData, overallCorrelation] = integratePerformanceScores(performanceResults, competencyTestData, periods, correlationMatrices)
    fprintf('▶ 개인별 성과점수 통합 중...\n');
    
    if isnumeric(competencyTestData.ID)
        testIDs = arrayfun(@num2str, competencyTestData.ID, 'UniformOutput', false);
    else
        testIDs = cellfun(@char, competencyTestData.ID, 'UniformOutput', false);
    end
    
    perfFields = fieldnames(performanceResults);
    allPerformanceData = table();
    
    fprintf('  - 수집 중인 시점: ');
    for i = 1:length(perfFields)
        fieldName = perfFields{i};
        periodNum = str2double(fieldName(end));
        result = performanceResults.(fieldName);
        
        fprintf('%s ', periods{periodNum});
        
        if isfield(result, 'cleanIDs')
            periodIDs = result.cleanIDs;
            
            tempTable = table();
            tempTable.ID = periodIDs;
            tempTable.PerformanceScore = result.performanceScores;
            tempTable.Period = repmat({periods{periodNum}}, length(periodIDs), 1);
            
            allPerformanceData = [allPerformanceData; tempTable];
        end
    end
    fprintf('\n');
    
    if height(allPerformanceData) == 0
        fprintf('  [경고] 통합할 성과점수 데이터가 없습니다.\n');
        integratedData = [];
        overallCorrelation = struct();
        return;
    end
    
    fprintf('  - 개인별 성과점수 평균 계산 중...\n');
    
    uniqueIDs = unique(allPerformanceData.ID);
    integratedTable = table();
    
    validCount = 0;
    for i = 1:length(uniqueIDs)
        personID = uniqueIDs{i};
        personData = allPerformanceData(strcmp(allPerformanceData.ID, personID), :);
        
        if height(personData) >= 1
            avgPerformanceScore = nanmean(personData.PerformanceScore);
            numPeriods = height(personData);
            
            newRow = table();
            newRow.ID = {personID};
            newRow.IntegratedPerformanceScore = avgPerformanceScore;
            newRow.NumPeriods = numPeriods;
            newRow.PerformanceScores = {personData.PerformanceScore'};
            newRow.Periods = {personData.Period'};
            
            integratedTable = [integratedTable; newRow];
            validCount = validCount + 1;
        end
    end
    
    fprintf('  - 통합 가능한 개인: %d명 (전체 %d명 중)\n', validCount, length(uniqueIDs));
    
    if validCount < 3
        fprintf('  [경고] 상관분석을 위한 데이터가 부족합니다 (%d명)\n', validCount);
        integratedData = [];
        overallCorrelation = struct();
        return;
    end
    
    fprintf('  - 역량검사점수와 매칭 중...\n');
    
    dummyTestIdx = 1:height(competencyTestData);
    [~, usedColumnName] = findCompetencyScoreColumn(competencyTestData, dummyTestIdx);
    
    if isempty(usedColumnName)
        fprintf('  [경고] 역량검사 점수 컬럼을 찾을 수 없습니다.\n');
        integratedData = [];
        overallCorrelation = struct();
        return;
    end
    
    rawCompetencyScores = competencyTestData.(usedColumnName);
    
    competencyScores = standardizeCompetencyScoresToHundred(rawCompetencyScores);
    
    [commonIDs, integratedIdx, testIdx] = intersect(integratedTable.ID, testIDs);
    
    if length(commonIDs) < 3
        fprintf('  [경고] 매칭된 데이터가 부족합니다 (%d명)\n', length(commonIDs));
        integratedData = [];
        overallCorrelation = struct();
        return;
    end
    
    finalTable = table();
    finalTable.ID = commonIDs;
    finalTable.CompetencyScore = competencyScores(testIdx);
    finalTable.PerformanceScore = integratedTable.IntegratedPerformanceScore(integratedIdx);
    finalTable.IntegratedPerformanceScore = integratedTable.IntegratedPerformanceScore(integratedIdx);
    finalTable.NumPeriods = integratedTable.NumPeriods(integratedIdx);
    
    for p = 1:length(periods)
        colName = sprintf('Performance_%s', periods{p});
        colData = nan(height(finalTable), 1);
        
        for i = 1:height(finalTable)
            personID = finalTable.ID{i};
            personIntegratedData = integratedTable(strcmp(integratedTable.ID, personID), :);
            if height(personIntegratedData) > 0
                periodScores = personIntegratedData.PerformanceScores{1};
                periodNames = personIntegratedData.Periods{1};
                
                periodIdx = find(strcmp(periodNames, periods{p}));
                if ~isempty(periodIdx)
                    colData(i) = periodScores(periodIdx);
                end
            end
        end
        
        finalTable.(colName) = colData;
    end
    
    fprintf('  - 최종 분석 대상: %d명\n', height(finalTable));
    fprintf('  - 평균 참여 시점: %.1f개\n', mean(finalTable.NumPeriods));
    
    validRows = ~isnan(finalTable.CompetencyScore) & ~isnan(finalTable.IntegratedPerformanceScore);
    cleanData = finalTable(validRows, :);
    
    if height(cleanData) < 3
        fprintf('  [경고] 상관분석을 위한 유효 데이터가 부족합니다 (%d명)\n', height(cleanData));
        integratedData = finalTable;
        overallCorrelation = struct();
        return;
    end
    
    [corrCoeff, pValue] = corr(cleanData.CompetencyScore, cleanData.IntegratedPerformanceScore);
    
    if pValue < 0.001
        significance = '***';
    elseif pValue < 0.01
        significance = '**';
    elseif pValue < 0.05
        significance = '*';
    else
        significance = 'ns';
    end
    
    overallCorrelation = struct();
    overallCorrelation.correlation = corrCoeff;
    overallCorrelation.pValue = pValue;
    overallCorrelation.significance = significance;
    overallCorrelation.sampleSize = height(cleanData);
    overallCorrelation.competencyMean = mean(cleanData.CompetencyScore);
    overallCorrelation.competencyStd = std(cleanData.CompetencyScore);
    overallCorrelation.performanceMean = mean(cleanData.IntegratedPerformanceScore);
    overallCorrelation.performanceStd = std(cleanData.IntegratedPerformanceScore);
    overallCorrelation.usedColumnName = usedColumnName;
    
    finalTable.CompetencyMean = repmat(overallCorrelation.competencyMean, height(finalTable), 1);
    finalTable.PerformanceMean = repmat(overallCorrelation.performanceMean, height(finalTable), 1);
    finalTable.OverallCorrelation = repmat(corrCoeff, height(finalTable), 1);
    finalTable.PValue = repmat(pValue, height(finalTable), 1);
    finalTable.Significance = repmat({significance}, height(finalTable), 1);
    
    integratedData = finalTable;
    
    fprintf('  - 종합 성과점수: 평균 %.1f점 (SD %.1f)\n', ...
        overallCorrelation.performanceMean, overallCorrelation.performanceStd);
    fprintf('  - 역량검사점수: 평균 %.1f점 (SD %.1f)\n', ...
        overallCorrelation.competencyMean, overallCorrelation.competencyStd);
end

function createDistributionVisualizations(correlationMatrices, periods, periodFields, performanceResults, integratedPerformanceData, overallCorrelation, competencyTestData)
    if isempty(periodFields)
        return;
    end
    
    fprintf('▶ 분포 기반 시각화 생성 중...\n');
    
    % 역량검사점수 분포 히스토그램
    figure('Name', '역량검사점수 분포 (100점 만점)', 'Position', [100, 100, 1400, 900]);
    
    for i = 1:min(length(periodFields), 5)
        fieldName = periodFields{i};
        periodNum = str2double(fieldName(end));
        result = correlationMatrices.(fieldName);
        
        subplot(2, 3, i);
        competencyData = result.cleanData(:, end);
        
        % 100점 만점으로 변환
        competencyData = standardizeCompetencyScoresToHundred(competencyData);
        validCompetencyData = competencyData(~isnan(competencyData));
        
        histogram(validCompetencyData, 20);
        title(sprintf('%s 역량검사점수 분포', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
        xlabel('역량검사점수 (100점 만점)');
        ylabel('빈도');
        grid on;
        
        meanScore = nanmean(validCompetencyData);
        stdScore = nanstd(validCompetencyData);
        text(0.6, 0.8, sprintf('평균: %.1f점\n표준편차: %.1f\nN: %d', meanScore, stdScore, length(validCompetencyData)), ...
             'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white');
    end
    
    fprintf('✓ 역량검사점수 분포 히스토그램 생성 완료\n');
    
    % 성과점수 분포 히스토그램
    if ~isempty(fieldnames(performanceResults))
        figure('Name', '성과점수 분포 (100점 만점)', 'Position', [150, 150, 1400, 900]);
        
        perfFields = fieldnames(performanceResults);
        
        for i = 1:min(length(perfFields), 5)
            fieldName = perfFields{i};
            periodNum = str2double(fieldName(end));
            result = performanceResults.(fieldName);
            
            subplot(2, 3, i);
            validPerformanceScores = result.performanceScores(~isnan(result.performanceScores));
            histogram(validPerformanceScores, 15);
            title(sprintf('%s 성과점수 분포', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
            xlabel('성과점수 (100점 만점)');
            ylabel('빈도');
            grid on;
            
            meanScore = result.performanceMean;
            stdScore = result.performanceStd;
            text(0.6, 0.8, sprintf('평균: %.1f점\n표준편차: %.1f\nN: %d', meanScore, stdScore, result.sampleSize), ...
                 'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white');
        end
        
        fprintf('✓ 성과점수 분포 히스토그램 생성 완료\n');
    end
    
    % 역량검사점수 vs 성과점수 산점도
    if ~isempty(fieldnames(performanceResults))
        figure('Name', '역량검사점수 vs 성과점수 산점도 (100점 만점)', 'Position', [200, 200, 1400, 900]);
        
        perfFields = fieldnames(performanceResults);
        
        for i = 1:min(length(perfFields), 5)
            fieldName = perfFields{i};
            periodNum = str2double(fieldName(end));
            result = performanceResults.(fieldName);
            
            subplot(2, 3, i);
            
            scatter(result.competencyTestScores, result.performanceScores, 50, 'filled');
            hold on;
            
            if length(result.competencyTestScores) > 1
                p = polyfit(result.competencyTestScores, result.performanceScores, 1);
                x_trend = linspace(min(result.competencyTestScores), max(result.competencyTestScores), 100);
                y_trend = polyval(p, x_trend);
                plot(x_trend, y_trend, 'r-', 'LineWidth', 2);
            end
            
            title(sprintf('%s: 역량검사 vs 성과점수', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
            xlabel('역량검사점수 (100점 만점)');
            ylabel('성과점수 (100점 만점)');
            grid on;
            xlim([0 100]);
            ylim([0 100]);
            
            corrText = sprintf('r = %.3f\np = %.3f\nN = %d', result.correlation, result.pValue, result.sampleSize);
            text(0.05, 0.95, corrText, 'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white', 'VerticalAlignment', 'top');
            
            hold off;
        end
        
        % 종합 상관분석 산점도
        if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData)
            subplot(2, 3, 6);
            
            scatter(integratedPerformanceData.CompetencyScore, integratedPerformanceData.PerformanceScore, 50, 'filled', 'MarkerFaceColor', [0.2 0.6 0.8]);
            hold on;
            
            if height(integratedPerformanceData) > 1
                p = polyfit(integratedPerformanceData.CompetencyScore, integratedPerformanceData.PerformanceScore, 1);
                x_trend = linspace(min(integratedPerformanceData.CompetencyScore), max(integratedPerformanceData.CompetencyScore), 100);
                y_trend = polyval(p, x_trend);
                plot(x_trend, y_trend, 'r-', 'LineWidth', 2);
            end
            
            title('종합: 역량검사 vs 성과점수', 'FontSize', 12, 'FontWeight', 'bold');
            xlabel('역량검사점수 (100점 만점)');
            ylabel('종합 성과점수 (100점 만점)');
            grid on;
            xlim([0 100]);
            ylim([0 100]);
            
            if isfield(overallCorrelation, 'correlation')
                corrText = sprintf('r = %.3f\np = %.3f\nN = %d', overallCorrelation.correlation, overallCorrelation.pValue, overallCorrelation.sampleSize);
                text(0.05, 0.95, corrText, 'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white', 'VerticalAlignment', 'top');
            end
            
            hold off;
        else
            subplot(2, 3, 6);
            text(0.5, 0.5, '종합 상관분석 데이터 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center', 'FontSize', 12);
            title('종합: 역량검사 vs 성과점수', 'FontSize', 12, 'FontWeight', 'bold');
        end
        
        fprintf('✓ 역량검사 vs 성과점수 산점도 생성 완료\n');
    end
    
    % 전체 역량검사점수 분포 (100점 만점)
    figure('Name', '전체 역량검사점수 분포 (100점 만점)', 'Position', [300, 300, 800, 600]);
    
    [~, usedColumnName] = findCompetencyScoreColumn(competencyTestData, 1:height(competencyTestData));
    if ~isempty(usedColumnName)
        allUniqueCompetencyScores = competencyTestData.(usedColumnName);
        
        % 100점 만점으로 변환
        allUniqueCompetencyScores = standardizeCompetencyScoresToHundred(allUniqueCompetencyScores);
        validAllCompetencyScores = allUniqueCompetencyScores(~isnan(allUniqueCompetencyScores));
        
        histogram(validAllCompetencyScores, 30);
        title('전체 역량검사점수 분포 (100점 만점)', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('역량검사점수 (100점 만점)', 'FontSize', 12);
        ylabel('빈도', 'FontSize', 12);
        grid on;
        xlim([0 100]);
        
        meanScore = nanmean(validAllCompetencyScores);
        stdScore = nanstd(validAllCompetencyScores);
        
        textStr = sprintf('평균: %.1f점\n표준편차: %.1f\nN: %d명\n범위: %.1f ~ %.1f점', ...
                         meanScore, stdScore, length(validAllCompetencyScores), ...
                         min(validAllCompetencyScores), max(validAllCompetencyScores));
        text(0.02, 0.98, textStr, 'Units', 'normalized', 'FontSize', 11, ...
             'BackgroundColor', 'white', 'EdgeColor', 'black', ...
             'VerticalAlignment', 'top', 'HorizontalAlignment', 'left');
        
        fprintf('✓ 전체 역량검사점수 분포 별도 figure 생성 완료\n');
    end
    
    % 종합 성과점수 분포 (100점 만점)
    if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData) && ismember('PerformanceScore', integratedPerformanceData.Properties.VariableNames)
        figure('Name', '종합 성과점수 분포 (100점 만점)', 'Position', [400, 400, 800, 600]);
        
        validIntegratedScores = integratedPerformanceData.PerformanceScore(~isnan(integratedPerformanceData.PerformanceScore));
        
        histogram(validIntegratedScores, 25);
        title('종합 성과점수 분포 (5개 시점 통합, 100점 만점)', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('종합 성과점수 (100점 만점)', 'FontSize', 12);
        ylabel('빈도', 'FontSize', 12);
        grid on;
        xlim([0 100]);
        
        meanScore = nanmean(integratedPerformanceData.PerformanceScore);
        stdScore = nanstd(integratedPerformanceData.PerformanceScore);
        numPeriods = nanmean(integratedPerformanceData.NumPeriods);
        
        textStr = sprintf('평균: %.1f점\n표준편차: %.1f\nN: %d명\n평균 참여횟수: %.1f회\n범위: %.1f ~ %.1f점', ...
                         meanScore, stdScore, length(validIntegratedScores), ...
                         numPeriods, min(validIntegratedScores), max(validIntegratedScores));
        text(0.02, 0.98, textStr, 'Units', 'normalized', 'FontSize', 11, ...
             'BackgroundColor', 'white', 'EdgeColor', 'black', ...
             'VerticalAlignment', 'top', 'HorizontalAlignment', 'left');
        
        fprintf('✓ 종합 성과점수 분포 별도 figure 생성 완료\n');
    end
    
    % Winsorization 효과 시각화 (Before/After 비교)
    figure('Name', 'Winsorization 효과', 'Position', [500, 100, 1200, 500]);
    
    % 첫 번째 period의 데이터를 예시로 사용
    if length(periodFields) >= 1
        fieldName = periodFields{1};
        periodNum = str2double(fieldName(end));
        
        % 원본 데이터 재로드 (Winsorization 전)
        if isfield(allData, sprintf('period%d', periodNum))
            selfData = allData.(sprintf('period%d', periodNum)).selfData;
            idCol = findIDColumn(selfData);
            [~, originalData] = extractQuestionData(selfData, idCol);
            
            if ~isempty(originalData)
                % 첫 번째 문항을 예시로 사용
                sampleCol = originalData(:, 1);
                validSample = sampleCol(~isnan(sampleCol));
                
                % Winsorization 적용
                winsorizedSample = winsorizeData(validSample, 0.05, 0.95);
                
                subplot(1, 2, 1);
                histogram(validSample, 30);
                title('Winsorization 전', 'FontSize', 12, 'FontWeight', 'bold');
                xlabel('값');
                ylabel('빈도');
                grid on;
                
                subplot(1, 2, 2);
                histogram(winsorizedSample, 30);
                title('Winsorization 후 (5%-95%)', 'FontSize', 12, 'FontWeight', 'bold');
                xlabel('값');
                ylabel('빈도');
                grid on;
                
                sgtitle(sprintf('%s 첫 번째 문항 예시', periods{periodNum}), 'FontSize', 14, 'FontWeight', 'bold');
                
                fprintf('✓ Winsorization 효과 시각화 생성 완료\n');
            end
        end
    end
    
    fprintf('✓ 모든 분포 기반 시각화 생성 완료\n');
end