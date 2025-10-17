%% 각 Period별 문항과 역량검사 종합점수 간 상관 매트릭스 생성
% 
% 목적: 각 시점별로 수집된 문항들과 역량검사 종합점수 간의 
%       전체 상관 매트릭스를 생성하고 분석
%
% 작성일: 2025년

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('Period별 문항-종합점수 상관 매트릭스 생성\n');
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
competencyTestPath = 'D:\project\HR데이터\결과\자가불소_revised\역량검사_가중치적용점수_2025-09-23_134448.xlsx';

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

%% 2-1. 종합성과 히스토그램 생성 (모든 기간 데이터 포함)
fprintf('\n[2-1단계] 종합성과 히스토그램 생성\n');
fprintf('----------------------------------------\n');

% 각 시점별 성과 관련 문항 정의
performanceQuestions = struct();
performanceQuestions.period1 = {'Q3', 'Q4', 'Q5', 'Q22', 'Q23', 'Q45', 'Q46', 'Q51'};  % 23년 상반기
performanceQuestions.period2 = {'Q4', 'Q21', 'Q23', 'Q25', 'Q32', 'Q33', 'Q34'}; % 23년 하반기
performanceQuestions.period3 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'}; % 24년 상반기
performanceQuestions.period4 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'}; % 24년 하반기
performanceQuestions.period5 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'}; % 25년 상반기

% 모든 기간의 종합성과점수 계산 (개인별로 통합)
allPerformanceData = table();

for p = 1:length(periods)
    fprintf('\n▶ %s 종합성과점수 계산 중...\n', periods{p});
    
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
    
    if length(perfIndices) < 3
        fprintf('  [경고] 성과 문항 인덱스를 찾을 수 없습니다.\n');
        continue;
    end
    
    % 성과 종합점수 계산
    performanceData = questionData(:, perfIndices);
    
    % 리커트 척도 표준화 적용
    standardizedPerformanceData = standardizeQuestionScales(performanceData, availableQuestions, p);
    
    % 각 응답자별 성과점수 계산 (표준화된 값의 평균)
    performanceScores = nanmean(standardizedPerformanceData, 2);
    
    % 결측치가 너무 많은 응답자 제외 (50% 이상 결측시)
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
    
    % 전체 테이블에 추가
    allPerformanceData = [allPerformanceData; tempTable];
end

% 개인별로 성과점수 평균 계산 (중복 ID 제거)
fprintf('\n▶ 개인별 종합성과점수 통합 중...\n');
uniqueIDs = unique(allPerformanceData.ID);
integratedPerformanceScores = [];
integratedIDs = {};

for i = 1:length(uniqueIDs)
    personID = uniqueIDs{i};
    personData = allPerformanceData(strcmp(allPerformanceData.ID, personID), :);
    
    % 해당 개인의 모든 시점 성과점수 평균 계산
    avgPerformanceScore = nanmean(personData.PerformanceScore);
    numPeriods = height(personData);
    
    integratedPerformanceScores(end+1) = avgPerformanceScore;
    integratedIDs{end+1} = personID;
    
    if mod(i, 20) == 0  % 20명마다 진행상황 출력
        fprintf('  진행: %d/%d명 처리 완료\n', i, length(uniqueIDs));
    end
end

fprintf('  - 고유한 개인 수: %d명\n', length(uniqueIDs));
fprintf('  - 평균 참여 시점: %.1f개\n', height(allPerformanceData) / length(uniqueIDs));

% 종합성과 히스토그램 생성 (고유한 개인별)
if ~isempty(integratedPerformanceScores)
    figure('Name', '전체 종합성과점수 분포', 'Position', [500, 500, 800, 600]);
    
    validScores = integratedPerformanceScores(~isnan(integratedPerformanceScores));
    
    histogram(validScores, 30);
    title('전체 종합성과점수 분포 (고유한 개인별)', 'FontSize', 14, 'FontWeight', 'bold');
    xlabel('종합성과점수 (표준화된 값)', 'FontSize', 12);
    ylabel('빈도', 'FontSize', 12);
    grid on;
    
    meanScore = mean(validScores);
    stdScore = std(validScores);
    avgParticipation = height(allPerformanceData) / length(uniqueIDs);
    
    % 통계 정보를 텍스트 박스로 표시
    textStr = sprintf('평균: %.3f\n표준편차: %.3f\nN: %d명 (고유한 개인)\n평균 참여: %.1f회\n범위: %.3f ~ %.3f', ...
                     meanScore, stdScore, length(validScores), avgParticipation, ...
                     min(validScores), max(validScores));
    text(0.02, 0.98, textStr, 'Units', 'normalized', 'FontSize', 11, ...
         'BackgroundColor', 'white', 'EdgeColor', 'black', ...
         'VerticalAlignment', 'top', 'HorizontalAlignment', 'left');
    
    fprintf('✓ 전체 종합성과점수 히스토그램 생성 완료: %d명 (고유한 개인)\n', length(validScores));
else
    fprintf('⚠️  종합성과점수 데이터가 없습니다.\n');
end

%% 3. 각 Period별 상관 분석
fprintf('\n[3단계] 각 Period별 문항 데이터 분석\n');
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
    
    % 문항 컬럼들 추출 (Q로 시작하는 숫자형 컬럼)
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
        fprintf('\n');  % 구분을 위한 빈 줄
        perfQuestions = performanceQuestions.(sprintf('period%d', p));
        displayPerformanceQuestionCorrelations(correlationMatrix, pValues, questionCols, perfQuestions);
    end
end

%% 4. 결과 저장
fprintf('\n[4단계] 결과 저장\n');
fprintf('----------------------------------------\n');

% 백업 폴더 확인 및 생성
backupDir = 'D:\project\HR데이터\결과\성과종합점수&역검_weighted\backup';
if ~exist(backupDir, 'dir')
    mkdir(backupDir);
    fprintf('✓ 백업 폴더 생성: %s\n', backupDir);
end

% 기존 파일들을 백업 폴더로 이동
existingFiles = dir('D:\project\HR데이터\결과\성과종합점수&역검_weighted\correlation_matrices_by_period_*.xlsx');
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

% 새 결과 파일명 생성 (최신 파일)
dateStr = datestr(now, 'yyyymmdd_HHMM');
outputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검_weighted\\correlation_matrices_by_period_%s.xlsx', dateStr);

% 각 Period별 상관 매트릭스를 별도 시트에 저장
savedSheets = {};
periodFields = fieldnames(correlationMatrices);

for i = 1:length(periodFields)
    fieldName = periodFields{i};
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

%% 5. 요약 테이블 생성 및 저장
% correlationMatrices가 비어있지 않은지 확인
if ~isempty(fieldnames(correlationMatrices))
    summaryTable = createSummaryTable(correlationMatrices, periods);
    
    try
        writetable(summaryTable, outputFileName, 'Sheet', '분석요약');
        savedSheets{end+1} = '분석요약';
        fprintf('✓ 분석 요약 저장 완료\n');
    catch ME
        fprintf('✗ 분석 요약 저장 실패: %s\n', ME.message);
    end
else
    fprintf('⚠️ 상관 매트릭스가 비어있어 요약 테이블을 생성할 수 없습니다.\n');
    summaryTable = table();  % 빈 테이블 생성
end

% MAT 파일로도 저장
matFileName = sprintf('D:\\project\\correlation_matrices_workspace_%s.mat', dateStr);
if ~isempty(fieldnames(correlationMatrices))
    save(matFileName, 'correlationMatrices', 'periods', 'summaryTable', 'allData');
else
    save(matFileName, 'periods', 'allData');  % correlationMatrices가 없을 때
end

%% 6. 분포 기반 시각화는 성과점수 분석 이후에 수행

%% 7. 최종 요약 출력
fprintf('\n========================================\n');
fprintf('Period별 상관 매트릭스 생성 완료\n');
fprintf('========================================\n');

if ~isempty(fieldnames(correlationMatrices))
    fprintf('📊 처리된 Period 수: %d개\n', length(fieldnames(correlationMatrices)));
else
    fprintf('⚠️ 처리된 Period가 없습니다.\n');
end

fprintf('📁 저장된 파일:\n');
fprintf('  • Excel: %s\n', outputFileName);
fprintf('  • MAT: %s\n', matFileName);

fprintf('\n📋 저장된 시트:\n');
for i = 1:length(savedSheets)
    fprintf('  • %s\n', savedSheets{i});
end

fprintf('\n📈 Period별 상관 매트릭스 크기:\n');
if exist('summaryTable', 'var') && height(summaryTable) > 0
    for i = 1:height(summaryTable)
        fprintf('  • %s (샘플: %d명, 문항: %d개)\n', ...
            summaryTable.Period{i}, summaryTable.SampleSize(i), summaryTable.NumQuestions(i));
    end
else
    fprintf('  (요약 테이블이 비어있습니다)\n');
end

%% 8. 성과 관련 문항 종합점수 분석 (역량검사와의 상관분석)
fprintf('\n[5단계] 성과 관련 문항 종합점수 분석 (역량검사와의 상관분석)\n');
fprintf('========================================\n');

performanceResults = struct();

for p = 1:length(periods)
    fprintf('\n▶ %s 성과점수 분석 중...\n', periods{p});
    
    % 해당 period의 상관 매트릭스가 있는지 확인
    if ~isfield(correlationMatrices, sprintf('period%d', p))
        fprintf('  [경고] %s 상관 매트릭스를 찾을 수 없습니다.\n', periods{p});
        continue;
    end
    
    result = correlationMatrices.(sprintf('period%d', p));
    perfQuestions = performanceQuestions.(sprintf('period%d', p));
    
    % 성과 관련 문항이 실제 데이터에 있는지 확인
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
    
    % 성과 관련 문항의 인덱스 찾기
    perfIndices = [];
    for q = 1:length(availableQuestions)
        qIdx = find(strcmp(result.questionNames, availableQuestions{q}));
        if ~isempty(qIdx)
            perfIndices(end+1) = qIdx;
        end
    end
    
    if length(perfIndices) < 3
        fprintf('  [경고] 성과 문항 인덱스를 찾을 수 없습니다.\n');
        continue;
    end
    
    % 성과 종합점수 계산 (리커트 척도 표준화 적용)
    performanceData = result.cleanData(:, perfIndices);
    
    % 각 문항별로 리커트 척도 정보 가져오기 및 표준화
    standardizedPerformanceData = standardizeQuestionScales(performanceData, availableQuestions, p);
    
    % 각 응답자별 성과점수 계산 (표준화된 값의 평균)
    performanceScores = nanmean(standardizedPerformanceData, 2);
    
    % 결측치가 너무 많은 응답자 제외 (50% 이상 결측시)
    validPerformanceRows = sum(~isnan(performanceData), 2) >= (length(perfIndices) * 0.5);
    cleanPerformanceScores = performanceScores(validPerformanceRows);
    cleanAllData = result.cleanData(validPerformanceRows, :);
    
    fprintf('  - 성과점수 계산 완료: %d명 (평균: %.2f, 표준편차: %.2f)\n', ...
        length(cleanPerformanceScores), mean(cleanPerformanceScores), std(cleanPerformanceScores));
    
    % 역량검사 종합점수와 성과점수 간 상관분석
    % cleanAllData에서 역량검사 종합점수 추출 (마지막 컬럼)
    rawCompetencyTestScores = cleanAllData(:, end);  % CompetencyTest_Total
    
    % 역량검사 점수도 min-max 스케일링 적용
    competencyTestScores = standardizeCompetencyScores(rawCompetencyTestScores);
    
    % 역량검사점수와 성과점수만으로 상관분석
    performanceCorrelationData = [competencyTestScores, cleanPerformanceScores];
    
    % 결측치가 있는 행 제거
    validRows = ~any(isnan(performanceCorrelationData), 2);
    cleanPerformanceCorrelationData = performanceCorrelationData(validRows, :);
    
    if size(cleanPerformanceCorrelationData, 1) < 3
        fprintf('  [경고] 역량검사-성과점수 상관분석을 위한 데이터가 부족합니다 (%d명)\n', ...
            size(cleanPerformanceCorrelationData, 1));
        continue;
    end
    
    try
        % 상관계수 계산
        [corrCoeff, pValue] = corr(cleanPerformanceCorrelationData(:,1), cleanPerformanceCorrelationData(:,2));
        
        % 성과점수 분석 결과 저장 (ID 정보도 포함)
        % validRows는 performanceCorrelationData에 대한 인덱스이므로, validPerformanceRows와 결합 필요
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
        
        % 상관분석 결과 출력
        sig_str = '';
        if pValue < 0.001, sig_str = '***';
        elseif pValue < 0.01, sig_str = '**';
        elseif pValue < 0.05, sig_str = '*';
        end
        
        fprintf('  ✓ 역량검사점수 vs 성과점수 상관분석 완료\n');
        fprintf('    → r = %.3f (p = %.3f) %s (N = %d)\n', ...
            corrCoeff, pValue, sig_str, size(cleanPerformanceCorrelationData, 1));
        fprintf('    → 역량검사점수: 평균 %.2f (SD %.2f)\n', ...
            mean(cleanPerformanceCorrelationData(:,1)), std(cleanPerformanceCorrelationData(:,1)));
        fprintf('    → 성과점수: 평균 %.2f (SD %.2f)\n', ...
            mean(cleanPerformanceCorrelationData(:,2)), std(cleanPerformanceCorrelationData(:,2)));
        
    catch ME
        fprintf('  [오류] 성과점수 상관 분석 실패: %s\n', ME.message);
    end
end

%% 9. 성과점수 분석 결과 저장
if ~isempty(fieldnames(performanceResults))
    fprintf('\n[6단계] 역량검사-성과점수 상관분석 결과 저장\n');
    fprintf('----------------------------------------\n');
    
    % 성과점수 분석 요약 테이블 생성
    perfSummaryTable = createPerformanceSummaryTable(performanceResults, periods);
    
    try
        writetable(perfSummaryTable, outputFileName, 'Sheet', '역량검사_성과점수_상관분석');
        fprintf('✓ 역량검사-성과점수 상관분석 결과 저장 완료\n');
    catch ME
        fprintf('✗ 역량검사-성과점수 상관분석 저장 실패: %s\n', ME.message);
    end
    
    % MAT 파일에 성과점수 결과도 추가 저장
    save(matFileName, 'correlationMatrices', 'periods', 'summaryTable', 'allData', ...
         'performanceResults', 'performanceQuestions', 'perfSummaryTable', '-append');
    
    fprintf('\n📊 역량검사-성과점수 상관분석 완료 - %d개 시점 처리됨\n', length(fieldnames(performanceResults)));
else
    fprintf('\n⚠️  성과점수 분석 결과가 없습니다.\n');
end

%% 10. 종합 성과점수 분석 (5개 시점 통합)
fprintf('\n[7단계] 종합 성과점수 분석 (5개 시점 통합)\n');
fprintf('========================================\n');

if ~isempty(fieldnames(performanceResults))
    % 각 개인별로 5개 시점의 성과점수를 통합
    [integratedPerformanceData, overallCorrelation] = integratePerformanceScores(performanceResults, competencyTestData, periods, correlationMatrices);
    
    if ~isempty(integratedPerformanceData)
        fprintf('✅ 종합 성과점수 분석 완료\n');
        fprintf('   → 전체 상관계수: r = %.3f (p = %.3f) %s (N = %d)\n', ...
            overallCorrelation.correlation, overallCorrelation.pValue, ...
            overallCorrelation.significance, overallCorrelation.sampleSize);
        
        % 종합 분석 결과를 Excel에 저장
        try
            writetable(integratedPerformanceData, outputFileName, 'Sheet', '종합성과점수분석');
            fprintf('✓ 종합 성과점수 분석 결과 저장 완료\n');
        catch ME
            fprintf('✗ 종합 성과점수 분석 저장 실패: %s\n', ME.message);
        end
        
        % MAT 파일에도 추가 저장
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

% 디버깅 정보
fprintf('▶ 시각화 생성 조건 확인:\n');
fprintf('  - correlationMatrices 필드 수: %d\n', length(fieldnames(correlationMatrices)));
fprintf('  - performanceResults 필드 수: %d\n', length(fieldnames(performanceResults)));
if exist('integratedPerformanceData', 'var')
    if istable(integratedPerformanceData)
        fprintf('  - integratedPerformanceData: %d행 테이블\n', height(integratedPerformanceData));
    else
        fprintf('  - integratedPerformanceData: 비어있음\n');
    end
else
    fprintf('  - integratedPerformanceData: 변수 없음\n');
end

if length(fieldnames(correlationMatrices)) > 0
    periodFields = fieldnames(correlationMatrices);

    % 시각화 함수 호출 전에 성과점수 분석 상태 확인
    if isempty(fieldnames(performanceResults))
        fprintf('⚠️  성과점수 분석 결과가 없어 일부 시각화가 생성되지 않습니다.\n');
        fprintf('   → 기본 시각화(역량검사점수 분포)만 생성됩니다.\n');
    end

    createDistributionVisualizations(correlationMatrices, periods, periodFields, performanceResults, integratedPerformanceData, overallCorrelation, competencyTestData);
else
    fprintf('❌ 상관 매트릭스가 없어 시각화를 생성할 수 없습니다.\n');
end

% 원래 디렉토리로 복귀

fprintf('\n========================================\n');
fprintf('전체 분석 완료\n');
fprintf('========================================\n');

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
            return;
        end
    end
end

function [questionCols, questionData] = extractQuestionData(selfData, idCol)

    colNames = selfData.Properties.VariableNames;
    questionCols = colNames(startsWith(colNames, 'Q') & ~endsWith(colNames, '_text'));
    
    if isempty(questionCols)
        questionData = [];
        return;
    end
    
    questionData = selfData{:, questionCols};
    
    % cell 배열인 경우 숫자로 변환
    if iscell(questionData)
        questionData = cell2mat(cellfun(@(x) str2double(x), questionData, 'UniformOutput', false));
    end
end

function standardizedIDs = extractAndStandardizeIDs(rawIDs)

    if isnumeric(rawIDs)
        standardizedIDs = string(rawIDs);
    elseif iscell(rawIDs)
        standardizedIDs = cellfun(@(x) char(x), rawIDs, 'UniformOutput', false);
    else
        standardizedIDs = cellstr(rawIDs);
    end
    
    % 빈 값이나 NaN 처리
    emptyIdx = cellfun(@isempty, standardizedIDs) | strcmp(standardizedIDs, 'NaN');
    standardizedIDs(emptyIdx) = {''};
end

function [competencyScore, usedColumnName] = findCompetencyScoreColumn(competencyTestData, testIdx)

    % 역량검사 종합점수 컬럼을 찾는 개선된 함수
    competencyScore = [];
    usedColumnName = '';
    
    if isempty(competencyTestData) || height(competencyTestData) == 0
        fprintf('    [경고] 역량검사 데이터가 비어있습니다\n');
        return;
    end
    
    colNames = competencyTestData.Properties.VariableNames;
    fprintf('    역량검사 데이터 컬럼 수: %d개\n', length(colNames));
    
    % 1단계: 키워드 기반 검색
    keywords = {'종합', '총점', 'total', 'sum', 'score', '점수', '가중', 'weighted'};
    candidateColumns = {};
    candidateScores = [];
    
    for i = 1:length(colNames)
        colName = lower(colNames{i});
        score = 0;
        
        for j = 1:length(keywords)
            if contains(colName, keywords{j})
                score = score + 1;
            end
        end
        
        if score > 0
            candidateColumns{end+1} = colNames{i};
            candidateScores(end+1) = score;
        end
    end
    
    if ~isempty(candidateColumns)
        [~, bestIdx] = max(candidateScores);
        bestCol = candidateColumns{bestIdx};
        colData = competencyTestData{:, bestCol};
        
        if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
            competencyScore = colData;
            usedColumnName = bestCol;
            fprintf('    사용된 역량점수 컬럼: "%s" (키워드 매치)\n', usedColumnName);
            return;
        end
    end
    
    % 2단계: 숫자형 컬럼 중 가장 적절한 것 찾기
    fprintf('    키워드 매치 실패 - 숫자형 컬럼 탐색 중...\n');
    
    candidateColumns = {};
    candidateScores = [];
    
    for i = 1:length(colNames)
        colData = competencyTestData{:, i};
        
        if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
            % 통계적 특성으로 점수 매기기
            score = 0;
            
            % 평균이 높은 경우 (0-100 척도일 가능성)
            meanVal = mean(colData, 'omitnan');
            if meanVal > 50 && meanVal < 100
                score = score + 2;
            end
            
            % 표준편차가 적당한 경우 (너무 균일하지 않음)
            stdVal = std(colData, 'omitnan');
            if stdVal > 5 && stdVal < 30
                score = score + 1;
            end
            
            % 범위가 적당한 경우
            rangeVal = max(colData) - min(colData);
            if rangeVal > 20 && rangeVal < 100
                score = score + 1;
            end
            
            if score > 0
                candidateColumns{end+1} = colNames{i};
                candidateScores(end+1) = score;
            end
        end
    end
    
    if ~isempty(candidateColumns)
        [~, bestIdx] = max(candidateScores);
        bestCol = candidateColumns{bestIdx};
        competencyScore = competencyTestData{:, bestCol};
        usedColumnName = bestCol;
        fprintf('    사용된 역량점수 컬럼: "%s" (통계적 특성 기반)\n', usedColumnName);
        return;
    end
    
    % 3단계: 마지막 숫자형 컬럼 사용
    fprintf('    통계적 특성 매치 실패 - 마지막 숫자형 컬럼 사용\n');
    
    for i = length(colNames):-1:1
        colData = competencyTestData{:, i};
        if isnumeric(colData) && ~all(isnan(colData))
            competencyScore = colData;
            usedColumnName = colNames{i};
            fprintf('    사용된 역량점수 컬럼: "%s" (마지막 숫자형 컬럼)\n', usedColumnName);
            return;
        end
    end
    
    fprintf('    [경고] 적절한 역량점수 컬럼을 찾을 수 없습니다\n');
end

function [matchedData, matchedIDs, sampleSize] = matchWithCompetencyTest(questionData, responseIDs, competencyTestData)

    % ID를 문자열로 통일
    if isnumeric(responseIDs)
        responseIDs = string(responseIDs);
    elseif iscell(responseIDs)
        responseIDs = cellfun(@(x) char(x), responseIDs, 'UniformOutput', false);
    else
        responseIDs = cellstr(responseIDs);
    end
    
    % 역량검사 데이터의 ID 컬럼 찾기
    competencyIDCol = findIDColumn(competencyTestData);
    if isempty(competencyIDCol)
        fprintf('    [경고] 역량검사 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        matchedData = [];
        matchedIDs = [];
        sampleSize = 0;
        return;
    end
    
    % 역량검사 ID 추출 및 표준화
    competencyIDs = extractAndStandardizeIDs(competencyTestData{:, competencyIDCol});
    
    % ID 매칭
    [commonIDs, questionIdx, competencyIdx] = intersect(responseIDs, competencyIDs);
    
    if isempty(commonIDs)
        fprintf('    [경고] 매칭되는 ID가 없습니다\n');
        matchedData = [];
        matchedIDs = [];
        sampleSize = 0;
        return;
    end
    
    % 매칭된 데이터 구성
    matchedData = questionData(questionIdx, :);
    matchedIDs = commonIDs;
    sampleSize = length(commonIDs);
    
    fprintf('    ✓ ID 매칭 완료: %d명\n', sampleSize);
end

function [correlationMatrix, pValues, cleanData, variableNames] = performCorrelationAnalysis(matchedData, questionCols)
    if isempty(matchedData)
        correlationMatrix = [];
        pValues = [];
        cleanData = [];
        variableNames = {};
        return;
    end
    
    % 문항 데이터 추출
    colNames = matchedData.Properties.VariableNames;
    questionCols = colNames(startsWith(colNames, 'Q') & ~endsWith(colNames, '_text'));
    
    if isempty(questionCols)
        correlationMatrix = [];
        pValues = [];
        cleanData = [];
        variableNames = {};
        return;
    end
    
    questionData = matchedData{:, questionCols};
    
    % cell 배열인 경우 숫자로 변환
    if iscell(questionData)
        questionData = cell2mat(cellfun(@(x) str2double(x), questionData, 'UniformOutput', false));
    end
    
    % 역량검사 점수 컬럼 찾기
    competencyScore = findCompetencyScoreColumn(matchedData, 1);
    
    if isempty(competencyScore)
        correlationMatrix = [];
        pValues = [];
        cleanData = [];
        variableNames = {};
        return;
    end
    
    % 상관분석 수행
    [correlationMatrix, pValues] = corr(questionData, competencyScore, 'rows', 'pairwise');
    
    % 결과 정리
    cleanData = questionData;
    variableNames = questionCols;
end

function displayTopCorrelations(correlationMatrix, pValues, questionCols)
    
    if size(correlationMatrix, 2) < 2
        fprintf('  상관분석 결과가 부족합니다\n');
        return;
    end
    
    % 상관계수 절댓값으로 정렬
    [sortedCorrs, sortedIdx] = sort(abs(correlationMatrix), 'descend');
    
    fprintf('  상위 5개 문항의 종합점수와의 상관:\n');
    for i = 1:min(5, length(sortedIdx))
        idx = sortedIdx(i);
        r = correlationMatrix(idx);
        p = pValues(idx);
        
        % 유의성 표시
        if p < 0.001
            sig = '***';
        elseif p < 0.01
            sig = '**';
        elseif p < 0.05
            sig = '*';
        else
            sig = '';
        end
        
        fprintf('    %s: r=%.3f (p=%.3f) %s\n', questionCols{idx}, r, p, sig);
    end
end

function displayPerformanceQuestionCorrelations(correlationMatrix, pValues, questionCols, performanceQuestions)
    if size(correlationMatrix, 2) < 2
        fprintf('  상관분석 결과가 부족합니다\n');
        return;
    end
    
    % 성과 관련 문항 찾기
    if isempty(performanceQuestions)
        fprintf('  성과 관련 문항이 정의되지 않았습니다\n');
        return;
    end
    
    fprintf('  성과 관련 문항의 역량검사 종합점수와의 상관:\n');
    for i = 1:length(performanceQuestions)
        qName = performanceQuestions{i};
        qIdx = find(strcmp(questionCols, qName));
        
        if ~isempty(qIdx)
            r = correlationMatrix(qIdx);
            p = pValues(qIdx);
            
            % 유의성 표시
            if p < 0.001
                sig = '***';
            elseif p < 0.01
                sig = '**';
            elseif p < 0.05
                sig = '*';
            else
                sig = '';
            end
            
            fprintf('    %s: r=%.3f (p=%.3f) %s\n', qName, r, p, sig);
        else
            fprintf('    %s: 문항을 찾을 수 없음\n', qName);
        end
    end
end

function createDistributionVisualizations(correlationMatrices, periods, periodFields, performanceResults, integratedPerformanceData, overallCorrelation, competencyTestData)
                
    if isempty(periodFields)
        fprintf('❌ 시각화를 생성할 수 없습니다: Period 데이터가 없습니다\n');
        return;
    end
    
    fprintf('▶ 분포 기반 시각화 생성 중...\n');
    
    % 1. 역량검사점수 분포 히스토그램
    figure('Name', '역량검사점수 분포', 'Position', [100, 100, 1400, 900]);
    
    numPeriods = length(periodFields);
    subplotRows = ceil(numPeriods / 2);
    subplotCols = 2;
    
    for periodNum = 1:numPeriods
        periodField = periodFields{periodNum};
        
        if isfield(correlationMatrices, periodField)
            periodData = correlationMatrices.(periodField);
            if isfield(periodData, 'competencyScores') && ~isempty(periodData.competencyScores)
                subplot(subplotRows, subplotCols, periodNum);
                histogram(periodData.competencyScores, 20);
                title(sprintf('%s 역량검사점수 분포', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
                xlabel('역량검사점수');
                ylabel('빈도');
                grid on;
            end
        end
    end
    
    fprintf('✓ 역량검사점수 분포 히스토그램 생성 완료\n');
    
    % 2. 성과점수 vs 역량검사점수 산점도
    if ~isempty(performanceResults)
        figure('Name', '성과점수 vs 역량검사점수', 'Position', [200, 200, 1400, 900]);
        
        for periodNum = 1:numPeriods
            periodField = periodFields{periodNum};
            
            if isfield(performanceResults, periodField) && isfield(correlationMatrices, periodField)
                periodPerfData = performanceResults.(periodField);
                periodCorrData = correlationMatrices.(periodField);
                
                if ~isempty(periodPerfData) && isfield(periodCorrData, 'competencyScores')
                    subplot(subplotRows, subplotCols, periodNum);
                    scatter(periodCorrData.competencyScores, periodPerfData{:, 2}, 'filled', 'Alpha', 0.6);
                    title(sprintf('%s: 역량검사 vs 성과점수', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
                    xlabel('역량검사점수');
                    ylabel('성과점수');
                    grid on;
                    
                    % 상관계수 계산 및 표시
                    [r, p] = corr(periodCorrData.competencyScores, periodPerfData{:, 2}, 'rows', 'pairwise');
                    text(0.05, 0.95, sprintf('r=%.3f, p=%.3f', r, p), 'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white');
                end
            end
        end
        
        fprintf('✓ 성과점수 vs 역량검사점수 산점도 생성 완료\n');
    end
    
    % 3. 통합 성과점수 vs 역량검사점수
    if ~isempty(integratedPerformanceData) && ~isempty(overallCorrelation)
        figure('Name', '통합 성과점수 vs 역량검사점수', 'Position', [300, 300, 800, 600]);
        
        scatter(integratedPerformanceData{:, 2}, integratedPerformanceData{:, 3}, 'filled', 'Alpha', 0.6);
        title('통합 성과점수 vs 역량검사점수', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('통합 성과점수');
        ylabel('역량검사점수');
        grid on;
        
        % 상관계수 표시
        text(0.05, 0.95, sprintf('r=%.3f', overallCorrelation), 'Units', 'normalized', 'FontSize', 12, 'BackgroundColor', 'white');
        
        fprintf('✓ 통합 성과점수 vs 역량검사점수 산점도 생성 완료\n');
    end
    
    fprintf('✓ 모든 시각화 생성 완료\n');
end

function standardizedData = standardizeQuestionScales(questionData, questionNames, periodNum)
    % 문항 척도 정보 가져오기
    scaleInfo = getQuestionScaleInfo(questionNames, periodNum);
    
    % 표준화 수행
    standardizedData = questionData;
    
    for i = 1:length(questionNames)
        qName = questionNames{i};
        
        if isfield(scaleInfo, qName)
            scale = scaleInfo.(qName);
            originalData = questionData(:, i);
            
            % 결측치가 아닌 데이터만 처리
            validIdx = ~isnan(originalData);
            if sum(validIdx) > 0
                % min-max 표준화
                minVal = scale.min;
                maxVal = scale.max;
                
                if maxVal > minVal
                    standardizedData(validIdx, i) = (originalData(validIdx) - minVal) / (maxVal - minVal);
                else
                    standardizedData(validIdx, i) = 0.5; % 모든 값이 같으면 중간값으로 설정
                end
            end
        end
    end
end

function scaleInfo = getQuestionScaleInfo(questionNames, periodNum)
    % 기본 척도 정보 가져오기
    defaultScales = getDefaultQuestionScales(periodNum);
    
    % 문항별 척도 정보 구성
    scaleInfo = struct();
    
    for i = 1:length(questionNames)
        qName = questionNames{i};
        
        % 기본 척도에서 찾기
        if isfield(defaultScales, qName)
            scaleInfo.(qName) = defaultScales.(qName);
        else
            % 기본값 설정 (1-5 척도)
            scaleInfo.(qName) = struct('min', 1, 'max', 5);
        end
    end
end

function defaultScales = getDefaultQuestionScales(periodNum)
    % 시점별 기본 척도 정보
    defaultScales = struct();
    
    % 모든 문항에 대해 기본 1-5 척도 설정
    commonQuestions = {'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', ...
                      'Q11', 'Q12', 'Q13', 'Q14', 'Q15', 'Q16', 'Q17', 'Q18', 'Q19', 'Q20', ...
                      'Q21', 'Q22', 'Q23', 'Q24', 'Q25', 'Q26', 'Q27', 'Q28', 'Q29', 'Q30', ...
                      'Q31', 'Q32', 'Q33', 'Q34', 'Q35', 'Q36', 'Q37', 'Q38', 'Q39', 'Q40', ...
                      'Q41', 'Q42', 'Q43', 'Q44', 'Q45', 'Q46', 'Q47', 'Q48', 'Q49', 'Q50', ...
                      'Q51', 'Q52', 'Q53', 'Q54', 'Q55', 'Q56', 'Q57', 'Q58', 'Q59', 'Q60'};
    
    for i = 1:length(commonQuestions)
        qName = commonQuestions{i};
        defaultScales.(qName) = struct('min', 1, 'max', 5);
    end
end

function standardizedScores = standardizeCompetencyScores(rawScores)
    % 역량검사 점수 min-max 표준화 함수
    if isempty(rawScores) || all(isnan(rawScores))
        standardizedScores = [];
        return;
    end
    
    % NaN 값 제거
    validScores = rawScores(~isnan(rawScores));
    
    if isempty(validScores)
        standardizedScores = [];
        return;
    end
    
    % min-max 표준화
    minScore = min(validScores);
    maxScore = max(validScores);
    
    if maxScore > minScore
        standardizedScores = (rawScores - minScore) / (maxScore - minScore);
    else
        standardizedScores = zeros(size(rawScores));
    end
    
    fprintf('    ✓ 역량검사 점수 표준화 완료 (원본 범위: %.1f~%.1f → [0,1])\n', minScore, maxScore);
end

function perfSummaryTable = createPerformanceSummaryTable(performanceResults, periods)
    % 역량검사-성과점수 상관분석 요약 테이블을 생성하는 함수
    if isempty(performanceResults)
        perfSummaryTable = table();
        return;
    end
    
    % 요약 데이터 수집
    summaryData = {};
    resultFields = fieldnames(performanceResults);
    
    for i = 1:length(resultFields)
        fieldName = resultFields{i};
        result = performanceResults.(fieldName);
        
        if ~isempty(result) && isfield(result, 'correlation') && isfield(result, 'pValue')
            summaryRow = {fieldName, result.sampleSize, result.correlation, result.pValue, ...
                         result.performanceMean, result.performanceStd, result.competencyMean, result.competencyStd};
            summaryData = [summaryData; summaryRow];
        end
    end
    
    % 테이블 생성
    if ~isempty(summaryData)
        perfSummaryTable = cell2table(summaryData, 'VariableNames', ...
            {'Period', 'SampleSize', 'Correlation', 'PValue', 'PerformanceMean', 'PerformanceStd', 'CompetencyMean', 'CompetencyStd'});
    else
        perfSummaryTable = table();
    end
end

function [integratedData, overallCorrelation] = integratePerformanceScores(performanceResults, competencyTestData, periods, correlationMatrices)
    % 5개 시점의 성과점수를 개인별로 통합하고 역량검사점수와 상관분석하는 함수
    integratedData = [];
    overallCorrelation = [];
    
    if isempty(performanceResults) || isempty(competencyTestData)
        fprintf('  [경고] 통합 분석을 위한 데이터가 부족합니다\n');
        return;
    end
    
    fprintf('  ▶ 5개 시점 성과점수 통합 분석 중...\n');
    
    % 역량검사 데이터에서 ID 추출
    competencyIDCol = findIDColumn(competencyTestData);
    if isempty(competencyIDCol)
        fprintf('  [경고] 역량검사 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        return;
    end
    
    competencyIDs = extractAndStandardizeIDs(competencyTestData{:, competencyIDCol});
    
    % 역량검사 점수 추출
    [competencyScores, usedColumnName] = findCompetencyScoreColumn(competencyTestData, 1:height(competencyTestData));
    if isempty(competencyScores)
        fprintf('  [경고] 역량검사 점수를 찾을 수 없습니다\n');
        return;
    end
    
    fprintf('  - 사용된 역량점수 컬럼: "%s"\n', usedColumnName);
    
    % 각 시점의 성과점수 수집
    allPerformanceData = [];
    resultFields = fieldnames(performanceResults);
    
    for i = 1:length(resultFields)
        fieldName = resultFields{i};
        result = performanceResults.(fieldName);
        
        if ~isempty(result) && isfield(result, 'performanceScores') && isfield(result, 'sampleIDs')
            % 성과점수와 ID 매칭
            perfIDs = result.sampleIDs;
            perfScores = result.performanceScores;
            
            % ID 표준화
            perfIDs = extractAndStandardizeIDs(perfIDs);
            
            % 역량검사 ID와 매칭
            [commonIDs, perfIdx, compIdx] = intersect(perfIDs, competencyIDs);
            
            if ~isempty(commonIDs)
                matchedPerfScores = perfScores(perfIdx);
                matchedCompScores = competencyScores(compIdx);
                
                % 데이터 추가
                for j = 1:length(commonIDs)
                    allPerformanceData = [allPerformanceData; {commonIDs{j}, matchedPerfScores(j), matchedCompScores(j)}];
                end
            end
        end
    end
    
    if isempty(allPerformanceData)
        fprintf('  [경고] 매칭되는 성과점수 데이터가 없습니다\n');
        return;
    end
    
    % 개인별로 성과점수 평균 계산
    uniqueIDs = unique(allPerformanceData(:, 1));
    integratedScores = [];
    
    for i = 1:length(uniqueIDs)
        id = uniqueIDs{i};
        idRows = strcmp(allPerformanceData(:, 1), id);
        
        if sum(idRows) > 0
            perfScores = cell2mat(allPerformanceData(idRows, 2));
            compScores = cell2mat(allPerformanceData(idRows, 3));
            
            % 성과점수는 평균, 역량검사점수는 첫 번째 값 사용
            avgPerfScore = mean(perfScores);
            compScore = compScores(1);
            
            integratedScores = [integratedScores; {id, avgPerfScore, compScore}];
        end
    end
    
    if isempty(integratedScores)
        fprintf('  [경고] 통합된 성과점수 데이터가 없습니다\n');
        return;
    end
    
    % 테이블로 변환
    integratedData = cell2table(integratedScores, 'VariableNames', {'ID', 'PerformanceScore', 'CompetencyScore'});
    
    % 상관분석 수행
    perfScores = integratedData.PerformanceScore;
    compScores = integratedData.CompetencyScore;
    
    [r, p] = corr(perfScores, compScores, 'rows', 'pairwise');
    
    overallCorrelation = struct();
    overallCorrelation.correlation = r;
    overallCorrelation.pValue = p;
    overallCorrelation.sampleSize = height(integratedData);
    
    fprintf('  ✓ 통합 분석 완료: %d명, r=%.3f (p=%.3f)\n', height(integratedData), r, p);
end

function summaryTable = createSummaryTable(correlationMatrices, periods)
    % 상관분석 결과 요약 테이블 생성
    if isempty(correlationMatrices)
        summaryTable = table();
        return;
    end
    
    % 요약 데이터 수집
    summaryData = {};
    resultFields = fieldnames(correlationMatrices);
    
    for i = 1:length(resultFields)
        fieldName = resultFields{i};
        result = correlationMatrices.(fieldName);
        
        if ~isempty(result) && isfield(result, 'correlationMatrix') && isfield(result, 'sampleSize')
            % 상관분석 결과 요약
            corrMatrix = result.correlationMatrix;
            pValueMatrix = result.pValueMatrix;
            
            % 유의한 상관관계 개수 (p < 0.05)
            significantCorrs = sum(pValueMatrix < 0.05, 'all');
            totalCorrs = numel(corrMatrix);
            
            % 최대 상관계수
            maxCorr = max(abs(corrMatrix(:)));
            
            % 평균 상관계수
            meanCorr = mean(abs(corrMatrix(:)), 'omitnan');
            
            summaryRow = {fieldName, result.sampleSize, totalCorrs, significantCorrs, maxCorr, meanCorr};
            summaryData = [summaryData; summaryRow];
        end
    end
    
    % 테이블 생성
    if ~isempty(summaryData)
        summaryTable = cell2table(summaryData, 'VariableNames', ...
            {'Period', 'SampleSize', 'TotalCorrelations', 'SignificantCorrelations', 'MaxCorrelation', 'MeanCorrelation'});
    else
        summaryTable = table();
    end
end

% ID를 기반으로 Period 데이터와 상위항목 데이터를 매칭하는 함수
%
% 입력:
%   periodData - Period 데이터 (테이블)
%   upperCategoryResults - 상위항목 분석 결과 구조체
%
% 출력:
%   matchedData - 매칭된 Period 데이터
%   matchedUpperScores - 매칭된 상위항목 점수
%   matchedIdx - 매칭된 인덱스

    matchedData = [];
    matchedUpperScores = [];
    matchedIdx = [];
    
    if isempty(periodData) || isempty(upperCategoryResults)
        return;
    end
    
    % Period 데이터에서 ID 추출
    periodIDCol = findIDColumn(periodData);
    if isempty(periodIDCol)
        fprintf('  [경고] Period 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        return;
    end
    
    periodIDs = extractAndStandardizeIDs(periodData{:, periodIDCol});
    
    % 상위항목 데이터에서 ID 추출
    if isfield(upperCategoryResults, 'commonIDs')
        upperIDs = upperCategoryResults.commonIDs;
    else
        fprintf('  [경고] 상위항목 데이터에서 ID를 찾을 수 없습니다\n');
        return;
    end
    
    % ID 매칭
    [commonIDs, periodIdx, upperIdx] = intersect(periodIDs, upperIDs);
    
    if isempty(commonIDs)
        fprintf('  [경고] 매칭되는 ID가 없습니다\n');
        return;
    end
    
    % 매칭된 데이터 구성
    matchedData = periodData(periodIdx, :);
    matchedIdx = periodIdx;
    
    % 상위항목 점수 추출
    if isfield(upperCategoryResults, 'scoreMatrix')
        matchedUpperScores = upperCategoryResults.scoreMatrix(upperIdx, :);
    else
        fprintf('  [경고] 상위항목 점수 데이터를 찾을 수 없습니다\n');
        matchedUpperScores = [];
    end
    
    fprintf('  ✓ ID 매칭 완료: %d명\n', length(commonIDs));
end

    % 상위항목 점수와 성과점수 간 상관분석 및 중다회귀분석
    upperCategoryResults = [];
    
    if isempty(upperCategoryData) || isempty(performanceResults)
        fprintf('  [경고] 상위항목 분석을 위한 데이터가 부족합니다\n');
        return;
    end
    
    fprintf('  ▶ 상위항목 성과분석 시작\n');
    
    % ID 컬럼 찾기
    upperIDCol = findIDColumn(upperCategoryData);
    if isempty(upperIDCol)
        fprintf('  [경고] 상위항목 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        return;
    end
    
    % 상위항목 점수 컬럼 식별
    colNames = upperCategoryData.Properties.VariableNames;
    scoreColumns = {};
    scoreColumnNames = {};
    
    for i = 1:width(upperCategoryData)
        if i == upperIDCol
            continue;
        end
        
        colName = colNames{i};
        colData = upperCategoryData{:, i};
        
        if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
            scoreColumns{end+1} = i;
            scoreColumnNames{end+1} = colName;
        end
    end
    
    fprintf('  - 발견된 상위항목: %d개\n', length(scoreColumnNames));
    
    if length(scoreColumnNames) < 2
        fprintf('  [경고] 분석에 필요한 상위항목이 부족합니다\n');
        return;
    end
    
    % ID 표준화
    upperIDs = extractAndStandardizeIDs(upperCategoryData{:, upperIDCol});
    
    % 통합 성과점수가 있는 경우 사용, 없으면 개별 시점 성과점수 사용
    if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData) && ...
       ismember('PerformanceScore', integratedPerformanceData.Properties.VariableNames)
        
        % 통합 성과점수 사용
        performanceIDs = integratedPerformanceData.ID;
        performanceScores = integratedPerformanceData.PerformanceScore;
        
        fprintf('  - 통합 성과점수 사용: %d명\n', height(integratedPerformanceData));
        
    else
        % 개별 시점 성과점수 사용
        performanceIDs = {};
        performanceScores = [];
        
        resultFields = fieldnames(performanceResults);
        for i = 1:length(resultFields)
            result = performanceResults.(resultFields{i});
            if ~isempty(result) && isfield(result, 'sampleIDs') && isfield(result, 'performanceScores')
                performanceIDs = [performanceIDs; result.sampleIDs];
                performanceScores = [performanceScores; result.performanceScores];
            end
        end
        
        if isempty(performanceIDs)
            fprintf('  [경고] 성과점수 데이터를 찾을 수 없습니다\n');
            return;
        end
        
        fprintf('  - 개별 시점 성과점수 사용: %d명\n', length(performanceIDs));
    end
    
    % ID 매칭
    [commonIDs, upperIdx, perfIdx] = intersect(upperIDs, performanceIDs);
    
    if length(commonIDs) < 10
        fprintf('  [경고] 매칭되는 데이터가 부족합니다 (%d명)\n', length(commonIDs));
        return;
    end
    
    fprintf('  - 매칭된 데이터: %d명\n', length(commonIDs));
    
    % 매칭된 상위항목 데이터 구성
    upperScoreMatrix = [];
    for i = 1:length(scoreColumns)
        colIdx = scoreColumns{i};
        colData = upperCategoryData{upperIdx, colIdx};
        upperScoreMatrix = [upperScoreMatrix, colData];
    end
    
    matchedPerformanceScores = performanceScores(perfIdx);
    
    % 상위항목은 pairwise correlation으로 처리
    validPerformanceRows = ~isnan(matchedPerformanceScores);
    cleanPerformanceScores = matchedPerformanceScores(validPerformanceRows);
    cleanUpperMatrix = upperScoreMatrix(validPerformanceRows, :);
    cleanCommonIDs = commonIDs(validPerformanceRows);
    
    fprintf('  - 유효한 데이터: %d명\n', length(cleanCommonIDs));
    
    % 각 상위항목별 유효 데이터 수 확인
    for i = 1:size(cleanUpperMatrix, 2)
        validCount = sum(~isnan(cleanUpperMatrix(:, i)));
        fprintf('    %s: %d명 유효\n', scoreColumnNames{i}, validCount);
    end
    
    if length(cleanCommonIDs) < 10
        fprintf('  [경고] 유효한 데이터가 부족합니다\n');
        return;
    end
    
    %% 1. 상관분석
    fprintf('\n  ▶ 상위항목-성과점수 상관분석\n');
    
    correlationResults = [];
    pValues = [];
    
    for i = 1:length(scoreColumnNames)
        [r, p] = corr(cleanUpperMatrix(:, i), cleanPerformanceScores, 'rows', 'pairwise');
        correlationResults(end+1) = r;
        pValues(end+1) = p;
        
        fprintf('    %s: r=%.3f, p=%.3f', scoreColumnNames{i}, r, p);
        if p < 0.001
            fprintf(' ***');
        elseif p < 0.01
            fprintf(' **');
        elseif p < 0.05
            fprintf(' *');
        end
        fprintf('\n');
    end
    
    % 상관분석 결과 테이블 생성
    correlationTable = table(scoreColumnNames', correlationResults', pValues', ...
        'VariableNames', {'UpperCategory', 'Correlation', 'PValue'});
    
    %% 2. 중다회귀분석
    fprintf('\n  ▶ 중다회귀분석\n');
    
    try
        % 최소 70% 이상의 상위항목 데이터가 있는 행만 사용
        missingCount = sum(isnan(cleanUpperMatrix), 2);
        maxMissing = floor(size(cleanUpperMatrix, 2) * 0.3); % 30%까지 결측 허용
        regressionValidRows = missingCount <= maxMissing;
        
        regressionMatrix = cleanUpperMatrix(regressionValidRows, :);
        regressionScores = cleanPerformanceScores(regressionValidRows);
        
        if size(regressionMatrix, 1) < 10
            fprintf('    [경고] 회귀분석을 위한 데이터가 부족합니다\n');
            regressionTable = table();
        else
            % 중다회귀분석 수행
            [b, bint, r, rint, stats] = regress(regressionScores, [ones(size(regressionMatrix, 1), 1), regressionMatrix]);
            
            % 결과 테이블 생성
            regressionTable = table(scoreColumnNames', b(2:end), bint(2:end, 1), bint(2:end, 2), ...
                'VariableNames', {'UpperCategory', 'Coefficient', 'CI_Lower', 'CI_Upper'});
            
            fprintf('    R² = %.3f, F = %.3f, p = %.3f\n', stats(1), stats(2), stats(3));
        end
        
        % 다중공선성 확인
        corrMatrix = corrcoef(regressionMatrix, 'rows', 'pairwise');
        maxCorr = max(corrMatrix(:));
        fprintf('    - 상위항목 간 최대 상관: %.3f\n', maxCorr);
        
        if maxCorr > 0.9
            fprintf('    ⚠️ 높은 다중공선성 감지 (r > 0.9)\n');
        end
        
    catch ME
        fprintf('    [경고] 중다회귀분석 실패: %s\n', ME.message);
        regressionTable = table();
    end
    
    %% 3. 예측 정확도 분석
    fprintf('\n  ▶ 예측 정확도 분석\n');
    
    try
        if size(regressionMatrix, 1) >= 10
            % 교차검증을 위한 예측 정확도 계산
            cv = cvpartition(size(regressionMatrix, 1), 'HoldOut', 0.3);
            trainIdx = cv.training;
            testIdx = cv.test;
            
            trainX = regressionMatrix(trainIdx, :);
            trainY = regressionScores(trainIdx);
            testX = regressionMatrix(testIdx, :);
            testY = regressionScores(testIdx);
            
            % 훈련 데이터로 모델 학습
            [b, ~, ~, ~, ~] = regress(trainY, [ones(size(trainX, 1), 1), trainX]);
            
            % 테스트 데이터로 예측
            predictedY = [ones(size(testX, 1), 1), testX] * b;
            
            % 예측 정확도 계산
            mse = mean((testY - predictedY).^2);
            rmse = sqrt(mse);
            mae = mean(abs(testY - predictedY));
            
            % 상관계수
            [r, p] = corr(testY, predictedY, 'rows', 'pairwise');
            
            fprintf('    RMSE: %.3f, MAE: %.3f, r=%.3f (p=%.3f)\n', rmse, mae, r, p);
            
            % 예측 정확도 테이블 생성
            predictionTable = table({rmse}, {mae}, {r}, {p}, ...
                'VariableNames', {'RMSE', 'MAE', 'Correlation', 'PValue'});
        else
            fprintf('    [경고] 예측 정확도 분석을 위한 데이터가 부족합니다\n');
            predictionTable = table();
        end
        
    catch ME
        fprintf('    [경고] 예측 정확도 분석 실패: %s\n', ME.message);
        predictionTable = table();
    end
    
    %% 결과 저장
    upperCategoryResults = struct();
    upperCategoryResults.correlationTable = correlationTable;
    upperCategoryResults.regressionTable = regressionTable;
    upperCategoryResults.predictionTable = predictionTable;
    upperCategoryResults.scoreColumnNames = {scoreColumnNames};
    upperCategoryResults.commonIDs = cleanCommonIDs;
    upperCategoryResults.scoreMatrix = cleanUpperMatrix;
    
    fprintf('  ✓ 상위항목 분석 완료\n');
end

    % 상위항목 분석 결과 시각화
    if isempty(upperCategoryResults)
        fprintf('  [경고] 상위항목 분석 결과가 없어 시각화를 생성할 수 없습니다\n');
        return;
    end
    
    fprintf('  ▶ 상위항목 분석 시각화 생성 중...\n');
    
    % 1. 상위항목-성과점수 상관관계 막대그래프
    if isfield(upperCategoryResults, 'correlationTable') && ~isempty(upperCategoryResults.correlationTable)
        figure('Name', '상위항목-성과점수 상관관계', 'Position', [100, 100, 1200, 800]);
        
        corrTable = upperCategoryResults.correlationTable;
        correlations = corrTable.Correlation;
        pValues = corrTable.PValue;
        categoryNames = corrTable.UpperCategory;
        
        % 상관계수 막대그래프
        subplot(2, 2, 1);
        bar(correlations);
        title('상위항목-성과점수 상관계수', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('상위항목');
        ylabel('상관계수');
        xticklabels(categoryNames);
        xtickangle(45);
        grid on;
        
        % 유의성 표시
        for i = 1:length(correlations)
            if pValues(i) < 0.001
                text(i, correlations(i) + 0.05, '***', 'HorizontalAlignment', 'center', 'FontSize', 12, 'Color', 'red');
            elseif pValues(i) < 0.01
                text(i, correlations(i) + 0.05, '**', 'HorizontalAlignment', 'center', 'FontSize', 12, 'Color', 'red');
            elseif pValues(i) < 0.05
                text(i, correlations(i) + 0.05, '*', 'HorizontalAlignment', 'center', 'FontSize', 12, 'Color', 'red');
            end
        end
        
        % p-value 막대그래프
        subplot(2, 2, 2);
        bar(pValues);
        title('상위항목-성과점수 p값', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('상위항목');
        ylabel('p값');
        xticklabels(categoryNames);
        xtickangle(45);
        yline(0.05, 'r--', 'LineWidth', 2, 'DisplayName', 'p=0.05');
        yline(0.01, 'r:', 'LineWidth', 2, 'DisplayName', 'p=0.01');
        grid on;
        legend;
        
        % 상관계수 절댓값으로 정렬
        subplot(2, 2, 3);
        [sortedCorrs, sortIdx] = sort(abs(correlations), 'descend');
        sortedNames = categoryNames(sortIdx);
        sortedPvals = pValues(sortIdx);
        
        bar(sortedCorrs);
        title('상위항목-성과점수 상관계수 (절댓값 기준 정렬)', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('상위항목');
        ylabel('|상관계수|');
        xticklabels(sortedNames);
        xtickangle(45);
        grid on;
        
        % 유의성 표시
        for i = 1:length(sortedCorrs)
            if sortedPvals(i) < 0.001
                text(i, sortedCorrs(i) + 0.05, '***', 'HorizontalAlignment', 'center', 'FontSize', 12, 'Color', 'red');
            elseif sortedPvals(i) < 0.01
                text(i, sortedCorrs(i) + 0.05, '**', 'HorizontalAlignment', 'center', 'FontSize', 12, 'Color', 'red');
            elseif sortedPvals(i) < 0.05
                text(i, sortedCorrs(i) + 0.05, '*', 'HorizontalAlignment', 'center', 'FontSize', 12, 'Color', 'red');
            end
        end
        
        % 상관계수 vs p값 산점도
        subplot(2, 2, 4);
        scatter(correlations, pValues, 100, 'filled');
        title('상관계수 vs p값', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('상관계수');
        ylabel('p값');
        xline(0, 'k--', 'LineWidth', 1);
        yline(0.05, 'r--', 'LineWidth', 2, 'DisplayName', 'p=0.05');
        yline(0.01, 'r:', 'LineWidth', 2, 'DisplayName', 'p=0.01');
        grid on;
        legend;
        
        % 상위항목명 표시
        for i = 1:length(correlations)
            text(correlations(i), pValues(i), categoryNames{i}, 'FontSize', 8, 'HorizontalAlignment', 'center');
        end
        
        fprintf('  ✓ 상위항목-성과점수 상관관계 시각화 완료\n');
    end
    
    % 2. 중다회귀분석 결과 시각화
    if isfield(upperCategoryResults, 'regressionTable') && ~isempty(upperCategoryResults.regressionTable)
        figure('Name', '상위항목 중다회귀분석 결과', 'Position', [200, 200, 1200, 600]);
        
        regTable = upperCategoryResults.regressionTable;
        coefficients = regTable.Coefficient;
        ciLower = regTable.CI_Lower;
        ciUpper = regTable.CI_Upper;
        categoryNames = regTable.UpperCategory;
        
        % 회귀계수 막대그래프
        subplot(1, 2, 1);
        bar(coefficients);
        title('상위항목 회귀계수', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('상위항목');
        ylabel('회귀계수');
        xticklabels(categoryNames);
        xtickangle(45);
        grid on;
        
        % 신뢰구간 표시
        hold on;
        errorbar(1:length(coefficients), coefficients, coefficients - ciLower, ciUpper - coefficients, 'k.', 'LineWidth', 2);
        hold off;
        
        % 회귀계수 절댓값으로 정렬
        subplot(1, 2, 2);
        [sortedCoeffs, sortIdx] = sort(abs(coefficients), 'descend');
        sortedNames = categoryNames(sortIdx);
        sortedCILower = ciLower(sortIdx);
        sortedCIUpper = ciUpper(sortIdx);
        
        bar(sortedCoeffs);
        title('상위항목 회귀계수 (절댓값 기준 정렬)', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('상위항목');
        ylabel('|회귀계수|');
        xticklabels(sortedNames);
        xtickangle(45);
        grid on;
        
        % 신뢰구간 표시
        hold on;
        errorbar(1:length(sortedCoeffs), sortedCoeffs, sortedCoeffs - sortedCILower, sortedCIUpper - sortedCoeffs, 'k.', 'LineWidth', 2);
        hold off;
        
        fprintf('  ✓ 상위항목 중다회귀분석 결과 시각화 완료\n');
    end
    
    % 3. 예측 정확도 시각화
    if isfield(upperCategoryResults, 'predictionTable') && ~isempty(upperCategoryResults.predictionTable)
        figure('Name', '상위항목 예측 정확도', 'Position', [300, 300, 800, 600]);
        
        predTable = upperCategoryResults.predictionTable;
        
        % 예측 정확도 지표들
        metrics = {'RMSE', 'MAE', 'Correlation'};
        values = [predTable.RMSE, predTable.MAE, predTable.Correlation];
        
        bar(values);
        title('상위항목 예측 정확도', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('예측 정확도 지표');
        ylabel('값');
        xticklabels(metrics);
        grid on;
        
        % 값 표시
        for i = 1:length(values)
            text(i, values(i) + 0.01, sprintf('%.3f', values(i)), 'HorizontalAlignment', 'center', 'FontSize', 12);
        end
        
        fprintf('  ✓ 상위항목 예측 정확도 시각화 완료\n');
    end
    
    fprintf('  ✓ 모든 상위항목 시각화 생성 완료\n');
end

function summaryTable = createSummaryTable(correlationMatrices, periods)
    % 상관분석 결과 요약 테이블 생성
    if isempty(correlationMatrices)
        summaryTable = table();
        return;
    end
    
    % 요약 데이터 수집
    summaryData = {};
    resultFields = fieldnames(correlationMatrices);
    
    for i = 1:length(resultFields)
        fieldName = resultFields{i};
        result = correlationMatrices.(fieldName);
        
        if ~isempty(result) && isfield(result, 'correlationMatrix') && isfield(result, 'sampleSize')
            % 상관분석 결과 요약
            corrMatrix = result.correlationMatrix;
            pValueMatrix = result.pValueMatrix;
            
            % 유의한 상관관계 개수 (p < 0.05)
            significantCorrs = sum(pValueMatrix < 0.05, 'all');
            totalCorrs = numel(corrMatrix);
            
            % 최대 상관계수
            maxCorr = max(abs(corrMatrix(:)));
            
            % 평균 상관계수
            meanCorr = mean(abs(corrMatrix(:)), 'omitnan');
            
            summaryRow = {fieldName, result.sampleSize, totalCorrs, significantCorrs, maxCorr, meanCorr};
            summaryData = [summaryData; summaryRow];
        end
    end
    
    % 테이블 생성
    if ~isempty(summaryData)
        summaryTable = cell2table(summaryData, 'VariableNames', ...
            {'Period', 'SampleSize', 'TotalCorrelations', 'SignificantCorrelations', 'MaxCorrelation', 'MeanCorrelation'});
    else
        summaryTable = table();
    end
end

function [integratedData, overallCorrelation] = integratePerformanceScores(performanceResults, competencyTestData, periods, correlationMatrices)
    % 5개 시점의 성과점수를 개인별로 통합하고 역량검사점수와 상관분석하는 함수
    integratedData = [];
    overallCorrelation = [];
    
    if isempty(performanceResults) || isempty(competencyTestData)
        fprintf('  [경고] 통합 분석을 위한 데이터가 부족합니다\n');
        return;
    end
    
    fprintf('  ▶ 5개 시점 성과점수 통합 분석 중...\n');
    
    % 역량검사 데이터에서 ID 추출
    competencyIDCol = findIDColumn(competencyTestData);
    if isempty(competencyIDCol)
        fprintf('  [경고] 역량검사 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        return;
    end
    
    competencyIDs = extractAndStandardizeIDs(competencyTestData{:, competencyIDCol});
    
    % 역량검사 점수 추출
    [competencyScores, usedColumnName] = findCompetencyScoreColumn(competencyTestData, 1:height(competencyTestData));
    if isempty(competencyScores)
        fprintf('  [경고] 역량검사 점수를 찾을 수 없습니다\n');
        return;
    end
    
    fprintf('  - 사용된 역량점수 컬럼: "%s"\n', usedColumnName);
    
    % 각 시점의 성과점수 수집
    allPerformanceData = [];
    resultFields = fieldnames(performanceResults);
    
    for i = 1:length(resultFields)
        fieldName = resultFields{i};
        result = performanceResults.(fieldName);
        
        if ~isempty(result) && isfield(result, 'performanceScores') && isfield(result, 'sampleIDs')
            % 성과점수와 ID 매칭
            perfIDs = result.sampleIDs;
            perfScores = result.performanceScores;
            
            % ID 표준화
            perfIDs = extractAndStandardizeIDs(perfIDs);
            
            % 역량검사 ID와 매칭
            [commonIDs, perfIdx, compIdx] = intersect(perfIDs, competencyIDs);
            
            if ~isempty(commonIDs)
                matchedPerfScores = perfScores(perfIdx);
                matchedCompScores = competencyScores(compIdx);
                
                % 데이터 추가
                for j = 1:length(commonIDs)
                    allPerformanceData = [allPerformanceData; {commonIDs{j}, matchedPerfScores(j), matchedCompScores(j)}];
                end
            end
        end
    end
    
    if isempty(allPerformanceData)
        fprintf('  [경고] 매칭되는 성과점수 데이터가 없습니다\n');
        return;
    end
    
    % 개인별로 성과점수 평균 계산
    uniqueIDs = unique(allPerformanceData(:, 1));
    integratedScores = [];
    
    for i = 1:length(uniqueIDs)
        id = uniqueIDs{i};
        idRows = strcmp(allPerformanceData(:, 1), id);
        
        if sum(idRows) > 0
            perfScores = cell2mat(allPerformanceData(idRows, 2));
            compScores = cell2mat(allPerformanceData(idRows, 3));
            
            % 성과점수는 평균, 역량검사점수는 첫 번째 값 사용
            avgPerfScore = mean(perfScores);
            compScore = compScores(1);
            
            integratedScores = [integratedScores; {id, avgPerfScore, compScore}];
        end
    end
    
    if isempty(integratedScores)
        fprintf('  [경고] 통합된 성과점수 데이터가 없습니다\n');
        return;
    end
    
    % 테이블로 변환
    integratedData = cell2table(integratedScores, 'VariableNames', {'ID', 'PerformanceScore', 'CompetencyScore'});
    
    % 상관분석 수행
    perfScores = integratedData.PerformanceScore;
    compScores = integratedData.CompetencyScore;
    
    [r, p] = corr(perfScores, compScores, 'rows', 'pairwise');
    
    overallCorrelation = struct();
    overallCorrelation.correlation = r;
    overallCorrelation.pValue = p;
    overallCorrelation.sampleSize = height(integratedData);
    
    fprintf('  ✓ 통합 분석 완료: %d명, r=%.3f (p=%.3f)\n', height(integratedData), r, p);
end

% ID를 기반으로 Period 데이터와 상위항목 데이터를 매칭하는 함수
%
% 입력:
%   periodData - Period 데이터 (테이블)
%   upperCategoryResults - 상위항목 분석 결과 구조체
%
% 출력:
%   matchedData - 매칭된 Period 데이터
%   matchedUpperScores - 매칭된 상위항목 점수
%   matchedIdx - 매칭된 인덱스

    matchedData = [];
    matchedUpperScores = [];
    matchedIdx = [];
    
    if isempty(periodData) || isempty(upperCategoryResults)
        return;
    end
    
    % Period 데이터에서 ID 추출
    periodIDCol = findIDColumn(periodData);
    if isempty(periodIDCol)
        fprintf('  [경고] Period 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        return;
    end
    
    periodIDs = extractAndStandardizeIDs(periodData{:, periodIDCol});
    
    % 상위항목 데이터에서 ID 추출
    if isfield(upperCategoryResults, 'commonIDs')
        upperIDs = upperCategoryResults.commonIDs;
    else
        fprintf('  [경고] 상위항목 데이터에서 ID를 찾을 수 없습니다\n');
        return;
    end
    
    % ID 매칭
    [commonIDs, periodIdx, upperIdx] = intersect(periodIDs, upperIDs);
    
    if isempty(commonIDs)
        fprintf('  [경고] 매칭되는 ID가 없습니다\n');
        return;
    end
    
    % 매칭된 데이터 구성
    matchedData = periodData(periodIdx, :);
    matchedIdx = periodIdx;
    
    % 상위항목 점수 추출
    if isfield(upperCategoryResults, 'scoreMatrix')
        matchedUpperScores = upperCategoryResults.scoreMatrix(upperIdx, :);
    else
        fprintf('  [경고] 상위항목 점수 데이터를 찾을 수 없습니다\n');
        matchedUpperScores = [];
    end
    
    fprintf('  ✓ ID 매칭 완료: %d명\n', length(commonIDs));
end

    % 상위항목 점수와 성과점수 간 상관분석 및 중다회귀분석
    upperCategoryResults = [];
    
    if isempty(upperCategoryData) || isempty(performanceResults)
        fprintf('  [경고] 상위항목 분석을 위한 데이터가 부족합니다\n');
        return;
    end
    
    fprintf('  ▶ 상위항목 성과분석 시작\n');
    
    % ID 컬럼 찾기
    upperIDCol = findIDColumn(upperCategoryData);
    if isempty(upperIDCol)
        fprintf('  [경고] 상위항목 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        return;
    end
    
    % 상위항목 점수 컬럼 식별
    colNames = upperCategoryData.Properties.VariableNames;
    scoreColumns = {};
    scoreColumnNames = {};
    
    for i = 1:width(upperCategoryData)
        if i == upperIDCol
            continue;
        end
        
        colName = colNames{i};
        colData = upperCategoryData{:, i};
        
        if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
            scoreColumns{end+1} = i;
            scoreColumnNames{end+1} = colName;
        end
    end
    
    fprintf('  - 발견된 상위항목: %d개\n', length(scoreColumnNames));
    
    if length(scoreColumnNames) < 2
        fprintf('  [경고] 분석에 필요한 상위항목이 부족합니다\n');
        return;
    end
    
    % ID 표준화
    upperIDs = extractAndStandardizeIDs(upperCategoryData{:, upperIDCol});
    
    % 통합 성과점수가 있는 경우 사용, 없으면 개별 시점 성과점수 사용
    if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData) && ...
       ismember('PerformanceScore', integratedPerformanceData.Properties.VariableNames)
        
        % 통합 성과점수 사용
        performanceIDs = integratedPerformanceData.ID;
        performanceScores = integratedPerformanceData.PerformanceScore;
        
        fprintf('  - 통합 성과점수 사용: %d명\n', height(integratedPerformanceData));
        
    else
        % 개별 시점 성과점수 사용
        performanceIDs = {};
        performanceScores = [];
        
        resultFields = fieldnames(performanceResults);
        for i = 1:length(resultFields)
            result = performanceResults.(resultFields{i});
            if ~isempty(result) && isfield(result, 'sampleIDs') && isfield(result, 'performanceScores')
                performanceIDs = [performanceIDs; result.sampleIDs];
                performanceScores = [performanceScores; result.performanceScores];
            end
        end
        
        if isempty(performanceIDs)
            fprintf('  [경고] 성과점수 데이터를 찾을 수 없습니다\n');
            return;
        end
        
        fprintf('  - 개별 시점 성과점수 사용: %d명\n', length(performanceIDs));
    end
    
    % ID 매칭
    [commonIDs, upperIdx, perfIdx] = intersect(upperIDs, performanceIDs);
    
    if length(commonIDs) < 10
        fprintf('  [경고] 매칭되는 데이터가 부족합니다 (%d명)\n', length(commonIDs));
        return;
    end
    
    fprintf('  - 매칭된 데이터: %d명\n', length(commonIDs));
    
    % 매칭된 상위항목 데이터 구성
    upperScoreMatrix = [];
    for i = 1:length(scoreColumns)
        colIdx = scoreColumns{i};
        colData = upperCategoryData{upperIdx, colIdx};
        upperScoreMatrix = [upperScoreMatrix, colData];
    end
    
    matchedPerformanceScores = performanceScores(perfIdx);
    
    % 상위항목은 pairwise correlation으로 처리
    validPerformanceRows = ~isnan(matchedPerformanceScores);
    cleanPerformanceScores = matchedPerformanceScores(validPerformanceRows);
    cleanUpperMatrix = upperScoreMatrix(validPerformanceRows, :);
    cleanCommonIDs = commonIDs(validPerformanceRows);
    
    fprintf('  - 유효한 데이터: %d명\n', length(cleanCommonIDs));
    
    % 각 상위항목별 유효 데이터 수 확인
    for i = 1:size(cleanUpperMatrix, 2)
        validCount = sum(~isnan(cleanUpperMatrix(:, i)));
        fprintf('    %s: %d명 유효\n', scoreColumnNames{i}, validCount);
    end
    
    if length(cleanCommonIDs) < 10
        fprintf('  [경고] 유효한 데이터가 부족합니다\n');
        return;
    end
    
    %% 1. 상관분석
    fprintf('\n  ▶ 상위항목-성과점수 상관분석\n');
    
    correlationResults = [];
    pValues = [];
    
    for i = 1:length(scoreColumnNames)
        [r, p] = corr(cleanUpperMatrix(:, i), cleanPerformanceScores, 'rows', 'pairwise');
        correlationResults(end+1) = r;
        pValues(end+1) = p;
        
        fprintf('    %s: r=%.3f, p=%.3f', scoreColumnNames{i}, r, p);
        if p < 0.001
            fprintf(' ***');
        elseif p < 0.01
            fprintf(' **');
        elseif p < 0.05
            fprintf(' *');
        end
        fprintf('\n');
    end
    
    % 상관분석 결과 테이블 생성
    correlationTable = table(scoreColumnNames', correlationResults', pValues', ...
        'VariableNames', {'UpperCategory', 'Correlation', 'PValue'});
    
    %% 2. 중다회귀분석
    fprintf('\n  ▶ 중다회귀분석\n');
    
    try
        % 최소 70% 이상의 상위항목 데이터가 있는 행만 사용
        missingCount = sum(isnan(cleanUpperMatrix), 2);
        maxMissing = floor(size(cleanUpperMatrix, 2) * 0.3); % 30%까지 결측 허용
        regressionValidRows = missingCount <= maxMissing;
        
        regressionMatrix = cleanUpperMatrix(regressionValidRows, :);
        regressionScores = cleanPerformanceScores(regressionValidRows);
        
        if size(regressionMatrix, 1) < 10
            fprintf('    [경고] 회귀분석을 위한 데이터가 부족합니다\n');
            regressionTable = table();
        else
            % 중다회귀분석 수행
            [b, bint, r, rint, stats] = regress(regressionScores, [ones(size(regressionMatrix, 1), 1), regressionMatrix]);
            
            % 결과 테이블 생성
            regressionTable = table(scoreColumnNames', b(2:end), bint(2:end, 1), bint(2:end, 2), ...
                'VariableNames', {'UpperCategory', 'Coefficient', 'CI_Lower', 'CI_Upper'});
            
            fprintf('    R² = %.3f, F = %.3f, p = %.3f\n', stats(1), stats(2), stats(3));
        end
        
        % 다중공선성 확인
        corrMatrix = corrcoef(regressionMatrix, 'rows', 'pairwise');
        maxCorr = max(corrMatrix(:));
        fprintf('    - 상위항목 간 최대 상관: %.3f\n', maxCorr);
        
        if maxCorr > 0.9
            fprintf('    ⚠️ 높은 다중공선성 감지 (r > 0.9)\n');
        end
        
    catch ME
        fprintf('    [경고] 중다회귀분석 실패: %s\n', ME.message);
        regressionTable = table();
    end
    
    %% 3. 예측 정확도 분석
    fprintf('\n  ▶ 예측 정확도 분석\n');
    
    try
        if size(regressionMatrix, 1) >= 10
            % 교차검증을 위한 예측 정확도 계산
            cv = cvpartition(size(regressionMatrix, 1), 'HoldOut', 0.3);
            trainIdx = cv.training;
            testIdx = cv.test;
            
            trainX = regressionMatrix(trainIdx, :);
            trainY = regressionScores(trainIdx);
            testX = regressionMatrix(testIdx, :);
            testY = regressionScores(testIdx);
            
            % 훈련 데이터로 모델 학습
            [b, ~, ~, ~, ~] = regress(trainY, [ones(size(trainX, 1), 1), trainX]);
            
            % 테스트 데이터로 예측
            predictedY = [ones(size(testX, 1), 1), testX] * b;
            
            % 예측 정확도 계산
            mse = mean((testY - predictedY).^2);
            rmse = sqrt(mse);
            mae = mean(abs(testY - predictedY));
            
            % 상관계수
            [r, p] = corr(testY, predictedY, 'rows', 'pairwise');
            
            fprintf('    RMSE: %.3f, MAE: %.3f, r=%.3f (p=%.3f)\n', rmse, mae, r, p);
            
            % 예측 정확도 테이블 생성
            predictionTable = table({rmse}, {mae}, {r}, {p}, ...
                'VariableNames', {'RMSE', 'MAE', 'Correlation', 'PValue'});
        else
            fprintf('    [경고] 예측 정확도 분석을 위한 데이터가 부족합니다\n');
            predictionTable = table();
        end
        
    catch ME
        fprintf('    [경고] 예측 정확도 분석 실패: %s\n', ME.message);
        predictionTable = table();
    end
    
    %% 결과 저장
    upperCategoryResults = struct();
    upperCategoryResults.correlationTable = correlationTable;
    upperCategoryResults.regressionTable = regressionTable;
    upperCategoryResults.predictionTable = predictionTable;
    upperCategoryResults.scoreColumnNames = {scoreColumnNames};
    upperCategoryResults.commonIDs = cleanCommonIDs;
    upperCategoryResults.scoreMatrix = cleanUpperMatrix;
    
    fprintf('  ✓ 상위항목 분석 완료\n');
end

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
    
    % 상관분석 수행
    validRows = ~isnan(finalTable.CompetencyScore) & ~isnan(finalTable.IntegratedPerformanceScore);
    cleanData = finalTable(validRows, :);
    
    if height(cleanData) < 3
        fprintf('  [경고] 상관분석을 위한 유효 데이터가 부족합니다 (%d명)\n', height(cleanData));
        integratedData = finalTable;
        overallCorrelation = struct();
        return;
    end
    
    [corrCoeff, pValue] = corr(cleanData.CompetencyScore, cleanData.IntegratedPerformanceScore);
    
    % 유의성 판정
    if pValue < 0.001
        significance = '***';
    elseif pValue < 0.01
        significance = '**';
    elseif pValue < 0.05
        significance = '*';
    else
        significance = 'ns';
    end
    
    % 결과 구조체 생성
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
    
    % 추가 통계 정보를 테이블에 추가
    finalTable.CompetencyMean = repmat(overallCorrelation.competencyMean, height(finalTable), 1);
    finalTable.PerformanceMean = repmat(overallCorrelation.performanceMean, height(finalTable), 1);
    finalTable.OverallCorrelation = repmat(corrCoeff, height(finalTable), 1);
    finalTable.PValue = repmat(pValue, height(finalTable), 1);
    finalTable.Significance = repmat({significance}, height(finalTable), 1);
    
    integratedData = finalTable;
    
    fprintf('  - 종합 성과점수: 평균 %.2f (SD %.2f)\n', ...
        overallCorrelation.performanceMean, overallCorrelation.performanceStd);
    fprintf('  - 역량검사점수: 평균 %.2f (SD %.2f)\n', ...
        overallCorrelation.competencyMean, overallCorrelation.competencyStd);
end

    % 상위항목 점수와 성과점수 간 상관분석 및 중다회귀분석

    fprintf('▶ 상위항목 성과분석 시작\n');

    % ID 컬럼 찾기
    upperIDCol = findIDColumn(upperCategoryData);
    if isempty(upperIDCol)
        fprintf('✗ 상위항목 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
        upperCategoryResults = [];
        return;
    end

    % 상위항목 점수 컬럼 식별
    colNames = upperCategoryData.Properties.VariableNames;
    scoreColumns = {};
    scoreColumnNames = {};

    for i = 1:width(upperCategoryData)
        if i == upperIDCol
            continue;
        end

        colName = colNames{i};
        colData = upperCategoryData{:, i};

        if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
            scoreColumns{end+1} = i;
            scoreColumnNames{end+1} = colName;
        end
    end

    fprintf('  - 발견된 상위항목: %d개\n', length(scoreColumnNames));

    if length(scoreColumnNames) < 2
        fprintf('✗ 분석에 필요한 상위항목이 부족합니다\n');
        upperCategoryResults = [];
        return;
    end

    % ID 표준화
    upperIDs = extractAndStandardizeIDs(upperCategoryData{:, upperIDCol});

    % 통합 성과점수가 있는 경우 사용, 없으면 개별 시점 성과점수 사용
    if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData) && ...
       ismember('PerformanceScore', integratedPerformanceData.Properties.VariableNames)

        fprintf('  - 통합 성과점수 사용\n');
        performanceIDs = integratedPerformanceData.ID;
        performanceScores = integratedPerformanceData.PerformanceScore;

    elseif ~isempty(fieldnames(performanceResults))

        fprintf('  - 개별 시점 성과점수 통합 중\n');
        % 개별 시점 성과점수들을 통합
        allPerformanceData = table();
        perfFields = fieldnames(performanceResults);

        for i = 1:length(perfFields)
            result = performanceResults.(perfFields{i});
            if isfield(result, 'cleanIDs') && isfield(result, 'performanceScores')
                tempTable = table();
                tempTable.ID = result.cleanIDs;
                tempTable.PerformanceScore = result.performanceScores;
                allPerformanceData = [allPerformanceData; tempTable];
            end
        end

        % 개인별 평균 계산
        uniqueIDs = unique(allPerformanceData.ID);
        performanceIDs = {};
        performanceScores = [];

        for i = 1:length(uniqueIDs)
            personID = uniqueIDs{i};
            personScores = allPerformanceData.PerformanceScore(strcmp(allPerformanceData.ID, personID));
            avgScore = nanmean(personScores);

            performanceIDs{end+1} = personID;
            performanceScores(end+1) = avgScore;
        end

    else
        fprintf('✗ 성과점수 데이터를 찾을 수 없습니다\n');
        upperCategoryResults = [];
        return;
    end

    % ID 매칭
    [commonIDs, upperIdx, perfIdx] = intersect(upperIDs, performanceIDs);

    if length(commonIDs) < 10
        fprintf('✗ 매칭된 데이터가 부족합니다 (%d명)\n', length(commonIDs));
        upperCategoryResults = [];
        return;
    end

    fprintf('  - 매칭된 데이터: %d명\n', length(commonIDs));

    % 매칭된 상위항목 데이터 구성
    upperScoreMatrix = [];
    for i = 1:length(scoreColumns)
        colIdx = scoreColumns{i};
        colData = upperCategoryData{upperIdx, colIdx};
        upperScoreMatrix = [upperScoreMatrix, colData];
    end

    matchedPerformanceScores = performanceScores(perfIdx);

    % 개선된 결측치 처리: 성과점수가 있는 사람만 필터링
    % 상위항목은 pairwise correlation으로 처리
    validPerformanceRows = ~isnan(matchedPerformanceScores);
    cleanPerformanceScores = matchedPerformanceScores(validPerformanceRows);
    cleanUpperMatrix = upperScoreMatrix(validPerformanceRows, :);
    cleanCommonIDs = commonIDs(validPerformanceRows);

    fprintf('  - 성과점수 기준 필터링 후: %d명\n', length(cleanPerformanceScores));

    % 각 상위항목별 유효 데이터 수 확인
    for i = 1:size(cleanUpperMatrix, 2)
        validCount = sum(~isnan(cleanUpperMatrix(:, i)));
        fprintf('    %s: %d명 유효\n', scoreColumnNames{i}, validCount);
    end

    if length(cleanPerformanceScores) < 10
        fprintf('✗ 분석에 충분한 데이터가 없습니다\n');
        upperCategoryResults = [];
        return;
    end

    %% 1. 상관분석
    fprintf('\n▶ 상위항목-성과점수 상관분석\n');

    correlationResults = [];
    pValues = [];

    for i = 1:length(scoreColumnNames)
        % Pairwise correlation (결측치가 있는 쌍만 제외하고 계산)
        [r, p] = corr(cleanUpperMatrix(:, i), cleanPerformanceScores, 'rows', 'pairwise');

        % 실제 사용된 데이터 개수 확인
        validPairs = ~isnan(cleanUpperMatrix(:, i)) & ~isnan(cleanPerformanceScores);
        actualN = sum(validPairs);

        correlationResults(i) = r;
        pValues(i) = p;

        sig_str = '';
        if p < 0.001, sig_str = '***';
        elseif p < 0.01, sig_str = '**';
        elseif p < 0.05, sig_str = '*';
        end

        fprintf('  %s: r = %.3f (p = %.3f) %s (N=%d)\n', ...
            scoreColumnNames{i}, r, p, sig_str, actualN);
    end

    % 상관분석 결과 테이블 생성
    correlationTable = table();
    correlationTable.UpperCategory = scoreColumnNames';
    correlationTable.Correlation = correlationResults';
    correlationTable.PValue = pValues';

    % 유의성 표시
    significance = cell(length(pValues), 1);
    for i = 1:length(pValues)
        if pValues(i) < 0.001
            significance{i} = '***';
        elseif pValues(i) < 0.01
            significance{i} = '**';
        elseif pValues(i) < 0.05
            significance{i} = '*';
        else
            significance{i} = 'ns';
        end
    end
    correlationTable.Significance = significance;

    %% 2. 중다회귀분석
    fprintf('\n▶ 중다회귀분석 (성과점수 예측)\n');

    try
        % 중다회귀분석을 위해 완전한 데이터만 사용 (더 관대한 기준)
        % 최소 70% 이상의 상위항목 데이터가 있는 행만 사용
        missingCount = sum(isnan(cleanUpperMatrix), 2);
        maxMissing = floor(size(cleanUpperMatrix, 2) * 0.3); % 30%까지 결측 허용
        regressionValidRows = missingCount <= maxMissing;

        regressionMatrix = cleanUpperMatrix(regressionValidRows, :);
        regressionPerformanceScores = cleanPerformanceScores(regressionValidRows);
        regressionIDs = cleanCommonIDs(regressionValidRows);

        fprintf('  - 중다회귀분석 데이터: %d명 (70%% 이상 완전한 데이터)\n', sum(regressionValidRows));

        if sum(regressionValidRows) < 15
            fprintf('  ⚠️ 중다회귀분석을 위한 데이터가 부족합니다 (N=%d)\n', sum(regressionValidRows));
            error('insufficient_data');
        end

        % 다중공선성 확인 (pairwise correlation)
        upperCorr = corrcoef(regressionMatrix, 'rows', 'pairwise');
        maxCorr = max(abs(upperCorr - eye(size(upperCorr))), [], 'all');
        fprintf('  - 상위항목 간 최대 상관: %.3f\n', maxCorr);

        if maxCorr > 0.9
            fprintf('  ⚠️ 높은 다중공선성 감지 (r > 0.9)\n');
        end

        % 중다회귀분석 수행 (완전한 사례만 사용)
        % 결측치가 있는 행 제거
        completeRows = ~any(isnan(regressionMatrix), 2);
        finalMatrix = regressionMatrix(completeRows, :);
        finalScores = regressionPerformanceScores(completeRows);
        finalIDs = regressionIDs(completeRows);

        fprintf('  - 실제 회귀분석 데이터: %d명 (완전한 사례)\n', length(finalScores));

        if size(finalScores, 1) == 1
            finalScores = finalScores';
        end
        [b, bint, r, rint, stats] = regress(finalScores, [ones(size(finalMatrix, 1), 1), finalMatrix]);

        rSquared = stats(1);
        fStat = stats(2);
        pValue = stats(3);

        fprintf('  - R² = %.3f (설명변량: %.1f%%)\n', rSquared, rSquared * 100);
        fprintf('  - F(%d,%d) = %.2f, p = %.3f\n', ...
            length(scoreColumnNames), length(cleanPerformanceScores) - length(scoreColumnNames) - 1, ...
            fStat, pValue);

        % 회귀계수 결과 테이블 생성
        regressionTable = table();
        predictorNames = [{'절편'}; scoreColumnNames'];
        regressionTable.Predictor = predictorNames;
        regressionTable.Coefficient = b;
        regressionTable.CI_Lower = bint(:, 1);
        regressionTable.CI_Upper = bint(:, 2);

        % t-검정
        se = (bint(:, 2) - bint(:, 1)) / (2 * 1.96); % 표준오차 추정
        tStats = b ./ se;
        df = length(cleanPerformanceScores) - length(b);
        pValuesReg = 2 * (1 - tcdf(abs(tStats), df));

        regressionTable.SE = se;
        regressionTable.tStat = tStats;
        regressionTable.PValue = pValuesReg;

        % 예측값 계산 (완전한 사례에 대해서만)
        predictedScores = [ones(size(finalMatrix, 1), 1), finalMatrix] * b;

        % 예측 정확도 평가
        mae = mean(abs(finalScores - predictedScores));
        rmse = sqrt(mean((finalScores - predictedScores).^2));

        fprintf('  - MAE (평균절대오차): %.3f\n', mae);
        fprintf('  - RMSE (평균제곱근오차): %.3f\n', rmse);

        % 예측 결과 테이블 생성 (완전한 사례만)
        predictionTable = table();
        predictionTable.ID = finalIDs;
        predictionTable.ActualPerformance = finalScores;
        predictionTable.PredictedPerformance = predictedScores;
        predictionTable.Residual = finalScores - predictedScores;
        predictionTable.AbsoluteError = abs(predictionTable.Residual);

        % 상위항목 점수도 포함 (완전한 사례에 대해서만)
        for i = 1:length(scoreColumnNames)
            predictionTable.(scoreColumnNames{i}) = finalMatrix(:, i);
        end

    catch ME
        fprintf('✗ 중다회귀분석 실패: %s\n', ME.message);
        regressionTable = table();
        predictionTable = table();
        rSquared = NaN;
        mae = NaN;
        rmse = NaN;
    end

    %% 결과 구조체 생성
    upperCategoryResults = struct();
    upperCategoryResults.correlationTable = correlationTable;
    upperCategoryResults.regressionTable = regressionTable;
    upperCategoryResults.predictionTable = predictionTable;
    upperCategoryResults.cleanData = cleanUpperMatrix;
    upperCategoryResults.upperScoreMatrix = cleanUpperMatrix;  % 13단계를 위한 필드 추가
    upperCategoryResults.cleanPerformanceScores = cleanPerformanceScores;
    upperCategoryResults.scoreColumnNames = {scoreColumnNames};
    upperCategoryResults.cleanIDs = {cleanCommonIDs};
    upperCategoryResults.matchedIDs = cleanCommonIDs;  % 13단계를 위한 필드 추가
    upperCategoryResults.rSquared = rSquared;
    upperCategoryResults.mae = mae;
    upperCategoryResults.rmse = rmse;

    fprintf('✓ 상위항목 성과분석 완료\n');
end

    % 상위항목 분석 결과 시각화

    fprintf('▶ 상위항목 분석 시각화 생성 중...\n');

    if isempty(upperCategoryResults)
        return;
    end

    %% 1. 상관계수 막대그래프
    figure('Name', '상위항목-성과점수 상관계수', 'Position', [100, 100, 1000, 600]);

    corrData = upperCategoryResults.correlationTable;
    categories = corrData.UpperCategory;
    correlations = corrData.Correlation;
    pValues = corrData.PValue;

    % 막대그래프 생성
    bars = bar(correlations);
    set(bars, 'FaceColor', [0.3 0.6 0.8]);

    % 유의한 상관계수 강조
    hold on;
    for i = 1:length(pValues)
        if pValues(i) < 0.05
            bar(i, correlations(i), 'FaceColor', [0.8 0.3 0.3]);
        end
    end

    set(gca, 'XTickLabel', categories, 'XTickLabelRotation', 45);
    xlabel('상위항목');
    ylabel('성과점수와의 상관계수');
    title('상위항목별 성과점수 상관분석', 'FontSize', 14, 'FontWeight', 'bold');
    grid on;

    % 유의성 표시
    for i = 1:length(correlations)
        y_pos = correlations(i) + sign(correlations(i)) * 0.02;
        if pValues(i) < 0.001
            text(i, y_pos, '***', 'HorizontalAlignment', 'center', 'FontWeight', 'bold');
        elseif pValues(i) < 0.01
            text(i, y_pos, '**', 'HorizontalAlignment', 'center', 'FontWeight', 'bold');
        elseif pValues(i) < 0.05
            text(i, y_pos, '*', 'HorizontalAlignment', 'center', 'FontWeight', 'bold');
        end
    end

    hold off;

    %% 2. 실제값 vs 예측값 산점도
    if isfield(upperCategoryResults, 'predictionTable') && height(upperCategoryResults.predictionTable) > 0
        figure('Name', '성과점수 예측 정확도', 'Position', [200, 200, 800, 600]);

        predData = upperCategoryResults.predictionTable;
        actual = predData.ActualPerformance;
        predicted = predData.PredictedPerformance;

        scatter(actual, predicted, 50, 'filled', 'MarkerFaceColor', [0.3 0.6 0.8]);
        hold on;

        % 완벽한 예측선 (y=x)
        minVal = min([actual; predicted]);
        maxVal = max([actual; predicted]);
        plot([minVal, maxVal], [minVal, maxVal], 'r--', 'LineWidth', 2);

        % 회귀선
        p = polyfit(actual, predicted, 1);
        x_line = linspace(minVal, maxVal, 100);
        y_line = polyval(p, x_line);
        plot(x_line, y_line, 'g-', 'LineWidth', 1.5);

        xlabel('실제 성과점수');
        ylabel('예측 성과점수');
        title('성과점수 예측 정확도', 'FontSize', 14, 'FontWeight', 'bold');

        % 통계 정보 표시
        r2_text = sprintf('R² = %.3f', upperCategoryResults.rSquared);
        mae_text = sprintf('MAE = %.3f', upperCategoryResults.mae);
        rmse_text = sprintf('RMSE = %.3f', upperCategoryResults.rmse);

        text(0.05, 0.95, {r2_text, mae_text, rmse_text}, ...
            'Units', 'normalized', 'FontSize', 11, ...
            'BackgroundColor', 'white', 'EdgeColor', 'black', ...
            'VerticalAlignment', 'top');

        legend('데이터 포인트', '완벽한 예측 (y=x)', '회귀선', 'Location', 'southeast');
        grid on;
        hold off;
    end

    %% 3. 상위항목별 성과점수 박스플롯
    if isfield(upperCategoryResults, 'cleanData') && ~isempty(upperCategoryResults.cleanData)
        figure('Name', '상위항목별 점수 분포', 'Position', [300, 300, 1200, 700]);

        cleanData = upperCategoryResults.cleanData;
        scoreNames = upperCategoryResults.scoreColumnNames{1};

        % 데이터 준비 (상위항목별로)
        allScores = [];
        allCategories = {};

        for i = 1:length(scoreNames)
            scores = cleanData(:, i);
            allScores = [allScores; scores];
            categories = repmat({scoreNames{i}}, length(scores), 1);
            allCategories = [allCategories; categories];
        end

        % 박스플롯 생성
        boxplot(allScores, allCategories);
        xlabel('상위항목');
        ylabel('점수');
        title('상위항목별 점수 분포', 'FontSize', 14, 'FontWeight', 'bold');
        set(gca, 'XTickLabelRotation', 45);
        grid on;
    end

    fprintf('✓ 상위항목 분석 시각화 완료\n');
end


%% function13

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
        
        % 문항과 종합점수 간 상관의 통계
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
        
        summaryTable = [summaryTable; newRow];
    end
end

% ID를 기반으로 Period 데이터와 상위항목 데이터를 매칭하는 함수
%
% 입력:
%   periodData - Period별 역량진단 데이터 (테이블)
%   upperCategoryResults - 상위항목 분석 결과 구조체
%
% 출력:
%   matchedData - 매칭된 Period 데이터
%   matchedUpperScores - 매칭된 상위항목 점수 매트릭스
%   matchedIdx - 매칭된 인덱스

    matchedData = [];
    matchedUpperScores = [];
    matchedIdx = [];

    try
        % Period 데이터에서 ID 컬럼 찾기
        periodIDCol = findIDColumn(periodData);
        if isempty(periodIDCol)
            fprintf('  ✗ Period 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
            return;
        end

        % Period 데이터의 ID 추출 및 표준화
        periodIDs = extractAndStandardizeIDs(periodData{:, periodIDCol});

        % 상위항목 결과에서 ID 가져오기 (cleanIDs 또는 matchedIDs 사용)
        upperIDs = [];
        if isfield(upperCategoryResults, 'matchedIDs') && ~isempty(upperCategoryResults.matchedIDs)
            upperIDs = upperCategoryResults.matchedIDs;
        elseif isfield(upperCategoryResults, 'cleanIDs') && ~isempty(upperCategoryResults.cleanIDs)
            upperIDs = upperCategoryResults.cleanIDs{1};  % cell 배열에서 추출
        else
            fprintf('  ✗ 상위항목 결과에서 ID를 찾을 수 없습니다\n');
            return;
        end

        % 상위항목 점수 매트릭스 가져오기
        if isfield(upperCategoryResults, 'upperScoreMatrix') && ~isempty(upperCategoryResults.upperScoreMatrix)
            upperScoreMatrix = upperCategoryResults.upperScoreMatrix;
        elseif isfield(upperCategoryResults, 'cleanData') && ~isempty(upperCategoryResults.cleanData)
            upperScoreMatrix = upperCategoryResults.cleanData;
        else
            fprintf('  ✗ 상위항목 점수 매트릭스를 찾을 수 없습니다\n');
            return;
        end

        % ID 매칭
        [commonIDs, periodIdx, upperIdx] = intersect(periodIDs, upperIDs);

        if length(commonIDs) < 5
            fprintf('  ✗ 매칭된 ID가 부족합니다 (%d개)\n', length(commonIDs));
            return;
        end

        % 매칭된 데이터 추출
        matchedData = periodData(periodIdx, :);
        matchedUpperScores = upperScoreMatrix(upperIdx, :);
        matchedIdx = periodIdx;

        fprintf('  ✓ ID 매칭 성공: %d명 (전체 Period: %d명, 상위항목: %d명)\n', ...
                length(commonIDs), length(periodIDs), length(upperIDs));

    catch ME
        fprintf('  ✗ ID 매칭 중 오류 발생: %s\n', ME.message);
        matchedData = [];
        matchedUpperScores = [];
        matchedIdx = [];
    end
end