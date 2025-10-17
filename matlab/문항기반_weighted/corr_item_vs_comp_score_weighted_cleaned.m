%% 각 Period별 문항과 역량검사 종합점수 간 상관 매트릭스 생성 (가중치 적용 버전)
%
% 목적: 각 시점별로 수집된 문항들과 역량검사 종합점수 간의
%       전체 상관 매트릭스를 생성하고 분석 (가중치 적용)
%
% 작성일: 2025년

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('Period별 문항-종합점수 상관 매트릭스 생성 (가중치 적용)\n');
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
competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    competencyTestData = readtable( competencyTestPath, ...
        'Sheet', '역량검사_종합점수', ...
        'VariableNamingRule', 'preserve');
    fprintf('✓ 역량검사 데이터 로드 완료: %d명\n', height(competencyTestData));
catch ME
    fprintf('✗ 역량검사 데이터 로드 실패: %s\n', ME.message);
    return;
end

% 신뢰가능성 필터링
fprintf('\n【STEP 1-1】 신뢰가능성 필터링\n');
fprintf('────────────────────────────────────────────\n');

% 신뢰가능성 컬럼 찾기
reliability_col_idx = find(contains(competencyTestData.Properties.VariableNames, '신뢰가능성'), 1);

if ~isempty(reliability_col_idx)
    colName = competencyTestData.Properties.VariableNames{reliability_col_idx};
    fprintf('▶ 신뢰가능성 컬럼 발견: %s\n', colName);

    % 신뢰불가 데이터 찾기
    reliability_data = competencyTestData{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(competencyTestData), 1);
    end

    fprintf('  신뢰불가 데이터: %d명\n', sum(unreliable_idx));

    % 신뢰가능한 데이터만 유지
    competencyTestData = competencyTestData(~unreliable_idx, :);
    fprintf('  ✓ 신뢰가능한 데이터: %d명\n', height(competencyTestData));
else
    fprintf('  ⚠ 신뢰가능성 컬럼이 없습니다. 모든 데이터를 사용합니다.\n');
end

%% 2. 23년 상반기 데이터 추가 (선택적)
fprintf('\n[2단계] 23년 상반기 데이터 추가 확인\n');
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

%% 3. 종합성과 히스토그램 생성 (모든 기간 데이터 포함)
fprintf('\n[3단계] 종합성과 히스토그램 생성\n');
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

%% 4. 각 Period별 상관 분석
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

%% 5. 결과 저장
fprintf('\n[5단계] 결과 저장\n');
fprintf('----------------------------------------\n');

% 백업 폴더 확인 및 생성
backupDir = 'D:\project\HR데이터\결과\성과종합점수&역검_weighted\backup';
if ~exist(backupDir, 'dir')
    mkdir(backupDir);
    fprintf('✓ 백업 폴더 생성: %s\n', backupDir);
end

% 기존 파일들을 백업 폴더로 이동
existingFiles = dir('D:\project\HR데이터\결과\성과종합점수&역검_weighted\correlation_matrices_weighted_*.xlsx');
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
existingMatFiles = dir('D:\project\HR데이터\결과\성과종합점수&역검_weighted\correlation_workspace_weighted_*.mat');
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
outputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검_weighted\\correlation_matrices_weighted_%s.xlsx', dateStr);

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

%% 6. 요약 테이블 생성 및 저장
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
matFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검_weighted\\correlation_workspace_weighted_%s.mat', dateStr);
if ~isempty(fieldnames(correlationMatrices))
    save(matFileName, 'correlationMatrices', 'periods', 'summaryTable', 'allData');
else
    save(matFileName, 'periods', 'allData');  % correlationMatrices가 없을 때
end

%% 7. 성과 관련 문항 종합점수 분석 (역량검사와의 상관분석)
fprintf('\n[6단계] 성과 관련 문항 종합점수 분석 (역량검사와의 상관분석)\n');
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

%% 8. 성과점수 분석 결과 저장
if ~isempty(fieldnames(performanceResults))
    fprintf('\n[7단계] 역량검사-성과점수 상관분석 결과 저장\n');
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

%% 9. 종합 성과점수 분석 (5개 시점 통합)
fprintf('\n[8단계] 종합 성과점수 분석 (5개 시점 통합)\n');
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

%% 10. 분포 기반 시각화
fprintf('\n[9단계] 분포 기반 시각화\n');
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

%% 11. 최종 요약 출력
fprintf('\n========================================\n');
fprintf('Period별 상관 매트릭스 생성 완료 (가중치 적용 버전)\n');
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

% 원래 디렉토리로 복귀
cd('D:\project\HR데이터\matlab\문항기반_weighted');
fprintf('\n✅ 분석 완료 - 원본 디렉토리로 복귀\n');

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

function [competencyScore, usedColumnName] = findCompetencyScoreColumn(competencyTestData, testIdx)
    % 역량검사 종합점수 컬럼을 찾는 개선된 함수
    %
    % 우선순위:
    % 1. 정확한 매치: 'Average_Competency_Score'
    % 2. 키워드 매치: '총점', '종합점수', '평균점수', '총합', '합계', 'total', 'average', 'score'
    % 3. 숫자형 컬럼 중 가장 적절한 것 (ID 제외, 분산이 있는 것)

    competencyScore = [];
    usedColumnName = '';
    colNames = competencyTestData.Properties.VariableNames;

    % 1단계: 정확한 매치
    exactMatches = {'Average_Competency_Score', 'CompetencyScore', 'Competency_Score'};
    for i = 1:length(exactMatches)
        if ismember(exactMatches{i}, colNames)
            competencyScore = competencyTestData.(exactMatches{i})(testIdx);
            usedColumnName = exactMatches{i};
            return;
        end
    end

    % 2단계: 키워드 매치 (한글 + 영문)
    scoreKeywords = {'총점', '종합점수', '평균점수', '총합', '합계', 'total', 'average', 'score', '점수'};

    for col = 1:width(competencyTestData)
        colName = colNames{col};
        colNameLower = lower(colName);

        % ID 컬럼은 제외
        if contains(colNameLower, {'id', '사번', 'empno'})
            continue;
        end

        % 키워드 매치 확인
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

    % 3단계: 숫자형 컬럼 중 가장 적절한 것 찾기
    fprintf('    키워드 매치 실패 - 숫자형 컬럼 탐색 중...\n');

    candidateColumns = {};
    candidateScores = [];

    for col = 2:width(competencyTestData)  % ID 컬럼(첫 번째) 제외
        colName = colNames{col};
        colNameLower = lower(colName);
        colData = competencyTestData{:, col};

        % ID 관련 컬럼 제외
        if contains(colNameLower, {'id', '사번', 'empno', 'employee'})
            continue;
        end

        % 숫자형이고 분산이 있는 컬럼만
        if isnumeric(colData) && ~all(isnan(colData))
            colVariance = var(colData, 'omitnan');
            if colVariance > 0
                candidateColumns{end+1} = colName;

                % 점수 매기기 (더 적절한 컬럼일수록 높은 점수)
                score = 0;

                % 평균이 합리적인 범위에 있는가 (1~100 사이)
                colMean = mean(colData, 'omitnan');
                if colMean >= 1 && colMean <= 100
                    score = score + 3;
                elseif colMean >= 0.1 && colMean <= 10
                    score = score + 2;
                end

                % 분산이 적절한가
                if colVariance > 0.1 && colVariance < 1000
                    score = score + 2;
                end

                % 결측치가 적은가
                missingRate = sum(isnan(colData)) / length(colData);
                if missingRate < 0.1
                    score = score + 1;
                end

                candidateScores(end+1) = score;

                fprintf('      후보 컬럼: "%s" (평균: %.2f, 분산: %.2f, 결측: %.1f%%, 점수: %.1f)\n', ...
                    colName, colMean, colVariance, missingRate*100, score);
            end
        end
    end

    % 가장 높은 점수의 컬럼 선택
    if ~isempty(candidateColumns)
        [~, bestIdx] = max(candidateScores);
        bestColumn = candidateColumns{bestIdx};
        competencyScore = competencyTestData.(bestColumn)(testIdx);
        usedColumnName = bestColumn;

        fprintf('      선택된 컬럼: "%s" (점수: %.1f)\n', bestColumn, candidateScores(bestIdx));
    else
        fprintf('      적절한 숫자형 컬럼을 찾을 수 없습니다\n');
    end
end

function [matchedData, matchedIDs, sampleSize] = matchWithCompetencyTest(questionData, responseIDs, competencyTestData)
    % ID를 문자열로 통일
    if isnumeric(competencyTestData.ID)
        testIDs = arrayfun(@num2str, competencyTestData.ID, 'UniformOutput', false);
    else
        testIDs = cellfun(@char, competencyTestData.ID, 'UniformOutput', false);
    end

    % 교집합 찾기
    [commonIDs, responseIdx, testIdx] = intersect(responseIDs, testIDs);

    % 매칭된 데이터 구성
    if length(commonIDs) >= 5
        matchedQuestionData = questionData(responseIdx, :);

        % 역량검사 종합점수 컬럼 찾기 (개선된 로직)
        [competencyScore, usedColumnName] = findCompetencyScoreColumn(competencyTestData, testIdx);

        if ~isempty(competencyScore)
            fprintf('    사용된 역량점수 컬럼: "%s"\n', usedColumnName);
            matchedData = [matchedQuestionData, competencyScore];
            matchedIDs = commonIDs;
            sampleSize = length(commonIDs);
        else
            fprintf('    [경고] 적절한 역량점수 컬럼을 찾을 수 없습니다\n');
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

    % 변수명 설정
    variableNames = [questionCols, {'CompetencyTest_Total'}];

    % 결측치 처리: 행별로 50% 이상 결측이면 제거
    validRows = sum(isnan(matchedData), 2) < (size(matchedData, 2) * 0.5);
    cleanData = matchedData(validRows, :);

    % 분산이 0인 변수 제거
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

    % 상관계수 매트릭스 계산
    try
        correlationMatrix = corrcoef(cleanData, 'Rows', 'pairwise');

        % p-value 계산
        n = size(cleanData, 1);
        tStat = correlationMatrix .* sqrt((n-2) ./ (1 - correlationMatrix.^2));
        pValues = 2 * (1 - tcdf(abs(tStat), n-2));

        % 대각선 요소 보정 (자기 자신과의 상관은 p=0)
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

    % 마지막 컬럼이 종합점수
    lastColIdx = size(correlationMatrix, 2);
    questionCorrs = correlationMatrix(1:end-1, lastColIdx);
    questionPvals = pValues(1:end-1, lastColIdx);

    % 상위 상관계수 출력
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

% 성과 관련 문항의 상관계수를 별도로 출력하는 함수 추가
function displayPerformanceQuestionCorrelations(correlationMatrix, pValues, questionCols, performanceQuestions)
    if size(correlationMatrix, 2) < 2
        return;
    end

    % 마지막 컬럼이 종합점수
    lastColIdx = size(correlationMatrix, 2);

    fprintf('  성과 관련 문항의 역량검사 종합점수와의 상관:\n');

    % 성과 관련 문항들만 찾아서 출력
    foundAny = false;
    for i = 1:length(performanceQuestions)
        perfQ = performanceQuestions{i};
        qIdx = find(strcmp(questionCols, perfQ));

        if ~isempty(qIdx)
            corr = correlationMatrix(qIdx, lastColIdx);
            pval = pValues(qIdx, lastColIdx);

            sig_str = '';
            if pval < 0.001, sig_str = '***';
            elseif pval < 0.01, sig_str = '**';
            elseif pval < 0.05, sig_str = '*';
            end

            fprintf('    %s: r=%.3f (p=%.3f) %s\n', perfQ, corr, pval, sig_str);
            foundAny = true;
        end
    end

    if ~foundAny
        fprintf('    (성과 관련 문항이 데이터에 없음)\n');
    end
end

function createDistributionVisualizations(correlationMatrices, periods, periodFields, performanceResults, integratedPerformanceData, overallCorrelation, competencyTestData)
    if isempty(periodFields)
        return;
    end

    fprintf('▶ 분포 기반 시각화 생성 중...\n');

    %% 1. 역량검사점수 분포 히스토그램
    figure('Name', '역량검사점수 분포', 'Position', [100, 100, 1400, 900]);

    % 서브플롯 배치: 2x3 (히스토그램 5개 + 종합 히스토그램 1개)
    for i = 1:length(periodFields)
        fieldName = periodFields{i};
        periodNum = str2double(fieldName(end));
        result = correlationMatrices.(fieldName);

        subplot(2, 3, i);
        competencyData = result.cleanData(:, end);
        validCompetencyData = competencyData(~isnan(competencyData));  % NaN 값 제거
        histogram(validCompetencyData, 20);  % 역량검사 총점 히스토그램 (20개 구간)
        title(sprintf('%s 역량검사점수 분포', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
        xlabel('역량검사점수');
        ylabel('빈도');
        grid on;

        % 통계량 표시 (NaN 값 제외)
        meanScore = nanmean(result.cleanData(:, end));
        stdScore = nanstd(result.cleanData(:, end));
        text(0.6, 0.8, sprintf('평균: %.1f\n표준편차: %.1f\nN: %d', meanScore, stdScore, length(validCompetencyData)), ...
             'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white');
    end

    fprintf('✓ 역량검사점수 분포 히스토그램 생성 완료\n');

    %% 2. 성과점수 분포 히스토그램 (성과점수 분석이 있는 경우)
    if ~isempty(fieldnames(performanceResults))
        figure('Name', '성과점수 분포', 'Position', [150, 150, 1400, 900]);

        perfFields = fieldnames(performanceResults);
        numPlots = length(perfFields) + 1;  % 각 시점 + 종합

        for i = 1:length(perfFields)
            fieldName = perfFields{i};
            periodNum = str2double(fieldName(end));
            result = performanceResults.(fieldName);

            subplot(2, 3, i);
            validPerformanceScores = result.performanceScores(~isnan(result.performanceScores));  % NaN 값 제거
            histogram(validPerformanceScores, 15);
            title(sprintf('%s 성과점수 분포', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
            xlabel('성과점수');
            ylabel('빈도');
            grid on;

            meanScore = result.performanceMean;
            stdScore = result.performanceStd;
            text(0.6, 0.8, sprintf('평균: %.2f\n표준편차: %.2f\nN: %d', meanScore, stdScore, result.sampleSize), ...
                 'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white');
        end

        fprintf('✓ 성과점수 분포 히스토그램 생성 완료\n');
    end

    %% 3. 역량검사점수 vs 성과점수 산점도 (선형추세선 포함)
    if ~isempty(fieldnames(performanceResults))
        figure('Name', '역량검사점수 vs 성과점수 산점도', 'Position', [200, 200, 1400, 900]);

        perfFields = fieldnames(performanceResults);

        for i = 1:length(perfFields)
            fieldName = perfFields{i};
            periodNum = str2double(fieldName(end));
            result = performanceResults.(fieldName);

            subplot(2, 3, i);

            % 산점도 그리기
            scatter(result.competencyTestScores, result.performanceScores, 50, 'filled');
            hold on;

            % 선형추세선 그리기
            if length(result.competencyTestScores) > 1
                p = polyfit(result.competencyTestScores, result.performanceScores, 1);
                x_trend = linspace(min(result.competencyTestScores), max(result.competencyTestScores), 100);
                y_trend = polyval(p, x_trend);
                plot(x_trend, y_trend, 'r-', 'LineWidth', 2);
            end

            title(sprintf('%s: 역량검사 vs 성과점수', strrep(periods{periodNum}, '_', ' ')), 'FontSize', 12, 'FontWeight', 'bold');
            xlabel('역량검사점수');
            ylabel('성과점수');
            grid on;

            % 상관계수 표시
            corrText = sprintf('r = %.3f\np = %.3f\nN = %d', result.correlation, result.pValue, result.sampleSize);
            text(0.05, 0.95, corrText, 'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white', 'VerticalAlignment', 'top');

            hold off;
        end

        % 종합 상관분석 산점도
        if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData) && all(ismember({'CompetencyScore', 'PerformanceScore'}, integratedPerformanceData.Properties.VariableNames))
            subplot(2, 3, 6);

            scatter(integratedPerformanceData.CompetencyScore, integratedPerformanceData.PerformanceScore, 50, 'filled', 'MarkerFaceColor', [0.2 0.6 0.8]);
            hold on;

            % 선형추세선
            if height(integratedPerformanceData) > 1
                p = polyfit(integratedPerformanceData.CompetencyScore, integratedPerformanceData.PerformanceScore, 1);
                x_trend = linspace(min(integratedPerformanceData.CompetencyScore), max(integratedPerformanceData.CompetencyScore), 100);
                y_trend = polyval(p, x_trend);
                plot(x_trend, y_trend, 'r-', 'LineWidth', 2);
            end

            title('종합: 역량검사 vs 성과점수', 'FontSize', 12, 'FontWeight', 'bold');
            xlabel('역량검사점수');
            ylabel('종합 성과점수');
            grid on;

            % 종합 상관계수 표시
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

    %% 4. 별도 figure: 전체 역량검사점수 분포
    figure('Name', '전체 역량검사점수 분포', 'Position', [300, 300, 800, 600]);

    % 역량검사 데이터에서 고유한 개인의 점수만 사용 (중복 제거)
    [~, usedColumnName] = findCompetencyScoreColumn(competencyTestData, 1:height(competencyTestData));
    if ~isempty(usedColumnName)
        allUniqueCompetencyScores = competencyTestData.(usedColumnName);
        validAllCompetencyScores = allUniqueCompetencyScores(~isnan(allUniqueCompetencyScores));  % NaN 값 제거

        histogram(validAllCompetencyScores, 30);
        title('전체 역량검사점수 분포', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('역량검사점수 (표준화된 값)', 'FontSize', 12);
        ylabel('빈도', 'FontSize', 12);
        grid on;

        meanScore = nanmean(validAllCompetencyScores);
        stdScore = nanstd(validAllCompetencyScores);

        % 통계 정보를 텍스트 박스로 표시
        textStr = sprintf('평균: %.3f\n표준편차: %.3f\nN: %d명\n범위: %.3f ~ %.3f', ...
                         meanScore, stdScore, length(validAllCompetencyScores), ...
                         min(validAllCompetencyScores), max(validAllCompetencyScores));
        text(0.02, 0.98, textStr, 'Units', 'normalized', 'FontSize', 11, ...
             'BackgroundColor', 'white', 'EdgeColor', 'black', ...
             'VerticalAlignment', 'top', 'HorizontalAlignment', 'left');

        fprintf('✓ 전체 역량검사점수 분포 별도 figure 생성 완료\n');
    end

    %% 5. 별도 figure: 종합 성과점수 분포 (5개 시점 통합 - 상관분석용)
    if ~isempty(integratedPerformanceData) && istable(integratedPerformanceData) && ismember('PerformanceScore', integratedPerformanceData.Properties.VariableNames)
        figure('Name', '종합 성과점수 분포 (상관분석용)', 'Position', [400, 400, 800, 600]);

        validIntegratedScores = integratedPerformanceData.PerformanceScore(~isnan(integratedPerformanceData.PerformanceScore));

        histogram(validIntegratedScores, 25);
        title('종합 성과점수 분포 (5개 시점 통합 - 상관분석용)', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('종합 성과점수 (표준화된 값)', 'FontSize', 12);
        ylabel('빈도', 'FontSize', 12);
        grid on;

        meanScore = nanmean(integratedPerformanceData.PerformanceScore);
        stdScore = nanstd(integratedPerformanceData.PerformanceScore);
        numPeriods = nanmean(integratedPerformanceData.NumPeriods);

        % 통계 정보를 텍스트 박스로 표시
        textStr = sprintf('평균: %.3f\n표준편차: %.3f\nN: %d명\n평균 참여횟수: %.1f회\n범위: %.3f ~ %.3f', ...
                         meanScore, stdScore, length(validIntegratedScores), ...
                         numPeriods, min(validIntegratedScores), max(validIntegratedScores));
        text(0.02, 0.98, textStr, 'Units', 'normalized', 'FontSize', 11, ...
             'BackgroundColor', 'white', 'EdgeColor', 'black', ...
             'VerticalAlignment', 'top', 'HorizontalAlignment', 'left');

        fprintf('✓ 종합 성과점수 분포 (상관분석용) 별도 figure 생성 완료\n');
    end

    fprintf('✓ 모든 분포 기반 시각화 생성 완료\n');
end

function standardizedData = standardizeQuestionScales(questionData, questionNames, periodNum)
    standardizedData = questionData;

    for i = 1:size(questionData, 2)
        columnData = questionData(:, i);
        validData = columnData(~isnan(columnData));

        if isempty(validData)
            continue;
        end

        % 항상 실제 데이터의 최소/최대값 사용
        actualMin = min(validData);
        actualMax = max(validData);

        % 디버깅 정보
        fprintf('    %s: 실제 범위 [%.1f, %.1f]', questionNames{i}, actualMin, actualMax);

        % 척도 정보 확인 (하지만 실제 데이터가 범위를 벗어나면 조정)
        scaleInfo = getQuestionScaleInfo(questionNames, periodNum);
        if isfield(scaleInfo, questionNames{i})
            theoreticalMin = scaleInfo.(questionNames{i}).min;
            theoreticalMax = scaleInfo.(questionNames{i}).max;

            % 실제 데이터가 이론적 범위를 벗어나면 실제 범위 사용
            minScale = min(theoreticalMin, actualMin);
            maxScale = max(theoreticalMax, actualMax);
            fprintf(' → 조정된 척도 [%.1f, %.1f]\n', minScale, maxScale);
        else
            % 척도 정보가 없으면 실제 데이터 범위 사용
            minScale = actualMin;
            maxScale = actualMax;
            fprintf(' → 실제 데이터 범위 사용\n');
        end

        % Min-Max 스케일링
        if maxScale > minScale
            standardizedData(:, i) = (columnData - minScale) / (maxScale - minScale);

            % 검증
            validStandardized = standardizedData(~isnan(standardizedData(:, i)), i);
            if max(validStandardized) > 1.0001 || min(validStandardized) < -0.0001
                warning('표준화 오류: %s [%.4f, %.4f]', ...
                    questionNames{i}, min(validStandardized), max(validStandardized));
            end
        else
            standardizedData(:, i) = 0.5 * ones(size(columnData));
            standardizedData(isnan(columnData), i) = NaN;
        end
    end
end

function scaleInfo = getQuestionScaleInfo(questionNames, periodNum)
    % 문항별 리커트 척도 정보를 반환하는 함수
    % 실제 questionInfo 테이블이 있다면 여기서 로드

    scaleInfo = struct();

    % questionInfo 테이블 로드 시도
    try
        % 일반적인 척도 정보 파일들을 찾아보기
        possibleFiles = {
            'D:\project\HR데이터\데이터\questionInfo.xlsx',
            'D:\project\HR데이터\데이터\문항정보.xlsx',
            'questionInfo.xlsx',
            '문항정보.xlsx'
        };

        questionInfo = [];
        for i = 1:length(possibleFiles)
            if exist(possibleFiles{i}, 'file')
                try
                    questionInfo = readtable(possibleFiles{i});
                    fprintf('    ✓ 문항 척도 정보 로드: %s\n', possibleFiles{i});
                    break;
                catch
                    continue;
                end
            end
        end

        if ~isempty(questionInfo) && height(questionInfo) > 0
            % 테이블에서 척도 정보 추출
            for i = 1:height(questionInfo)
                if ismember('QuestionID', questionInfo.Properties.VariableNames) && ...
                   ismember('MinScale', questionInfo.Properties.VariableNames) && ...
                   ismember('MaxScale', questionInfo.Properties.VariableNames)

                    qid = questionInfo.QuestionID{i};
                    minVal = questionInfo.MinScale(i);
                    maxVal = questionInfo.MaxScale(i);

                    if ismember(qid, questionNames)
                        scaleInfo.(qid) = struct('min', minVal, 'max', maxVal);
                    end
                end
            end
        end
    catch
        % 파일 로드 실패시 무시
    end

    % 기본 척도 정보 (일반적인 문항들)
    defaultScales = getDefaultQuestionScales(periodNum);

    % 기본값으로 채우기
    for i = 1:length(questionNames)
        qName = questionNames{i};
        if ~isfield(scaleInfo, qName)
            if isfield(defaultScales, qName)
                scaleInfo.(qName) = defaultScales.(qName);
            else
                % 완전 기본값: 1~5 척도
                scaleInfo.(qName) = struct('min', 1, 'max', 5);
            end
        end
    end
end

function defaultScales = getDefaultQuestionScales(periodNum)
    % 시점별 문항의 기본 척도 정보
    % 실제 데이터에 맞게 수정 필요

    defaultScales = struct();

    % 일반적인 성과 관련 문항들의 척도 (예시)
    performanceQuestions = {'Q3', 'Q4', 'Q5', 'Q21', 'Q22', 'Q23', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33', 'Q34', 'Q45', 'Q46', 'Q51'};

    for i = 1:length(performanceQuestions)
        qName = performanceQuestions{i};

        % 문항별 특별한 척도가 있다면 여기서 설정
        if ismember(qName, {'Q3', 'Q4', 'Q5'})
            defaultScales.(qName) = struct('min', 1, 'max', 4); % 4점 척도
        elseif ismember(qName, {'Q45', 'Q46', 'Q51'})
            defaultScales.(qName) = struct('min', 1, 'max', 7); % 7점 척도
        else
            defaultScales.(qName) = struct('min', 1, 'max', 5); % 5점 척도 (기본)
        end
    end
end

function standardizedScores = standardizeCompetencyScores(rawScores)
    % 역량검사 점수 min-max 표준화 함수

    validScores = rawScores(~isnan(rawScores));

    if isempty(validScores)
        standardizedScores = rawScores;
        return;
    end

    minScore = min(validScores);
    maxScore = max(validScores);

    if maxScore > minScore
        standardizedScores = (rawScores - minScore) / (maxScore - minScore);
    else
        standardizedScores = 0.5 * ones(size(rawScores));
        standardizedScores(isnan(rawScores)) = NaN;
    end

    fprintf('    ✓ 역량검사 점수 표준화 완료 (원본 범위: %.1f~%.1f → [0,1])\n', minScore, maxScore);
end

function perfSummaryTable = createPerformanceSummaryTable(performanceResults, periods)
    % 역량검사-성과점수 상관분석 요약 테이블을 생성하는 함수
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

        % 역량검사점수 통계
        newRow.CompetencyMean = result.competencyMean;
        newRow.CompetencyStd = result.competencyStd;

        % 성과점수 통계
        newRow.PerformanceMean = result.performanceMean;
        newRow.PerformanceStd = result.performanceStd;

        % 상관분석 결과
        newRow.Correlation = result.correlation;
        newRow.PValue = result.pValue;

        % 유의성 판정
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
    % 5개 시점의 성과점수를 개인별로 통합하고 역량검사점수와 상관분석하는 함수

    fprintf('▶ 개인별 성과점수 통합 중...\n');

    % 역량검사 데이터에서 ID 추출
    if isnumeric(competencyTestData.ID)
        testIDs = arrayfun(@num2str, competencyTestData.ID, 'UniformOutput', false);
    else
        testIDs = cellfun(@char, competencyTestData.ID, 'UniformOutput', false);
    end

    % 각 시점별 성과점수 데이터 수집
    perfFields = fieldnames(performanceResults);
    allPerformanceData = table();

    fprintf('  - 수집 중인 시점: ');
    for i = 1:length(perfFields)
        fieldName = perfFields{i};
        periodNum = str2double(fieldName(end));
        result = performanceResults.(fieldName);

        fprintf('%s ', periods{periodNum});

        % 해당 시점의 성과점수 데이터 가져오기
        if isfield(result, 'cleanIDs')
            % performanceResults에 저장된 ID와 성과점수 사용
            periodIDs = result.cleanIDs;

            % 성과점수와 매칭
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

    % 개인별로 성과점수 평균 계산
    fprintf('  - 개인별 성과점수 평균 계산 중...\n');

    uniqueIDs = unique(allPerformanceData.ID);
    integratedTable = table();

    validCount = 0;
    for i = 1:length(uniqueIDs)
        personID = uniqueIDs{i};
        personData = allPerformanceData(strcmp(allPerformanceData.ID, personID), :);

        % 최소 1개 시점 이상의 데이터가 있는 경우 포함 (단일 시점도 포함)
        if height(personData) >= 1
            avgPerformanceScore = nanmean(personData.PerformanceScore);  % NaN 값 무시하고 평균 계산
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

    % 역량검사점수와 매칭
    fprintf('  - 역량검사점수와 매칭 중...\n');

    % 역량검사 점수 컬럼 찾기 (기존 함수 재사용)
    dummyTestIdx = 1:height(competencyTestData);
    [~, usedColumnName] = findCompetencyScoreColumn(competencyTestData, dummyTestIdx);

    if isempty(usedColumnName)
        fprintf('  [경고] 역량검사 점수 컬럼을 찾을 수 없습니다.\n');
        integratedData = [];
        overallCorrelation = struct();
        return;
    end

    rawCompetencyScores = competencyTestData.(usedColumnName);

    % 역량검사 점수 표준화
    competencyScores = standardizeCompetencyScores(rawCompetencyScores);

    % ID 매칭
    [commonIDs, integratedIdx, testIdx] = intersect(integratedTable.ID, testIDs);

    if length(commonIDs) < 3
        fprintf('  [경고] 매칭된 데이터가 부족합니다 (%d명)\n', length(commonIDs));
        integratedData = [];
        overallCorrelation = struct();
        return;
    end

    % 최종 분석 데이터 구성
    finalTable = table();
    finalTable.ID = commonIDs;
    finalTable.CompetencyScore = competencyScores(testIdx);
    finalTable.PerformanceScore = integratedTable.IntegratedPerformanceScore(integratedIdx);  % 시각화를 위해 PerformanceScore로 명명
    finalTable.IntegratedPerformanceScore = integratedTable.IntegratedPerformanceScore(integratedIdx);  % 기존 이름도 유지
    finalTable.NumPeriods = integratedTable.NumPeriods(integratedIdx);

    % 각 시점별 성과점수도 추가
    for p = 1:length(periods)
        colName = sprintf('Performance_%s', periods{p});
        colData = nan(height(finalTable), 1);

        for j = 1:height(finalTable)
            personID = finalTable.ID{j};
            periodData = allPerformanceData(strcmp(allPerformanceData.ID, personID) & strcmp(allPerformanceData.Period, periods{p}), :);
            if height(periodData) > 0
                colData(j) = periodData.PerformanceScore(1);
            end
        end

        finalTable.(colName) = colData;
    end

    % 상관분석 수행
    validRows = ~isnan(finalTable.CompetencyScore) & ~isnan(finalTable.IntegratedPerformanceScore);
    cleanCompetencyScores = finalTable.CompetencyScore(validRows);
    cleanPerformanceScores = finalTable.IntegratedPerformanceScore(validRows);

    if length(cleanCompetencyScores) < 3
        fprintf('  [경고] 상관분석을 위한 최종 데이터가 부족합니다 (%d명)\n', length(cleanCompetencyScores));
        integratedData = finalTable;
        overallCorrelation = struct();
        return;
    end

    % 상관계수 계산
    [corrCoeff, pValue] = corr(cleanCompetencyScores, cleanPerformanceScores);

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
    overallCorrelation.sampleSize = length(cleanCompetencyScores);
    overallCorrelation.competencyMean = mean(cleanCompetencyScores);
    overallCorrelation.competencyStd = std(cleanCompetencyScores);
    overallCorrelation.performanceMean = mean(cleanPerformanceScores);
    overallCorrelation.performanceStd = std(cleanPerformanceScores);

    integratedData = finalTable;

    fprintf('  ✓ 개인별 성과점수 통합 및 상관분석 완료: %d명\n', length(cleanCompetencyScores));
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