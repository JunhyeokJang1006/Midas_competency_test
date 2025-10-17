%% 역량검사 vs 역량진단 상관분석 및 단순회귀 (최종 완전판)
% 
% 작성일: 2025년
% 목적: 역량검사 평균 점수와 역량진단 요인분석 점수 간 관계 분석
% 
% 주요 기능:
% 1. 기존 역량진단 결과 로드
% 2. 역량검사 데이터 로드 및 전처리
% 3. 데이터 매칭
% 4. 상관분석
% 5. 단순회귀분석 (역량검사 → 역량진단)
% 6. 결과 저장 및 해석

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('역량검사 vs 역량진단 상관분석 및 회귀분석\n');
fprintf('========================================\n\n');

%% 1. 기존 분석 결과 로드
fprintf('[1단계] 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

consolidatedScores = [];
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};

% MAT 파일에서 역량진단 결과 로드
matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    matFileName = matFiles(idx).name;
    
    fprintf('MAT 파일 로드: %s\n', matFileName);
    try
        loadedData = load(matFileName, 'consolidatedScores', 'periods');
        if isfield(loadedData, 'consolidatedScores')
            consolidatedScores = loadedData.consolidatedScores;
        end
        if isfield(loadedData, 'periods')
            periods = loadedData.periods;
        end
        
        if istable(consolidatedScores)
            fprintf('✓ 역량진단 데이터 로드 완료: %d명\n', height(consolidatedScores));
            fprintf('  - 컬럼 수: %d개\n', width(consolidatedScores));
        else
            fprintf('✗ 역량진단 데이터가 테이블 형태가 아닙니다.\n');
            consolidatedScores = [];
        end
    catch ME
        fprintf('✗ MAT 파일 로드 실패: %s\n', ME.message);
        consolidatedScores = [];
    end
end

% MAT 파일 로드에 실패한 경우 Excel 파일 시도
if isempty(consolidatedScores)
    fprintf('\nExcel 파일에서 데이터 로드를 시도합니다.\n');
    
    excelFiles = dir('competency_performance_correlation_results_*.xlsx');
    if ~isempty(excelFiles)
        [~, idx] = max([excelFiles.datenum]);
        excelFileName = excelFiles(idx).name;
        
        fprintf('Excel 파일 로드: %s\n', excelFileName);
        try
            consolidatedScores = readtable(excelFileName, 'Sheet', '역량진단_통합점수', 'VariableNamingRule', 'preserve');
            fprintf('✓ 역량진단 데이터 로드 완료: %d명\n', height(consolidatedScores));
        catch ME
            fprintf('✗ Excel 파일 로드 실패: %s\n', ME.message);
            consolidatedScores = [];
        end
    end
end

if isempty(consolidatedScores)
    fprintf('✗ 분석 결과 파일을 찾을 수 없습니다.\n');
    fprintf('먼저 역량진단 분석을 실행해주세요.\n');
    return;
end

%% 2. 역량검사 데이터 로드
fprintf('\n[2단계] 역량검사 데이터 로드\n');
fprintf('----------------------------------------\n');

competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    % 파일의 시트 정보 확인
    [~, sheets] = xlsfinfo(competencyTestPath);
    fprintf('발견된 시트: %s\n', strjoin(sheets, ', '));
    
    % '역량검사_종합점수' 시트 우선 시도
    sheetToLoad = '역량검사_종합점수';
    if ismember(sheetToLoad, sheets)
        fprintf('로드할 시트: %s\n', sheetToLoad);
        competencyTestData = readtable(competencyTestPath, 'Sheet', sheetToLoad, 'VariableNamingRule', 'preserve');
    else
        % 대안으로 적절한 시트 찾기
        for i = 1:length(sheets)
            if contains(lower(sheets{i}), {'역량', 'competency', '점수', 'score'})
                sheetToLoad = sheets{i};
                break;
            end
        end
        
        if strcmp(sheetToLoad, '역량검사_종합점수')  % 여전히 찾지 못한 경우
            sheetToLoad = sheets{1};  % 첫 번째 시트 사용
        end
        
        fprintf('로드할 시트: %s\n', sheetToLoad);
        competencyTestData = readtable(competencyTestPath, 'Sheet', sheetToLoad, 'VariableNamingRule', 'preserve');
    end
    
    fprintf('✓ 역량검사 데이터 로드 완료: %d명, %d컬럼\n', height(competencyTestData), width(competencyTestData));
    fprintf('컬럼명: %s\n', strjoin(competencyTestData.Properties.VariableNames, ', '));
    
catch ME
    fprintf('✗ 역량검사 데이터 로드 실패: %s\n', ME.message);
    return;
end

%% 3. 역량검사 데이터 전처리
fprintf('\n[3단계] 역량검사 데이터 전처리\n');
fprintf('----------------------------------------\n');

% ID 컬럼 찾기
testIdCol = findIDColumn(competencyTestData);
testIDs = extractAndStandardizeIDs(competencyTestData{:, testIdCol});
fprintf('✓ ID 컬럼 사용: %s (%d명)\n', competencyTestData.Properties.VariableNames{testIdCol}, length(testIDs));

% 역량 점수 컬럼 확인
colNames = competencyTestData.Properties.VariableNames;
competencyScoreCols = {};

fprintf('\n▶ 역량 점수 컬럼 확인\n');
for col = 1:width(competencyTestData)
    colName = colNames{col};
    colData = competencyTestData{:, col};
    
    % ID 컬럼이 아니고 숫자형인 경우
    if col ~= testIdCol && isnumeric(colData)
        validCount = sum(~isnan(colData));
        
        if validCount > 0
            competencyScoreCols{end+1} = colName;
            fprintf('  ✓ %s: %d명 유효 (평균: %.2f, 범위: %.1f~%.1f)\n', ...
                colName, validCount, nanmean(colData), nanmin(colData), nanmax(colData));
        end
    end
end

fprintf('총 발견된 역량 점수 컬럼: %d개\n', length(competencyScoreCols));

if isempty(competencyScoreCols)
    fprintf('✗ 역량 점수 컬럼을 찾을 수 없습니다.\n');
    return;
end

%% 4. 역량 평균 점수 계산
fprintf('\n▶ 역량 평균 점수 계산\n');

competencyTestScores = table();
competencyTestScores.ID = testIDs;

% 개별 역량 점수들을 테이블에 추가
for i = 1:length(competencyScoreCols)
    colName = competencyScoreCols{i};
    competencyTestScores.(colName) = competencyTestData.(colName);
end

% 전체 역량 평균 점수 계산
competencyMatrix = table2array(competencyTestScores(:, competencyScoreCols));
competencyTestScores.Average_Competency_Score = mean(competencyMatrix, 2, 'omitnan');
competencyTestScores.Valid_Competency_Count = sum(~isnan(competencyMatrix), 2);

% 통계 요약
validAvgCount = sum(~isnan(competencyTestScores.Average_Competency_Score));
fprintf('✓ 역량 평균 점수 계산 완료\n');
fprintf('  - 유효한 평균 점수: %d명 / %d명 (%.1f%%)\n', ...
    validAvgCount, height(competencyTestScores), 100*validAvgCount/height(competencyTestScores));

if validAvgCount > 0
    avgScores = competencyTestScores.Average_Competency_Score(~isnan(competencyTestScores.Average_Competency_Score));
    fprintf('  - 전체 평균: %.3f (±%.3f)\n', mean(avgScores), std(avgScores));
    fprintf('  - 범위: %.3f ~ %.3f\n', min(avgScores), max(avgScores));
end

%% 5. 데이터 매칭
fprintf('\n[4단계] 데이터 매칭\n');
fprintf('----------------------------------------\n');

% ID를 문자열로 통일
if isnumeric(consolidatedScores.ID)
    diagnosticIDs = arrayfun(@num2str, consolidatedScores.ID, 'UniformOutput', false);
else
    diagnosticIDs = cellfun(@char, consolidatedScores.ID, 'UniformOutput', false);
end

if isnumeric(competencyTestScores.ID)
    testIDs_str = arrayfun(@num2str, competencyTestScores.ID, 'UniformOutput', false);
else
    testIDs_str = cellfun(@char, competencyTestScores.ID, 'UniformOutput', false);
end

% 교집합 찾기
[commonIDs, diagnosticIdx, testIdx] = intersect(diagnosticIDs, testIDs_str);

fprintf('매칭 결과:\n');
fprintf('  - 역량진단 데이터: %d명\n', height(consolidatedScores));
fprintf('  - 역량검사 데이터: %d명\n', height(competencyTestScores));
fprintf('  - 공통 ID: %d명\n', length(commonIDs));
fprintf('  - 매칭률: %.1f%%\n', 100 * length(commonIDs) / min(height(consolidatedScores), height(competencyTestScores)));

if length(commonIDs) < 5
    fprintf('✗ 매칭된 데이터가 너무 적습니다 (최소 5명 필요).\n');
    fprintf('ID 형식을 확인해주세요.\n');
    fprintf('샘플 ID (역량진단): %s\n', strjoin(diagnosticIDs(1:min(3, end)), ', '));
    fprintf('샘플 ID (역량검사): %s\n', strjoin(testIDs_str(1:min(3, end)), ', '));
    return;
end

% 매칭된 통합 데이터 생성
analysisData = table();
analysisData.ID = commonIDs;
analysisData.Test_Average = competencyTestScores.Average_Competency_Score(testIdx);

% 역량진단 점수 추가
if ismember('AverageStdScore', consolidatedScores.Properties.VariableNames)
    analysisData.Diagnostic_Average = consolidatedScores.AverageStdScore(diagnosticIdx);
end

% 시점별 점수도 추가
for p = 1:length(periods)
    colName = sprintf('Period%d_Score', p);
    if ismember(colName, consolidatedScores.Properties.VariableNames)
        analysisData.(sprintf('Diagnostic_Period%d', p)) = consolidatedScores.(colName)(diagnosticIdx);
    end
end

fprintf('✓ 통합 데이터 생성 완료: %d명\n', height(analysisData));

%% 6. 상관분석
fprintf('\n[5단계] 상관분석\n');
fprintf('----------------------------------------\n');

correlationResults = struct();

% 분석할 변수 쌍 정의
analysisVars = {};
if ismember('Diagnostic_Average', analysisData.Properties.VariableNames)
    analysisVars{end+1} = {'Test_Average', 'Diagnostic_Average', '역량검사 평균', '역량진단 전체 평균'};
end

% 시점별 분석도 추가
for p = 1:length(periods)
    diagVar = sprintf('Diagnostic_Period%d', p);
    if ismember(diagVar, analysisData.Properties.VariableNames)
        analysisVars{end+1} = {'Test_Average', diagVar, '역량검사 평균', sprintf('역량진단 %s', periods{p})};
    end
end

fprintf('분석할 변수 쌍: %d개\n\n', length(analysisVars));

% 상관분석 실행
for i = 1:length(analysisVars)
    varPair = analysisVars{i};
    xVar = varPair{1};
    yVar = varPair{2};
    xName = varPair{3};
    yName = varPair{4};
    
    if ismember(xVar, analysisData.Properties.VariableNames) && ...
       ismember(yVar, analysisData.Properties.VariableNames)
        
        xData = analysisData.(xVar);
        yData = analysisData.(yVar);
        
        % 유효한 데이터만 선택
        validIdx = ~isnan(xData) & ~isnan(yData);
        validCount = sum(validIdx);
        
        if validCount >= 5
            % 상관계수 계산
            r = corrcoef(xData(validIdx), yData(validIdx));
            correlation = r(1, 2);
            
            % 유의성 검정
            t_stat = correlation * sqrt((validCount - 2) / (1 - correlation^2));
            p_value = 2 * (1 - tcdf(abs(t_stat), validCount - 2));
            
            % 결과 출력
            fprintf('%s vs %s:\n', xName, yName);
            fprintf('  r = %.3f (n=%d, p=%.3f)', correlation, validCount, p_value);
            if p_value < 0.001
                fprintf(' ***');
            elseif p_value < 0.01
                fprintf(' **');
            elseif p_value < 0.05
                fprintf(' *');
            end
            fprintf('\n');
            
            % 효과크기 해석
            absR = abs(correlation);
            if absR >= 0.7
                fprintf('  → 강한 상관\n');
            elseif absR >= 0.5
                fprintf('  → 보통 상관\n');
            elseif absR >= 0.3
                fprintf('  → 약한 상관\n');
            else
                fprintf('  → 매우 약한 상관\n');
            end
            fprintf('\n');
            
            % 결과 저장
            resultKey = sprintf('%s_vs_%s', xVar, yVar);
            correlationResults.(resultKey) = struct(...
                'x_var', xVar, 'y_var', yVar, ...
                'x_name', xName, 'y_name', yName, ...
                'correlation', correlation, ...
                'n', validCount, 'p_value', p_value);
            
        else
            fprintf('%s vs %s: 데이터 부족 (n=%d)\n\n', xName, yName, validCount);
        end
    end
end



%% ===== 보조 함수 정의 =====
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

function idCol = findIDColumn(dataTable)
    idCol = [];
    colNames = dataTable.Properties.VariableNames;
    
    for col = 1:width(dataTable)
        colName = lower(colNames{col});
        colData = dataTable{:, col};
        
        % ID 컬럼 조건 확인
        isIDColumn = contains(colName, {'id', '사번', 'empno', 'employee'}) || ...
                     strcmp(colNames{col}, 'Var1');
        
        hasValidData = (isnumeric(colData) && ~all(isnan(colData))) || ...
                       (iscell(colData) && ~all(cellfun(@isempty, colData))) || ...
                       (isstring(colData) && ~all(ismissing(colData)));
        
        if isIDColumn && hasValidData
            idCol = col;
            break;
        end
    end
    
    % ID 컬럼을 찾지 못한 경우 첫 번째 컬럼 사용
    if isempty(idCol)
        idCol = 1;
    end
end