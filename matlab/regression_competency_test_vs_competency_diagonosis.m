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

%% 7. 단순회귀분석
fprintf('\n[6단계] 단순회귀분석: 역량검사 → 역량진단\n');
fprintf('========================================\n');

regressionResults = struct();

% 주요 회귀분석: 역량검사 평균 → 역량진단 평균
if ismember('Test_Average', analysisData.Properties.VariableNames) && ...
   ismember('Diagnostic_Average', analysisData.Properties.VariableNames)
    
    fprintf('▶ 주요 분석: 역량검사 평균 → 역량진단 평균\n');
    
    X = analysisData.Test_Average;
    Y = analysisData.Diagnostic_Average;
    
    % 결측치 제거
    validIdx = ~isnan(X) & ~isnan(Y);
    X_valid = X(validIdx);
    Y_valid = Y(validIdx);
    n = length(X_valid);
    
    fprintf('유효한 데이터: %d명\n', n);
    
    if n >= 5
        try
            % 회귀분석 실행
            X_matrix = [ones(n, 1), X_valid];
            beta = X_matrix \ Y_valid;
            
            intercept = beta(1);
            slope = beta(2);
            
            % 예측값 및 잔차
            Y_pred = X_matrix * beta;
            residuals = Y_valid - Y_pred;
            
            % 결정계수 (R²)
            SS_tot = sum((Y_valid - mean(Y_valid)).^2);
            SS_res = sum(residuals.^2);
            R_squared = 1 - (SS_res / SS_tot);
            R_squared_adj = 1 - ((SS_res/(n-2)) / (SS_tot/(n-1)));
            
            % 표준오차
            MSE = SS_res / (n - 2);
            SE_matrix = sqrt(MSE * inv(X_matrix' * X_matrix));
            SE_intercept = SE_matrix(1, 1);
            SE_slope = SE_matrix(2, 2);
            
            % t-검정
            t_slope = slope / SE_slope;
            t_intercept = intercept / SE_intercept;
            p_slope = 2 * (1 - tcdf(abs(t_slope), n-2));
            p_intercept = 2 * (1 - tcdf(abs(t_intercept), n-2));
            
            % F-검정
            F_stat = (SS_tot - SS_res) / (SS_res / (n-2));
            p_F = 1 - fcdf(F_stat, 1, n-2);
            
            % 결과 출력
            fprintf('\n회귀분석 결과:\n');
            fprintf('══════════════════════════════════════\n');
            fprintf('회귀식: 역량진단 = %.3f + %.3f × 역량검사\n', intercept, slope);
            fprintf('\n계수 추정치:\n');
            fprintf('  절편: %.3f (SE=%.3f, t=%.3f, p=%.3f)', intercept, SE_intercept, t_intercept, p_intercept);
            if p_intercept < 0.05, fprintf(' *'); end
            fprintf('\n');
            fprintf('  기울기: %.3f (SE=%.3f, t=%.3f, p=%.3f)', slope, SE_slope, t_slope, p_slope);
            if p_slope < 0.001, fprintf(' ***');
            elseif p_slope < 0.01, fprintf(' **');
            elseif p_slope < 0.05, fprintf(' *'); end
            fprintf('\n\n');
            
            fprintf('모형 적합도:\n');
            fprintf('  R² = %.3f (설명된 분산: %.1f%%)\n', R_squared, R_squared*100);
            fprintf('  조정된 R² = %.3f\n', R_squared_adj);
            fprintf('  F통계량 = %.3f (p=%.3f)', F_stat, p_F);
            if p_F < 0.001, fprintf(' ***');
            elseif p_F < 0.01, fprintf(' **');
            elseif p_F < 0.05, fprintf(' *'); end
            fprintf('\n');
            fprintf('  RMSE = %.3f\n\n', sqrt(MSE));
            
            % 실용적 해석
            fprintf('실용적 해석:\n');
            fprintf('══════════════════════════════════════\n');
            if p_slope < 0.05
                if slope > 0
                    fprintf('✓ 역량검사 점수가 1점 증가하면, 역량진단 점수가 %.3f점 증가합니다.\n', slope);
                    fprintf('✓ 역량검사는 역량진단 성과의 %.1f%%를 설명합니다.\n', R_squared*100);
                    
                    if ~isempty(X_valid)
                        example_increase = std(X_valid);
                        predicted_change = slope * example_increase;
                        fprintf('✓ 예시: 역량검사가 %.1f점(1SD) 증가 → 역량진단 %.3f점 증가 예상\n', ...
                            example_increase, predicted_change);
                    end
                else
                    fprintf('✓ 역량검사 점수가 1점 증가하면, 역량진단 점수가 %.3f점 감소합니다.\n', abs(slope));
                end
            else
                fprintf('✗ 역량검사와 역량진단 간에 통계적으로 유의한 예측 관계가 없습니다.\n');
                fprintf('  (p = %.3f > 0.05)\n', p_slope);
            end
            
            % 결과 저장
            regressionResults.main = struct(...
                'intercept', intercept, 'slope', slope, ...
                'SE_intercept', SE_intercept, 'SE_slope', SE_slope, ...
                't_intercept', t_intercept, 't_slope', t_slope, ...
                'p_intercept', p_intercept, 'p_slope', p_slope, ...
                'R_squared', R_squared, 'R_squared_adj', R_squared_adj, ...
                'F_stat', F_stat, 'p_F', p_F, 'RMSE', sqrt(MSE), 'n', n);
            
        catch ME
            fprintf('✗ 회귀분석 실패: %s\n', ME.message);
        end
    else
        fprintf('✗ 데이터 부족 (최소 5명 필요, 현재 %d명)\n', n);
    end
end

% 시점별 회귀분석
fprintf('\n▶ 시점별 회귀분석\n');
fprintf('----------------------------------------\n');

for p = 1:length(periods)
    diagVar = sprintf('Diagnostic_Period%d', p);
    
    if ismember(diagVar, analysisData.Properties.VariableNames) && ...
       ismember('Test_Average', analysisData.Properties.VariableNames)
        
        X = analysisData.Test_Average;
        Y = analysisData.(diagVar);
        
        validIdx = ~isnan(X) & ~isnan(Y);
        validCount = sum(validIdx);
        
        if validCount >= 5
            X_valid = X(validIdx);
            Y_valid = Y(validIdx);
            
            try
                X_matrix = [ones(validCount, 1), X_valid];
                beta = X_matrix \ Y_valid;
                
                Y_pred = X_matrix * beta;
                SS_tot = sum((Y_valid - mean(Y_valid)).^2);
                SS_res = sum((Y_valid - Y_pred).^2);
                R_squared = 1 - (SS_res / SS_tot);
                
                MSE = SS_res / (validCount - 2);
                SE_slope = sqrt(MSE * inv(X_matrix' * X_matrix));
                SE_slope = SE_slope(2, 2);
                
                t_slope = beta(2) / SE_slope;
                p_slope = 2 * (1 - tcdf(abs(t_slope), validCount-2));
                
                fprintf('%s: β=%.3f (R²=%.3f, p=%.3f, n=%d)', ...
                    periods{p}, beta(2), R_squared, p_slope, validCount);
                if p_slope < 0.05, fprintf(' *'); end
                fprintf('\n');
                
                % 결과 저장
                regressionResults.(sprintf('period%d', p)) = struct(...
                    'slope', beta(2), 'intercept', beta(1), ...
                    'R_squared', R_squared, 'p_slope', p_slope, 'n', validCount);
                
            catch
                fprintf('%s: 회귀분석 실패\n', periods{p});
            end
        else
            fprintf('%s: 데이터 부족 (n=%d)\n', periods{p}, validCount);
        end
    end
end

%% 8. 결과 저장
fprintf('\n[7단계] 결과 저장\n');
fprintf('----------------------------------------\n');

% 결과 파일명 생성
dateStr = datestr(now, 'yyyymmdd_HHMM');
outputFileName = sprintf('역량검사진단_분석결과_%s.xlsx', dateStr);

% 통합 분석 데이터 저장
writetable(analysisData, outputFileName, 'Sheet', '통합_분석데이터');
fprintf('✓ 통합 분석 데이터 저장: %d명\n', height(analysisData));

% 상관분석 결과 저장
corrFields = fieldnames(correlationResults);
if ~isempty(corrFields)
    corrTable = table();
    corrTable.X_Variable = cell(length(corrFields), 1);
    corrTable.Y_Variable = cell(length(corrFields), 1);
    corrTable.X_Name = cell(length(corrFields), 1);
    corrTable.Y_Name = cell(length(corrFields), 1);
    corrTable.Correlation = NaN(length(corrFields), 1);
    corrTable.N = NaN(length(corrFields), 1);
    corrTable.P_Value = NaN(length(corrFields), 1);
    corrTable.Significance = cell(length(corrFields), 1);
    
    for i = 1:length(corrFields)
        field = corrFields{i};
        result = correlationResults.(field);
        
        corrTable.X_Variable{i} = result.x_var;
        corrTable.Y_Variable{i} = result.y_var;
        corrTable.X_Name{i} = result.x_name;
        corrTable.Y_Name{i} = result.y_name;
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
    
    writetable(corrTable, outputFileName, 'Sheet', '상관분석_결과');
    fprintf('✓ 상관분석 결과 저장: %d개 분석\n', length(corrFields));
end

% 회귀분석 결과 저장
regFields = fieldnames(regressionResults);
if ~isempty(regFields)
    regTable = table();
    regTable.Analysis = regFields;
    regTable.Intercept = NaN(length(regFields), 1);
    regTable.Slope = NaN(length(regFields), 1);
    regTable.R_Squared = NaN(length(regFields), 1);
    regTable.P_Slope = NaN(length(regFields), 1);
    regTable.N = NaN(length(regFields), 1);
    regTable.Significance = cell(length(regFields), 1);
    
    for i = 1:length(regFields)
        field = regFields{i};
        result = regressionResults.(field);
        
        regTable.Intercept(i) = result.intercept;
        regTable.Slope(i) = result.slope;
        regTable.R_Squared(i) = result.R_squared;
        regTable.P_Slope(i) = result.p_slope;
        regTable.N(i) = result.n;
        
        if result.p_slope < 0.001
            regTable.Significance{i} = '***';
        elseif result.p_slope < 0.01
            regTable.Significance{i} = '**';
        elseif result.p_slope < 0.05
            regTable.Significance{i} = '*';
        else
            regTable.Significance{i} = 'n.s.';
        end
    end
    
    writetable(regTable, outputFileName, 'Sheet', '회귀분석_결과');
    fprintf('✓ 회귀분석 결과 저장: %d개 분석\n', length(regFields));
end

% 역량검사 원본 데이터도 저장
writetable(competencyTestScores, outputFileName, 'Sheet', '역량검사_원본데이터');
fprintf('✓ 역량검사 원본 데이터 저장: %d명\n', height(competencyTestScores));

% MAT 파일로도 저장
matFileName = sprintf('역량검사진단_분석결과_%s.mat', dateStr);
save(matFileName, 'analysisData', 'competencyTestScores', 'correlationResults', ...
     'regressionResults', 'periods', 'consolidatedScores');
fprintf('✓ MAT 파일 저장: %s\n', matFileName);

%% 9. 최종 요약
fprintf('\n[8단계] 분석 결과 최종 요약\n');
fprintf('========================================\n');

fprintf('📊 데이터 현황\n');
fprintf('  • 역량검사 데이터: %d명 (%d개 역량 점수)\n', height(competencyTestScores), length(competencyScoreCols));
fprintf('  • 역량진단 데이터: %d명\n', height(consolidatedScores));
fprintf('  • 최종 분석 대상: %d명 (매칭률: %.1f%%)\n', height(analysisData), ...
    100 * height(analysisData) / min(height(competencyTestScores), height(consolidatedScores)));

if ismember('Test_Average', analysisData.Properties.VariableNames)
    testScores = analysisData.Test_Average(~isnan(analysisData.Test_Average));
    if ~isempty(testScores)
        fprintf('  • 역량검사 평균: %.3f (±%.3f, 범위: %.1f~%.1f)\n', ...
            mean(testScores), std(testScores), min(testScores), max(testScores));
    end
end

if ismember('Diagnostic_Average', analysisData.Properties.VariableNames)
    diagScores = analysisData.Diagnostic_Average(~isnan(analysisData.Diagnostic_Average));
    if ~isempty(diagScores)
        fprintf('  • 역량진단 평균: %.3f (±%.3f, 범위: %.1f~%.1f)\n', ...
            mean(diagScores), std(diagScores), min(diagScores), max(diagScores));
    end
end

fprintf('\n🔗 주요 상관분석 결과\n');
if exist('corrFields', 'var') && ~isempty(corrFields)
    maxCorr = -1;
    maxCorrResult = [];
    significantCount = 0;
    
    for i = 1:length(corrFields)
        result = correlationResults.(corrFields{i});
        absCorr = abs(result.correlation);
        
        if absCorr > maxCorr
            maxCorr = absCorr;
            maxCorrResult = result;
        end
        
        if result.p_value < 0.05
            significantCount = significantCount + 1;
        end
    end
    
    fprintf('  • 총 분석 쌍: %d개\n', length(corrFields));
    fprintf('  • 유의한 상관: %d개\n', significantCount);
    
    if ~isempty(maxCorrResult)
        sig_str = '';
        if maxCorrResult.p_value < 0.001, sig_str = '***';
        elseif maxCorrResult.p_value < 0.01, sig_str = '**';
        elseif maxCorrResult.p_value < 0.05, sig_str = '*';
        end
        fprintf('  • 최고 상관: %s vs %s (r=%.3f) %s\n', ...
            maxCorrResult.x_name, maxCorrResult.y_name, maxCorrResult.correlation, sig_str);
    end
else
    fprintf('  • 상관분석 결과 없음\n');
end

fprintf('\n📈 주요 회귀분석 결과\n');
if isfield(regressionResults, 'main')
    mainResult = regressionResults.main;
    fprintf('  • 회귀식: 역량진단 = %.3f + %.3f × 역량검사\n', mainResult.intercept, mainResult.slope);
    fprintf('  • 설명력: R² = %.3f (%.1f%%)\n', mainResult.R_squared, mainResult.R_squared*100);
    fprintf('  • 통계적 유의성: p = %.3f', mainResult.p_slope);
    if mainResult.p_slope < 0.001, fprintf(' ***');
    elseif mainResult.p_slope < 0.01, fprintf(' **');
    elseif mainResult.p_slope < 0.05, fprintf(' *');
    end
    fprintf('\n');
    
    % 해석
    if mainResult.p_slope < 0.05
        if mainResult.slope > 0
            fprintf('  • 해석: 역량검사 1점 증가 → 역량진단 %.3f점 증가\n', mainResult.slope);
        else
            fprintf('  • 해석: 역량검사 1점 증가 → 역량진단 %.3f점 감소\n', abs(mainResult.slope));
        end
        
        % 효과크기 평가
        if mainResult.R_squared >= 0.25
            fprintf('  • 효과크기: 큰 효과 (R² ≥ 0.25)\n');
        elseif mainResult.R_squared >= 0.09
            fprintf('  • 효과크기: 중간 효과 (R² ≥ 0.09)\n');
        elseif mainResult.R_squared >= 0.01
            fprintf('  • 효과크기: 작은 효과 (R² ≥ 0.01)\n');
        else
            fprintf('  • 효과크기: 매우 작은 효과\n');
        end
    else
        fprintf('  • 해석: 통계적으로 유의한 예측 관계 없음\n');
    end
else
    fprintf('  • 주요 회귀분석 결과 없음\n');
end

fprintf('\n🎯 시점별 회귀분석 요약\n');
for p = 1:length(periods)
    fieldName = sprintf('period%d', p);
    if isfield(regressionResults, fieldName)
        result = regressionResults.(fieldName);
        fprintf('  • %s: β=%.3f (R²=%.3f, p=%.3f)', periods{p}, result.slope, result.R_squared, result.p_slope);
        if result.p_slope < 0.05, fprintf(' *'); end
        fprintf('\n');
    end
end

fprintf('\n📋 실용적 결론\n');
fprintf('----------------------------------------\n');
if isfield(regressionResults, 'main') && regressionResults.main.p_slope < 0.05
    mainResult = regressionResults.main;
    if mainResult.slope > 0
        fprintf('✓ 역량검사와 역량진단 간에 양의 선형관계가 확인되었습니다.\n');
        fprintf('✓ 역량검사를 통해 역량진단 성과를 어느 정도 예측할 수 있습니다.\n');
        fprintf('✓ 역량 개발 프로그램의 효과를 사전에 예측하는 도구로 활용 가능합니다.\n');
    else
        fprintf('⚠ 역량검사와 역량진단 간에 음의 관계가 발견되었습니다.\n');
        fprintf('⚠ 측정 방식이나 평가 기준의 차이를 검토해볼 필요가 있습니다.\n');
    end
else
    fprintf('× 역량검사와 역량진단 간에 통계적으로 유의한 관계가 발견되지 않았습니다.\n');
    fprintf('× 두 측정 도구가 서로 다른 역량 측면을 측정할 가능성이 있습니다.\n');
    fprintf('× 추가적인 변수나 조절효과를 고려한 분석이 필요할 수 있습니다.\n');
end

fprintf('\n📁 저장된 파일\n');
fprintf('  • Excel: %s\n', outputFileName);
fprintf('  • MAT: %s\n', matFileName);

fprintf('\n✅ 역량검사-역량진단 분석 완료!\n');
fprintf('   간결하고 독립적인 분석으로 핵심 관계를 명확히 파악했습니다.\n');

fprintf('\n%s\n', repmat('=', 1, 60));
fprintf('분석 완료 시각: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('%s\n', repmat('=', 1, 60));


%% ===============================================================
%  단순 상관분석 기반 하위항목 분석
%  1단계: 개별 상관분석으로 유의미한 특성 식별
%  2단계: 선별된 특성들로만 예측 모델 구축
%% ===============================================================

%% 10. 단계별 하위항목 분석


clc
fprintf('\n[9단계] 단계별 하위항목 분석\n');
fprintf('========================================\n');
fprintf('접근법: 단순 상관분석 → 특성 선별 → 예측 모델\n\n');

%% 10-1. 하위항목 데이터 로드 (기존과 동일)
fprintf('▶ 하위항목 데이터 로드\n');

subitemPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
try
    subitemData = readtable(subitemPath, 'Sheet', '역량검사_하위항목', 'VariableNamingRule', 'preserve');
    fprintf('✓ 하위항목 데이터 로드 완료: %d명, %d컬럼\n', height(subitemData), width(subitemData));
catch ME
    fprintf('✗ 하위항목 데이터 로드 실패: %s\n', ME.message);
    return;
end

% ID 및 점수 컬럼 식별
subitemIdCol = findIDColumn(subitemData);
subitemIDs = extractAndStandardizeIDs(subitemData{:, subitemIdCol});

colNames = subitemData.Properties.VariableNames;
subitemScoreCols = {};

for col = 1:width(subitemData)
    colName = colNames{col};
    colData = subitemData{:, col};
    
    if col ~= subitemIdCol && isnumeric(colData) && sum(~isnan(colData)) > 10
        subitemScoreCols{end+1} = colName;
    end
end

fprintf('✓ 발견된 하위항목: %d개\n', length(subitemScoreCols));

%% 10-2. 성과 데이터와 매칭 (기존과 동일)
fprintf('\n▶ 성과 데이터와 매칭\n');

if exist('consolidatedScores', 'var') && istable(consolidatedScores)
    performanceData = consolidatedScores;
else
    fprintf('✗ 성과 데이터가 없습니다.\n');
    return;
end

% ID 매칭
if isnumeric(performanceData.ID)
    performanceIDs = arrayfun(@num2str, performanceData.ID, 'UniformOutput', false);
else
    performanceIDs = cellfun(@char, performanceData.ID, 'UniformOutput', false);
end

if isnumeric(subitemIDs)
    subitemIDs_str = arrayfun(@num2str, subitemIDs, 'UniformOutput', false);
else
    subitemIDs_str = cellfun(@char, subitemIDs, 'UniformOutput', false);
end

[commonIDs_sub, perfIdx, subIdx] = intersect(performanceIDs, subitemIDs_str);
fprintf('매칭 결과: %d명\n', length(commonIDs_sub));

if length(commonIDs_sub) < 15
    fprintf('✗ 매칭된 데이터가 너무 적습니다.\n');
    return;
end

%% 10-3. 분석용 데이터 구성
fprintf('\n▶ 분석용 데이터 구성\n');

% 특성 행렬 구성
X_all = [];
featureNames_all = {};

for i = 1:length(subitemScoreCols)
    featureNames_all{end+1} = subitemScoreCols{i};
    X_all = [X_all, subitemData.(subitemScoreCols{i})(subIdx)];
end

% 타겟 변수 설정
if ismember('AverageStdScore', performanceData.Properties.VariableNames)
    Y_all = performanceData.AverageStdScore(perfIdx);
    targetName = '역량진단 평균점수';
else
    % 대안 타겟 찾기
    scoreCols = performanceData.Properties.VariableNames;
    scoreCol = '';
    for col = 1:width(performanceData)
        if contains(scoreCols{col}, 'Score') || contains(scoreCols{col}, '점수')
            scoreCol = scoreCols{col};
            break;
        end
    end
    
    if ~isempty(scoreCol)
        Y_all = performanceData.(scoreCol)(perfIdx);
        targetName = scoreCol;
    else
        fprintf('✗ 적절한 성과 타겟 변수를 찾을 수 없습니다.\n');
        return;
    end
end

% 결측치 제거
validIdx = ~isnan(Y_all);
for i = 1:size(X_all, 2)
    validIdx = validIdx & ~isnan(X_all(:, i));
end

X_clean = X_all(validIdx, :);
Y_clean = Y_all(validIdx);
commonIDs_clean = commonIDs_sub(validIdx);

fprintf('✓ 최종 분석 데이터: %d명, %d개 특성\n', length(Y_clean), size(X_clean, 2));
fprintf('✓ 타겟 변수: %s\n', targetName);

%% 10-4. 1단계: 개별 상관분석
fprintf('\n▶ 1단계: 개별 상관분석 수행\n');

n_features = size(X_clean, 2);
correlation_results = table();
correlation_results.Feature = featureNames_all';
correlation_results.Correlation = NaN(n_features, 1);
correlation_results.P_Value = NaN(n_features, 1);
correlation_results.N_Valid = NaN(n_features, 1);

fprintf('각 하위항목과 %s 간의 상관분석:\n', targetName);
fprintf('%-25s %10s %10s %10s %10s\n', '특성명', '상관계수', 'p-value', '유효N', '유의성');
fprintf('%s\n', repmat('-', 1, 75));

significant_features = {};
significant_correlations = [];
significant_p_values = [];

for i = 1:n_features
    feature_data = X_clean(:, i);
    
    % 상관분석
    [r, p] = corr(feature_data, Y_clean);
    
    correlation_results.Correlation(i) = r;
    correlation_results.P_Value(i) = p;
    correlation_results.N_Valid(i) = length(Y_clean);
    
    % 유의성 표시
    sig_str = '';
    if p < 0.001
        sig_str = '***';
    elseif p < 0.01
        sig_str = '**';
    elseif p < 0.05
        sig_str = '*';
    end
    
    % 출력
    feature_name = featureNames_all{i};
    if length(feature_name) > 25
        feature_name = [feature_name(1:22), '...'];
    end
    
    fprintf('%-25s %10.3f %10.3f %10d %10s\n', ...
        feature_name, r, p, length(Y_clean), sig_str);
    
    % 유의미한 특성 수집 (p < 0.05 또는 |r| > 0.2)
    if p < 0.05 || abs(r) > 0.2
        significant_features{end+1} = featureNames_all{i};
        significant_correlations(end+1) = r;
        significant_p_values(end+1) = p;
    end
end

% 상관계수 순으로 정렬
[~, sort_idx] = sort(abs(correlation_results.Correlation), 'descend');
correlation_results = correlation_results(sort_idx, :);

fprintf('\n상관분석 요약:\n');
fprintf('  - 전체 특성: %d개\n', n_features);
fprintf('  - 유의미한 특성 (p<0.05 또는 |r|>0.2): %d개\n', length(significant_features));

% 유의미한 상관계수가 있는지 확인
strong_corr_count = sum(abs(correlation_results.Correlation) > 0.3);
moderate_corr_count = sum(abs(correlation_results.Correlation) > 0.2);
weak_corr_count = sum(abs(correlation_results.Correlation) > 0.1);

fprintf('  - 강한 상관 (|r|>0.3): %d개\n', strong_corr_count);
fprintf('  - 보통 상관 (|r|>0.2): %d개\n', moderate_corr_count);
fprintf('  - 약한 상관 (|r|>0.1): %d개\n', weak_corr_count);

%% 10-5. 상위 상관계수 특성들 상세 분석
fprintf('\n▶ 상위 상관계수 특성들 상세 분석\n');

top_n = min(15, height(correlation_results));
fprintf('상위 %d개 특성 상세 정보:\n', top_n);
fprintf('%-25s %10s %10s %10s %15s\n', '특성명', '상관계수', 'p-value', '효과크기', '해석');
fprintf('%s\n', repmat('-', 1, 85));

for i = 1:top_n
    feature = correlation_results.Feature{i};
    r = correlation_results.Correlation(i);
    p = correlation_results.P_Value(i);
    
    % 효과크기 해석
    if abs(r) >= 0.5
        effect_size = '큰 효과';
    elseif abs(r) >= 0.3
        effect_size = '중간 효과';
    elseif abs(r) >= 0.1
        effect_size = '작은 효과';
    else
        effect_size = '미미한 효과';
    end
    
    % 실용적 해석
    if p < 0.05 && abs(r) > 0.2
        interpretation = '유의미';
    elseif p < 0.05
        interpretation = '통계적 유의';
    elseif abs(r) > 0.2
        interpretation = '실용적 의미';
    else
        interpretation = '의미 제한적';
    end
    
    % 특성명 길이 조정
    if length(feature) > 25
        feature = [feature(1:22), '...'];
    end
    
    fprintf('%-25s %10.3f %10.3f %10s %15s\n', ...
        feature, r, p, effect_size, interpretation);
end

%% 10-6. 2단계: 유의미한 특성들로 예측 모델 구축
fprintf('\n▶ 2단계: 선별된 특성들로 예측 모델 구축\n');

% 유의미한 특성 선별 기준 설정
selection_criteria = struct();
selection_criteria.min_abs_corr = 0.15;  % 최소 절대 상관계수
selection_criteria.max_p_value = 0.10;   % 최대 p-value
selection_criteria.max_features = 10;    % 최대 특성 수

% 기준에 따른 특성 선별
selected_mask = (abs(correlation_results.Correlation) >= selection_criteria.min_abs_corr) & ...
                (correlation_results.P_Value <= selection_criteria.max_p_value);

selected_features_table = correlation_results(selected_mask, :);

% 상위 N개로 제한
if height(selected_features_table) > selection_criteria.max_features
    selected_features_table = selected_features_table(1:selection_criteria.max_features, :);
end

if height(selected_features_table) == 0
    fprintf('선별 기준을 만족하는 특성이 없습니다.\n');
    fprintf('기준을 완화하여 상위 5개 특성을 선택합니다.\n');
    selected_features_table = correlation_results(1:min(5, height(correlation_results)), :);
end

selected_feature_names = selected_features_table.Feature;
n_selected = length(selected_feature_names);

fprintf('선별된 특성 (%d개):\n', n_selected);
for i = 1:n_selected
    fprintf('  %d. %s (r=%.3f, p=%.3f)\n', i, selected_feature_names{i}, ...
        selected_features_table.Correlation(i), selected_features_table.P_Value(i));
end

% 선별된 특성들의 데이터 추출
selected_indices = zeros(n_selected, 1);
for i = 1:n_selected
    selected_indices(i) = find(strcmp(featureNames_all, selected_feature_names{i}));
end

X_selected = X_clean(:, selected_indices);

%% 10-7. 선별된 특성들로 다양한 모델 테스트
fprintf('\n▶ 선별된 특성들로 다양한 모델 테스트\n');

% 데이터 표준화
X_selected_std = zeros(size(X_selected));
for i = 1:size(X_selected, 2)
    X_selected_std(:, i) = (X_selected(:, i) - mean(X_selected(:, i))) / (std(X_selected(:, i)) + eps);
end

% 훈련/테스트 분할
test_ratio = 0.25;
if length(Y_clean) < 20
    test_ratio = 0.2;  % 데이터가 적으면 테스트 비율 줄임
end

cv_partition = cvpartition(length(Y_clean), 'HoldOut', test_ratio);
X_train = X_selected_std(training(cv_partition), :);
X_test = X_selected_std(test(cv_partition), :);
Y_train = Y_clean(training(cv_partition));
Y_test = Y_clean(test(cv_partition));

fprintf('훈련: %d명, 테스트: %d명, 특성: %d개\n', length(Y_train), length(Y_test), size(X_train, 2));

% 모델 성능 저장
model_results = struct();

% 1. 다중선형회귀 (OLS)
fprintf('\n1. 다중선형회귀 (OLS):\n');
try
    if size(X_train, 2) < size(X_train, 1)  % 특성 수 < 샘플 수
        beta_ols = [ones(size(X_train, 1), 1), X_train] \ Y_train;
        Y_pred_ols = [ones(size(X_test, 1), 1), X_test] * beta_ols;
        
        r2_ols = 1 - sum((Y_test - Y_pred_ols).^2) / sum((Y_test - mean(Y_test)).^2);
        rmse_ols = sqrt(mean((Y_test - Y_pred_ols).^2));
        corr_ols = corr(Y_test, Y_pred_ols);
        
        model_results.ols.r2 = r2_ols;
        model_results.ols.rmse = rmse_ols;
        model_results.ols.corr = corr_ols;
        model_results.ols.coeffs = beta_ols(2:end);
        
        fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f\n', r2_ols, rmse_ols, corr_ols);
    else
        fprintf('   건너뜀 (특성 수 > 샘플 수)\n');
        model_results.ols.r2 = NaN;
    end
catch
    fprintf('   실패\n');
    model_results.ols.r2 = NaN;
end

% 2. Ridge 회귀
fprintf('\n2. Ridge 회귀:\n');
try
    [B_ridge, FitInfo_ridge] = lasso(X_train, Y_train, 'Alpha', 0, 'CV', 3);
    Y_pred_ridge = X_test * B_ridge(:, FitInfo_ridge.IndexMinMSE) + FitInfo_ridge.Intercept(FitInfo_ridge.IndexMinMSE);
    
    r2_ridge = 1 - sum((Y_test - Y_pred_ridge).^2) / sum((Y_test - mean(Y_test)).^2);
    rmse_ridge = sqrt(mean((Y_test - Y_pred_ridge).^2));
    corr_ridge = corr(Y_test, Y_pred_ridge);
    
    model_results.ridge.r2 = r2_ridge;
    model_results.ridge.rmse = rmse_ridge;
    model_results.ridge.corr = corr_ridge;
    model_results.ridge.coeffs = B_ridge(:, FitInfo_ridge.IndexMinMSE);
    model_results.ridge.lambda = FitInfo_ridge.Lambda(FitInfo_ridge.IndexMinMSE);
    
    fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f, λ = %.4f\n', ...
        r2_ridge, rmse_ridge, corr_ridge, model_results.ridge.lambda);
catch
    fprintf('   실패\n');
    model_results.ridge.r2 = NaN;
end

% 3. Elastic Net (관대한 설정)
fprintf('\n3. Elastic Net (관대한 설정):\n');
try
    [B_elastic, FitInfo_elastic] = lasso(X_train, Y_train, 'Alpha', 0.2, 'CV', 3, 'LambdaRatio', 1e-4);
    Y_pred_elastic = X_test * B_elastic(:, FitInfo_elastic.IndexMinMSE) + FitInfo_elastic.Intercept(FitInfo_elastic.IndexMinMSE);
    
    r2_elastic = 1 - sum((Y_test - Y_pred_elastic).^2) / sum((Y_test - mean(Y_test)).^2);
    rmse_elastic = sqrt(mean((Y_test - Y_pred_elastic).^2));
    corr_elastic = corr(Y_test, Y_pred_elastic);
    
    model_results.elastic.r2 = r2_elastic;
    model_results.elastic.rmse = rmse_elastic;
    model_results.elastic.corr = corr_elastic;
    model_results.elastic.coeffs = B_elastic(:, FitInfo_elastic.IndexMinMSE);
    model_results.elastic.lambda = FitInfo_elastic.Lambda(FitInfo_elastic.IndexMinMSE);
    
    fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f, λ = %.4f\n', ...
        r2_elastic, rmse_elastic, corr_elastic, model_results.elastic.lambda);
catch
    fprintf('   실패\n');
    model_results.elastic.r2 = NaN;
end

% 4. 단순 평균 기반 모델
fprintf('\n4. 단순 평균 기반:\n');
% 상관계수를 가중치로 사용한 선형결합
weights = selected_features_table.Correlation / sum(abs(selected_features_table.Correlation));
Y_pred_weighted = X_test * weights;

% 스케일 조정
Y_pred_weighted = Y_pred_weighted * std(Y_train) / std(Y_pred_weighted) + mean(Y_train);

r2_weighted = 1 - sum((Y_test - Y_pred_weighted).^2) / sum((Y_test - mean(Y_test)).^2);
rmse_weighted = sqrt(mean((Y_test - Y_pred_weighted).^2));
corr_weighted = corr(Y_test, Y_pred_weighted);

model_results.weighted.r2 = r2_weighted;
model_results.weighted.rmse = rmse_weighted;
model_results.weighted.corr = corr_weighted;
model_results.weighted.coeffs = weights;

fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f\n', r2_weighted, rmse_weighted, corr_weighted);
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

%% 7. 단순회귀분석
fprintf('\n[6단계] 단순회귀분석: 역량검사 → 역량진단\n');
fprintf('========================================\n');

regressionResults = struct();

% 주요 회귀분석: 역량검사 평균 → 역량진단 평균
if ismember('Test_Average', analysisData.Properties.VariableNames) && ...
   ismember('Diagnostic_Average', analysisData.Properties.VariableNames)
    
    fprintf('▶ 주요 분석: 역량검사 평균 → 역량진단 평균\n');
    
    X = analysisData.Test_Average;
    Y = analysisData.Diagnostic_Average;
    
    % 결측치 제거
    validIdx = ~isnan(X) & ~isnan(Y);
    X_valid = X(validIdx);
    Y_valid = Y(validIdx);
    n = length(X_valid);
    
    fprintf('유효한 데이터: %d명\n', n);
    
    if n >= 5
        try
            % 회귀분석 실행
            X_matrix = [ones(n, 1), X_valid];
            beta = X_matrix \ Y_valid;
            
            intercept = beta(1);
            slope = beta(2);
            
            % 예측값 및 잔차
            Y_pred = X_matrix * beta;
            residuals = Y_valid - Y_pred;
            
            % 결정계수 (R²)
            SS_tot = sum((Y_valid - mean(Y_valid)).^2);
            SS_res = sum(residuals.^2);
            R_squared = 1 - (SS_res / SS_tot);
            R_squared_adj = 1 - ((SS_res/(n-2)) / (SS_tot/(n-1)));
            
            % 표준오차
            MSE = SS_res / (n - 2);
            SE_matrix = sqrt(MSE * inv(X_matrix' * X_matrix));
            SE_intercept = SE_matrix(1, 1);
            SE_slope = SE_matrix(2, 2);
            
            % t-검정
            t_slope = slope / SE_slope;
            t_intercept = intercept / SE_intercept;
            p_slope = 2 * (1 - tcdf(abs(t_slope), n-2));
            p_intercept = 2 * (1 - tcdf(abs(t_intercept), n-2));
            
            % F-검정
            F_stat = (SS_tot - SS_res) / (SS_res / (n-2));
            p_F = 1 - fcdf(F_stat, 1, n-2);
            
            % 결과 출력
            fprintf('\n회귀분석 결과:\n');
            fprintf('══════════════════════════════════════\n');
            fprintf('회귀식: 역량진단 = %.3f + %.3f × 역량검사\n', intercept, slope);
            fprintf('\n계수 추정치:\n');
            fprintf('  절편: %.3f (SE=%.3f, t=%.3f, p=%.3f)', intercept, SE_intercept, t_intercept, p_intercept);
            if p_intercept < 0.05, fprintf(' *'); end
            fprintf('\n');
            fprintf('  기울기: %.3f (SE=%.3f, t=%.3f, p=%.3f)', slope, SE_slope, t_slope, p_slope);
            if p_slope < 0.001, fprintf(' ***');
            elseif p_slope < 0.01, fprintf(' **');
            elseif p_slope < 0.05, fprintf(' *'); end
            fprintf('\n\n');
            
            fprintf('모형 적합도:\n');
            fprintf('  R² = %.3f (설명된 분산: %.1f%%)\n', R_squared, R_squared*100);
            fprintf('  조정된 R² = %.3f\n', R_squared_adj);
            fprintf('  F통계량 = %.3f (p=%.3f)', F_stat, p_F);
            if p_F < 0.001, fprintf(' ***');
            elseif p_F < 0.01, fprintf(' **');
            elseif p_F < 0.05, fprintf(' *'); end
            fprintf('\n');
            fprintf('  RMSE = %.3f\n\n', sqrt(MSE));
            
            % 실용적 해석
            fprintf('실용적 해석:\n');
            fprintf('══════════════════════════════════════\n');
            if p_slope < 0.05
                if slope > 0
                    fprintf('✓ 역량검사 점수가 1점 증가하면, 역량진단 점수가 %.3f점 증가합니다.\n', slope);
                    fprintf('✓ 역량검사는 역량진단 성과의 %.1f%%를 설명합니다.\n', R_squared*100);
                    
                    if ~isempty(X_valid)
                        example_increase = std(X_valid);
                        predicted_change = slope * example_increase;
                        fprintf('✓ 예시: 역량검사가 %.1f점(1SD) 증가 → 역량진단 %.3f점 증가 예상\n', ...
                            example_increase, predicted_change);
                    end
                else
                    fprintf('✓ 역량검사 점수가 1점 증가하면, 역량진단 점수가 %.3f점 감소합니다.\n', abs(slope));
                end
            else
                fprintf('✗ 역량검사와 역량진단 간에 통계적으로 유의한 예측 관계가 없습니다.\n');
                fprintf('  (p = %.3f > 0.05)\n', p_slope);
            end
            
            % 결과 저장
            regressionResults.main = struct(...
                'intercept', intercept, 'slope', slope, ...
                'SE_intercept', SE_intercept, 'SE_slope', SE_slope, ...
                't_intercept', t_intercept, 't_slope', t_slope, ...
                'p_intercept', p_intercept, 'p_slope', p_slope, ...
                'R_squared', R_squared, 'R_squared_adj', R_squared_adj, ...
                'F_stat', F_stat, 'p_F', p_F, 'RMSE', sqrt(MSE), 'n', n);
            
        catch ME
            fprintf('✗ 회귀분석 실패: %s\n', ME.message);
        end
    else
        fprintf('✗ 데이터 부족 (최소 5명 필요, 현재 %d명)\n', n);
    end
end

% 시점별 회귀분석
fprintf('\n▶ 시점별 회귀분석\n');
fprintf('----------------------------------------\n');

for p = 1:length(periods)
    diagVar = sprintf('Diagnostic_Period%d', p);
    
    if ismember(diagVar, analysisData.Properties.VariableNames) && ...
       ismember('Test_Average', analysisData.Properties.VariableNames)
        
        X = analysisData.Test_Average;
        Y = analysisData.(diagVar);
        
        validIdx = ~isnan(X) & ~isnan(Y);
        validCount = sum(validIdx);
        
        if validCount >= 5
            X_valid = X(validIdx);
            Y_valid = Y(validIdx);
            
            try
                X_matrix = [ones(validCount, 1), X_valid];
                beta = X_matrix \ Y_valid;
                
                Y_pred = X_matrix * beta;
                SS_tot = sum((Y_valid - mean(Y_valid)).^2);
                SS_res = sum((Y_valid - Y_pred).^2);
                R_squared = 1 - (SS_res / SS_tot);
                
                MSE = SS_res / (validCount - 2);
                SE_slope = sqrt(MSE * inv(X_matrix' * X_matrix));
                SE_slope = SE_slope(2, 2);
                
                t_slope = beta(2) / SE_slope;
                p_slope = 2 * (1 - tcdf(abs(t_slope), validCount-2));
                
                fprintf('%s: β=%.3f (R²=%.3f, p=%.3f, n=%d)', ...
                    periods{p}, beta(2), R_squared, p_slope, validCount);
                if p_slope < 0.05, fprintf(' *'); end
                fprintf('\n');
                
                % 결과 저장
                regressionResults.(sprintf('period%d', p)) = struct(...
                    'slope', beta(2), 'intercept', beta(1), ...
                    'R_squared', R_squared, 'p_slope', p_slope, 'n', validCount);
                
            catch
                fprintf('%s: 회귀분석 실패\n', periods{p});
            end
        else
            fprintf('%s: 데이터 부족 (n=%d)\n', periods{p}, validCount);
        end
    end
end

%% 8. 결과 저장
fprintf('\n[7단계] 결과 저장\n');
fprintf('----------------------------------------\n');

% 결과 파일명 생성
dateStr = datestr(now, 'yyyymmdd_HHMM');
outputFileName = sprintf('역량검사진단_분석결과_%s.xlsx', dateStr);

% 통합 분석 데이터 저장
writetable(analysisData, outputFileName, 'Sheet', '통합_분석데이터');
fprintf('✓ 통합 분석 데이터 저장: %d명\n', height(analysisData));

% 상관분석 결과 저장
corrFields = fieldnames(correlationResults);
if ~isempty(corrFields)
    corrTable = table();
    corrTable.X_Variable = cell(length(corrFields), 1);
    corrTable.Y_Variable = cell(length(corrFields), 1);
    corrTable.X_Name = cell(length(corrFields), 1);
    corrTable.Y_Name = cell(length(corrFields), 1);
    corrTable.Correlation = NaN(length(corrFields), 1);
    corrTable.N = NaN(length(corrFields), 1);
    corrTable.P_Value = NaN(length(corrFields), 1);
    corrTable.Significance = cell(length(corrFields), 1);
    
    for i = 1:length(corrFields)
        field = corrFields{i};
        result = correlationResults.(field);
        
        corrTable.X_Variable{i} = result.x_var;
        corrTable.Y_Variable{i} = result.y_var;
        corrTable.X_Name{i} = result.x_name;
        corrTable.Y_Name{i} = result.y_name;
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
    
    writetable(corrTable, outputFileName, 'Sheet', '상관분석_결과');
    fprintf('✓ 상관분석 결과 저장: %d개 분석\n', length(corrFields));
end

% 회귀분석 결과 저장
regFields = fieldnames(regressionResults);
if ~isempty(regFields)
    regTable = table();
    regTable.Analysis = regFields;
    regTable.Intercept = NaN(length(regFields), 1);
    regTable.Slope = NaN(length(regFields), 1);
    regTable.R_Squared = NaN(length(regFields), 1);
    regTable.P_Slope = NaN(length(regFields), 1);
    regTable.N = NaN(length(regFields), 1);
    regTable.Significance = cell(length(regFields), 1);
    
    for i = 1:length(regFields)
        field = regFields{i};
        result = regressionResults.(field);
        
        regTable.Intercept(i) = result.intercept;
        regTable.Slope(i) = result.slope;
        regTable.R_Squared(i) = result.R_squared;
        regTable.P_Slope(i) = result.p_slope;
        regTable.N(i) = result.n;
        
        if result.p_slope < 0.001
            regTable.Significance{i} = '***';
        elseif result.p_slope < 0.01
            regTable.Significance{i} = '**';
        elseif result.p_slope < 0.05
            regTable.Significance{i} = '*';
        else
            regTable.Significance{i} = 'n.s.';
        end
    end
    
    writetable(regTable, outputFileName, 'Sheet', '회귀분석_결과');
    fprintf('✓ 회귀분석 결과 저장: %d개 분석\n', length(regFields));
end

% 역량검사 원본 데이터도 저장
writetable(competencyTestScores, outputFileName, 'Sheet', '역량검사_원본데이터');
fprintf('✓ 역량검사 원본 데이터 저장: %d명\n', height(competencyTestScores));

% MAT 파일로도 저장
matFileName = sprintf('역량검사진단_분석결과_%s.mat', dateStr);
save(matFileName, 'analysisData', 'competencyTestScores', 'correlationResults', ...
     'regressionResults', 'periods', 'consolidatedScores');
fprintf('✓ MAT 파일 저장: %s\n', matFileName);

%% 9. 최종 요약
fprintf('\n[8단계] 분석 결과 최종 요약\n');
fprintf('========================================\n');

fprintf('📊 데이터 현황\n');
fprintf('  • 역량검사 데이터: %d명 (%d개 역량 점수)\n', height(competencyTestScores), length(competencyScoreCols));
fprintf('  • 역량진단 데이터: %d명\n', height(consolidatedScores));
fprintf('  • 최종 분석 대상: %d명 (매칭률: %.1f%%)\n', height(analysisData), ...
    100 * height(analysisData) / min(height(competencyTestScores), height(consolidatedScores)));

if ismember('Test_Average', analysisData.Properties.VariableNames)
    testScores = analysisData.Test_Average(~isnan(analysisData.Test_Average));
    if ~isempty(testScores)
        fprintf('  • 역량검사 평균: %.3f (±%.3f, 범위: %.1f~%.1f)\n', ...
            mean(testScores), std(testScores), min(testScores), max(testScores));
    end
end

if ismember('Diagnostic_Average', analysisData.Properties.VariableNames)
    diagScores = analysisData.Diagnostic_Average(~isnan(analysisData.Diagnostic_Average));
    if ~isempty(diagScores)
        fprintf('  • 역량진단 평균: %.3f (±%.3f, 범위: %.1f~%.1f)\n', ...
            mean(diagScores), std(diagScores), min(diagScores), max(diagScores));
    end
end

fprintf('\n🔗 주요 상관분석 결과\n');
if exist('corrFields', 'var') && ~isempty(corrFields)
    maxCorr = -1;
    maxCorrResult = [];
    significantCount = 0;
    
    for i = 1:length(corrFields)
        result = correlationResults.(corrFields{i});
        absCorr = abs(result.correlation);
        
        if absCorr > maxCorr
            maxCorr = absCorr;
            maxCorrResult = result;
        end
        
        if result.p_value < 0.05
            significantCount = significantCount + 1;
        end
    end
    
    fprintf('  • 총 분석 쌍: %d개\n', length(corrFields));
    fprintf('  • 유의한 상관: %d개\n', significantCount);
    
    if ~isempty(maxCorrResult)
        sig_str = '';
        if maxCorrResult.p_value < 0.001, sig_str = '***';
        elseif maxCorrResult.p_value < 0.01, sig_str = '**';
        elseif maxCorrResult.p_value < 0.05, sig_str = '*';
        end
        fprintf('  • 최고 상관: %s vs %s (r=%.3f) %s\n', ...
            maxCorrResult.x_name, maxCorrResult.y_name, maxCorrResult.correlation, sig_str);
    end
else
    fprintf('  • 상관분석 결과 없음\n');
end

fprintf('\n📈 주요 회귀분석 결과\n');
if isfield(regressionResults, 'main')
    mainResult = regressionResults.main;
    fprintf('  • 회귀식: 역량진단 = %.3f + %.3f × 역량검사\n', mainResult.intercept, mainResult.slope);
    fprintf('  • 설명력: R² = %.3f (%.1f%%)\n', mainResult.R_squared, mainResult.R_squared*100);
    fprintf('  • 통계적 유의성: p = %.3f', mainResult.p_slope);
    if mainResult.p_slope < 0.001, fprintf(' ***');
    elseif mainResult.p_slope < 0.01, fprintf(' **');
    elseif mainResult.p_slope < 0.05, fprintf(' *');
    end
    fprintf('\n');
    
    % 해석
    if mainResult.p_slope < 0.05
        if mainResult.slope > 0
            fprintf('  • 해석: 역량검사 1점 증가 → 역량진단 %.3f점 증가\n', mainResult.slope);
        else
            fprintf('  • 해석: 역량검사 1점 증가 → 역량진단 %.3f점 감소\n', abs(mainResult.slope));
        end
        
        % 효과크기 평가
        if mainResult.R_squared >= 0.25
            fprintf('  • 효과크기: 큰 효과 (R² ≥ 0.25)\n');
        elseif mainResult.R_squared >= 0.09
            fprintf('  • 효과크기: 중간 효과 (R² ≥ 0.09)\n');
        elseif mainResult.R_squared >= 0.01
            fprintf('  • 효과크기: 작은 효과 (R² ≥ 0.01)\n');
        else
            fprintf('  • 효과크기: 매우 작은 효과\n');
        end
    else
        fprintf('  • 해석: 통계적으로 유의한 예측 관계 없음\n');
    end
else
    fprintf('  • 주요 회귀분석 결과 없음\n');
end

fprintf('\n🎯 시점별 회귀분석 요약\n');
for p = 1:length(periods)
    fieldName = sprintf('period%d', p);
    if isfield(regressionResults, fieldName)
        result = regressionResults.(fieldName);
        fprintf('  • %s: β=%.3f (R²=%.3f, p=%.3f)', periods{p}, result.slope, result.R_squared, result.p_slope);
        if result.p_slope < 0.05, fprintf(' *'); end
        fprintf('\n');
    end
end

fprintf('\n📋 실용적 결론\n');
fprintf('----------------------------------------\n');
if isfield(regressionResults, 'main') && regressionResults.main.p_slope < 0.05
    mainResult = regressionResults.main;
    if mainResult.slope > 0
        fprintf('✓ 역량검사와 역량진단 간에 양의 선형관계가 확인되었습니다.\n');
        fprintf('✓ 역량검사를 통해 역량진단 성과를 어느 정도 예측할 수 있습니다.\n');
        fprintf('✓ 역량 개발 프로그램의 효과를 사전에 예측하는 도구로 활용 가능합니다.\n');
    else
        fprintf('⚠ 역량검사와 역량진단 간에 음의 관계가 발견되었습니다.\n');
        fprintf('⚠ 측정 방식이나 평가 기준의 차이를 검토해볼 필요가 있습니다.\n');
    end
else
    fprintf('× 역량검사와 역량진단 간에 통계적으로 유의한 관계가 발견되지 않았습니다.\n');
    fprintf('× 두 측정 도구가 서로 다른 역량 측면을 측정할 가능성이 있습니다.\n');
    fprintf('× 추가적인 변수나 조절효과를 고려한 분석이 필요할 수 있습니다.\n');
end

fprintf('\n📁 저장된 파일\n');
fprintf('  • Excel: %s\n', outputFileName);
fprintf('  • MAT: %s\n', matFileName);

fprintf('\n✅ 역량검사-역량진단 분석 완료!\n');
fprintf('   간결하고 독립적인 분석으로 핵심 관계를 명확히 파악했습니다.\n');

fprintf('\n%s\n', repmat('=', 1, 60));
fprintf('분석 완료 시각: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('%s\n', repmat('=', 1, 60));


%% ===============================================================
%  단순 상관분석 기반 하위항목 분석
%  1단계: 개별 상관분석으로 유의미한 특성 식별
%  2단계: 선별된 특성들로만 예측 모델 구축
%% ===============================================================

%% 10. 단계별 하위항목 분석


clc
fprintf('\n[9단계] 단계별 하위항목 분석\n');
fprintf('========================================\n');
fprintf('접근법: 단순 상관분석 → 특성 선별 → 예측 모델\n\n');

%% 10-1. 하위항목 데이터 로드 (기존과 동일)
fprintf('▶ 하위항목 데이터 로드\n');

subitemPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
try
    subitemData = readtable(subitemPath, 'Sheet', '역량검사_하위항목', 'VariableNamingRule', 'preserve');
    fprintf('✓ 하위항목 데이터 로드 완료: %d명, %d컬럼\n', height(subitemData), width(subitemData));
catch ME
    fprintf('✗ 하위항목 데이터 로드 실패: %s\n', ME.message);
    return;
end

% ID 및 점수 컬럼 식별
subitemIdCol = findIDColumn(subitemData);
subitemIDs = extractAndStandardizeIDs(subitemData{:, subitemIdCol});

colNames = subitemData.Properties.VariableNames;
subitemScoreCols = {};

for col = 1:width(subitemData)
    colName = colNames{col};
    colData = subitemData{:, col};
    
    if col ~= subitemIdCol && isnumeric(colData) && sum(~isnan(colData)) > 10
        subitemScoreCols{end+1} = colName;
    end
end

fprintf('✓ 발견된 하위항목: %d개\n', length(subitemScoreCols));

%% 10-2. 성과 데이터와 매칭 (기존과 동일)
fprintf('\n▶ 성과 데이터와 매칭\n');

if exist('consolidatedScores', 'var') && istable(consolidatedScores)
    performanceData = consolidatedScores;
else
    fprintf('✗ 성과 데이터가 없습니다.\n');
    return;
end

% ID 매칭
if isnumeric(performanceData.ID)
    performanceIDs = arrayfun(@num2str, performanceData.ID, 'UniformOutput', false);
else
    performanceIDs = cellfun(@char, performanceData.ID, 'UniformOutput', false);
end

if isnumeric(subitemIDs)
    subitemIDs_str = arrayfun(@num2str, subitemIDs, 'UniformOutput', false);
else
    subitemIDs_str = cellfun(@char, subitemIDs, 'UniformOutput', false);
end

[commonIDs_sub, perfIdx, subIdx] = intersect(performanceIDs, subitemIDs_str);
fprintf('매칭 결과: %d명\n', length(commonIDs_sub));

if length(commonIDs_sub) < 15
    fprintf('✗ 매칭된 데이터가 너무 적습니다.\n');
    return;
end

%% 10-3. 분석용 데이터 구성
fprintf('\n▶ 분석용 데이터 구성\n');

% 특성 행렬 구성
X_all = [];
featureNames_all = {};

for i = 1:length(subitemScoreCols)
    featureNames_all{end+1} = subitemScoreCols{i};
    X_all = [X_all, subitemData.(subitemScoreCols{i})(subIdx)];
end

% 타겟 변수 설정
if ismember('AverageStdScore', performanceData.Properties.VariableNames)
    Y_all = performanceData.AverageStdScore(perfIdx);
    targetName = '역량진단 평균점수';
else
    % 대안 타겟 찾기
    scoreCols = performanceData.Properties.VariableNames;
    scoreCol = '';
    for col = 1:width(performanceData)
        if contains(scoreCols{col}, 'Score') || contains(scoreCols{col}, '점수')
            scoreCol = scoreCols{col};
            break;
        end
    end
    
    if ~isempty(scoreCol)
        Y_all = performanceData.(scoreCol)(perfIdx);
        targetName = scoreCol;
    else
        fprintf('✗ 적절한 성과 타겟 변수를 찾을 수 없습니다.\n');
        return;
    end
end

% 결측치 제거
validIdx = ~isnan(Y_all);
for i = 1:size(X_all, 2)
    validIdx = validIdx & ~isnan(X_all(:, i));
end

X_clean = X_all(validIdx, :);
Y_clean = Y_all(validIdx);
commonIDs_clean = commonIDs_sub(validIdx);

fprintf('✓ 최종 분석 데이터: %d명, %d개 특성\n', length(Y_clean), size(X_clean, 2));
fprintf('✓ 타겟 변수: %s\n', targetName);

%% 10-4. 1단계: 개별 상관분석
fprintf('\n▶ 1단계: 개별 상관분석 수행\n');

n_features = size(X_clean, 2);
correlation_results = table();
correlation_results.Feature = featureNames_all';
correlation_results.Correlation = NaN(n_features, 1);
correlation_results.P_Value = NaN(n_features, 1);
correlation_results.N_Valid = NaN(n_features, 1);

fprintf('각 하위항목과 %s 간의 상관분석:\n', targetName);
fprintf('%-25s %10s %10s %10s %10s\n', '특성명', '상관계수', 'p-value', '유효N', '유의성');
fprintf('%s\n', repmat('-', 1, 75));

significant_features = {};
significant_correlations = [];
significant_p_values = [];

for i = 1:n_features
    feature_data = X_clean(:, i);
    
    % 상관분석
    [r, p] = corr(feature_data, Y_clean);
    
    correlation_results.Correlation(i) = r;
    correlation_results.P_Value(i) = p;
    correlation_results.N_Valid(i) = length(Y_clean);
    
    % 유의성 표시
    sig_str = '';
    if p < 0.001
        sig_str = '***';
    elseif p < 0.01
        sig_str = '**';
    elseif p < 0.05
        sig_str = '*';
    end
    
    % 출력
    feature_name = featureNames_all{i};
    if length(feature_name) > 25
        feature_name = [feature_name(1:22), '...'];
    end
    
    fprintf('%-25s %10.3f %10.3f %10d %10s\n', ...
        feature_name, r, p, length(Y_clean), sig_str);
    
    % 유의미한 특성 수집 (p < 0.05 또는 |r| > 0.2)
    if p < 0.05 || abs(r) > 0.2
        significant_features{end+1} = featureNames_all{i};
        significant_correlations(end+1) = r;
        significant_p_values(end+1) = p;
    end
end

% 상관계수 순으로 정렬
[~, sort_idx] = sort(abs(correlation_results.Correlation), 'descend');
correlation_results = correlation_results(sort_idx, :);

fprintf('\n상관분석 요약:\n');
fprintf('  - 전체 특성: %d개\n', n_features);
fprintf('  - 유의미한 특성 (p<0.05 또는 |r|>0.2): %d개\n', length(significant_features));

% 유의미한 상관계수가 있는지 확인
strong_corr_count = sum(abs(correlation_results.Correlation) > 0.3);
moderate_corr_count = sum(abs(correlation_results.Correlation) > 0.2);
weak_corr_count = sum(abs(correlation_results.Correlation) > 0.1);

fprintf('  - 강한 상관 (|r|>0.3): %d개\n', strong_corr_count);
fprintf('  - 보통 상관 (|r|>0.2): %d개\n', moderate_corr_count);
fprintf('  - 약한 상관 (|r|>0.1): %d개\n', weak_corr_count);

%% 10-5. 상위 상관계수 특성들 상세 분석
fprintf('\n▶ 상위 상관계수 특성들 상세 분석\n');

top_n = min(15, height(correlation_results));
fprintf('상위 %d개 특성 상세 정보:\n', top_n);
fprintf('%-25s %10s %10s %10s %15s\n', '특성명', '상관계수', 'p-value', '효과크기', '해석');
fprintf('%s\n', repmat('-', 1, 85));

for i = 1:top_n
    feature = correlation_results.Feature{i};
    r = correlation_results.Correlation(i);
    p = correlation_results.P_Value(i);
    
    % 효과크기 해석
    if abs(r) >= 0.5
        effect_size = '큰 효과';
    elseif abs(r) >= 0.3
        effect_size = '중간 효과';
    elseif abs(r) >= 0.1
        effect_size = '작은 효과';
    else
        effect_size = '미미한 효과';
    end
    
    % 실용적 해석
    if p < 0.05 && abs(r) > 0.2
        interpretation = '유의미';
    elseif p < 0.05
        interpretation = '통계적 유의';
    elseif abs(r) > 0.2
        interpretation = '실용적 의미';
    else
        interpretation = '의미 제한적';
    end
    
    % 특성명 길이 조정
    if length(feature) > 25
        feature = [feature(1:22), '...'];
    end
    
    fprintf('%-25s %10.3f %10.3f %10s %15s\n', ...
        feature, r, p, effect_size, interpretation);
end

%% 10-6. 2단계: 유의미한 특성들로 예측 모델 구축
fprintf('\n▶ 2단계: 선별된 특성들로 예측 모델 구축\n');

% 유의미한 특성 선별 기준 설정
selection_criteria = struct();
selection_criteria.min_abs_corr = 0.15;  % 최소 절대 상관계수
selection_criteria.max_p_value = 0.10;   % 최대 p-value
selection_criteria.max_features = 10;    % 최대 특성 수

% 기준에 따른 특성 선별
selected_mask = (abs(correlation_results.Correlation) >= selection_criteria.min_abs_corr) & ...
                (correlation_results.P_Value <= selection_criteria.max_p_value);

selected_features_table = correlation_results(selected_mask, :);

% 상위 N개로 제한
if height(selected_features_table) > selection_criteria.max_features
    selected_features_table = selected_features_table(1:selection_criteria.max_features, :);
end

if height(selected_features_table) == 0
    fprintf('선별 기준을 만족하는 특성이 없습니다.\n');
    fprintf('기준을 완화하여 상위 5개 특성을 선택합니다.\n');
    selected_features_table = correlation_results(1:min(5, height(correlation_results)), :);
end

selected_feature_names = selected_features_table.Feature;
n_selected = length(selected_feature_names);

fprintf('선별된 특성 (%d개):\n', n_selected);
for i = 1:n_selected
    fprintf('  %d. %s (r=%.3f, p=%.3f)\n', i, selected_feature_names{i}, ...
        selected_features_table.Correlation(i), selected_features_table.P_Value(i));
end

% 선별된 특성들의 데이터 추출
selected_indices = zeros(n_selected, 1);
for i = 1:n_selected
    selected_indices(i) = find(strcmp(featureNames_all, selected_feature_names{i}));
end

X_selected = X_clean(:, selected_indices);

%% 10-7. 선별된 특성들로 다양한 모델 테스트
fprintf('\n▶ 선별된 특성들로 다양한 모델 테스트\n');

% 데이터 표준화
X_selected_std = zeros(size(X_selected));
for i = 1:size(X_selected, 2)
    X_selected_std(:, i) = (X_selected(:, i) - mean(X_selected(:, i))) / (std(X_selected(:, i)) + eps);
end

% 훈련/테스트 분할
test_ratio = 0.25;
if length(Y_clean) < 20
    test_ratio = 0.2;  % 데이터가 적으면 테스트 비율 줄임
end

cv_partition = cvpartition(length(Y_clean), 'HoldOut', test_ratio);
X_train = X_selected_std(training(cv_partition), :);
X_test = X_selected_std(test(cv_partition), :);
Y_train = Y_clean(training(cv_partition));
Y_test = Y_clean(test(cv_partition));

fprintf('훈련: %d명, 테스트: %d명, 특성: %d개\n', length(Y_train), length(Y_test), size(X_train, 2));

% 모델 성능 저장
model_results = struct();

% 1. 다중선형회귀 (OLS)
fprintf('\n1. 다중선형회귀 (OLS):\n');
try
    if size(X_train, 2) < size(X_train, 1)  % 특성 수 < 샘플 수
        beta_ols = [ones(size(X_train, 1), 1), X_train] \ Y_train;
        Y_pred_ols = [ones(size(X_test, 1), 1), X_test] * beta_ols;
        
        r2_ols = 1 - sum((Y_test - Y_pred_ols).^2) / sum((Y_test - mean(Y_test)).^2);
        rmse_ols = sqrt(mean((Y_test - Y_pred_ols).^2));
        corr_ols = corr(Y_test, Y_pred_ols);
        
        model_results.ols.r2 = r2_ols;
        model_results.ols.rmse = rmse_ols;
        model_results.ols.corr = corr_ols;
        model_results.ols.coeffs = beta_ols(2:end);
        
        fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f\n', r2_ols, rmse_ols, corr_ols);
    else
        fprintf('   건너뜀 (특성 수 > 샘플 수)\n');
        model_results.ols.r2 = NaN;
    end
catch
    fprintf('   실패\n');
    model_results.ols.r2 = NaN;
end

% 2. Ridge 회귀
fprintf('\n2. Ridge 회귀:\n');
try
    [B_ridge, FitInfo_ridge] = lasso(X_train, Y_train, 'Alpha', 0, 'CV', 3);
    Y_pred_ridge = X_test * B_ridge(:, FitInfo_ridge.IndexMinMSE) + FitInfo_ridge.Intercept(FitInfo_ridge.IndexMinMSE);
    
    r2_ridge = 1 - sum((Y_test - Y_pred_ridge).^2) / sum((Y_test - mean(Y_test)).^2);
    rmse_ridge = sqrt(mean((Y_test - Y_pred_ridge).^2));
    corr_ridge = corr(Y_test, Y_pred_ridge);
    
    model_results.ridge.r2 = r2_ridge;
    model_results.ridge.rmse = rmse_ridge;
    model_results.ridge.corr = corr_ridge;
    model_results.ridge.coeffs = B_ridge(:, FitInfo_ridge.IndexMinMSE);
    model_results.ridge.lambda = FitInfo_ridge.Lambda(FitInfo_ridge.IndexMinMSE);
    
    fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f, λ = %.4f\n', ...
        r2_ridge, rmse_ridge, corr_ridge, model_results.ridge.lambda);
catch
    fprintf('   실패\n');
    model_results.ridge.r2 = NaN;
end

% 3. Elastic Net (관대한 설정)
fprintf('\n3. Elastic Net (관대한 설정):\n');
try
    [B_elastic, FitInfo_elastic] = lasso(X_train, Y_train, 'Alpha', 0.2, 'CV', 3, 'LambdaRatio', 1e-4);
    Y_pred_elastic = X_test * B_elastic(:, FitInfo_elastic.IndexMinMSE) + FitInfo_elastic.Intercept(FitInfo_elastic.IndexMinMSE);
    
    r2_elastic = 1 - sum((Y_test - Y_pred_elastic).^2) / sum((Y_test - mean(Y_test)).^2);
    rmse_elastic = sqrt(mean((Y_test - Y_pred_elastic).^2));
    corr_elastic = corr(Y_test, Y_pred_elastic);
    
    model_results.elastic.r2 = r2_elastic;
    model_results.elastic.rmse = rmse_elastic;
    model_results.elastic.corr = corr_elastic;
    model_results.elastic.coeffs = B_elastic(:, FitInfo_elastic.IndexMinMSE);
    model_results.elastic.lambda = FitInfo_elastic.Lambda(FitInfo_elastic.IndexMinMSE);
    
    fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f, λ = %.4f\n', ...
        r2_elastic, rmse_elastic, corr_elastic, model_results.elastic.lambda);
catch
    fprintf('   실패\n');
    model_results.elastic.r2 = NaN;
end

% 4. 단순 평균 기반 모델
fprintf('\n4. 단순 평균 기반:\n');
% 상관계수를 가중치로 사용한 선형결합
weights = selected_features_table.Correlation / sum(abs(selected_features_table.Correlation));
Y_pred_weighted = X_test * weights;

% 스케일 조정
Y_pred_weighted = Y_pred_weighted * std(Y_train) / std(Y_pred_weighted) + mean(Y_train);

r2_weighted = 1 - sum((Y_test - Y_pred_weighted).^2) / sum((Y_test - mean(Y_test)).^2);
rmse_weighted = sqrt(mean((Y_test - Y_pred_weighted).^2));
corr_weighted = corr(Y_test, Y_pred_weighted);

model_results.weighted.r2 = r2_weighted;
model_results.weighted.rmse = rmse_weighted;
model_results.weighted.corr = corr_weighted;
model_results.weighted.coeffs = weights;

fprintf('   R² = %.4f, RMSE = %.4f, 상관계수 = %.4f\n', r2_weighted, rmse_weighted, corr_weighted);

%% 10-8. 최고 성능 모델 선택 및 분석
fprintf('\n▶ 모델 성능 비교 및 최종 결과\n');

% 성능 비교
model_names = {'OLS', 'Ridge', 'Elastic Net', 'Weighted'};
r2_values = [model_results.ols.r2, model_results.ridge.r2, model_results.elastic.r2, model_results.weighted.r2];
rmse_values = [model_results.ols.rmse, model_results.ridge.rmse, model_results.elastic.rmse, model_results.weighted.rmse];

fprintf('모델 성능 비교:\n');
fprintf('%-12s %10s %10s\n', '모델', 'R²', 'RMSE');
fprintf('%s\n', repmat('-', 1, 35));

for i = 1:length(model_names)
    if ~isnan(r2_values(i))
        fprintf('%-12s %10.4f %10.4f\n', model_names{i}, r2_values(i), rmse_values(i));
    else
        fprintf('%-12s %10s %10s\n', model_names{i}, 'N/A', 'N/A');
    end
end

% 최고 성능 모델 선택
valid_r2 = r2_values(~isnan(r2_values));
if ~isempty(valid_r2)
    [best_r2, best_idx_temp] = max(valid_r2);
    valid_indices = find(~isnan(r2_values));
    best_idx = valid_indices(best_idx_temp);
    best_model_name = model_names{best_idx};
    
    fprintf('\n최고 성능 모델: %s (R² = %.4f)\n', best_model_name, best_r2);
else
    fprintf('\n모든 모델이 실패했습니다.\n');
    best_r2 = NaN;
    best_model_name = 'None';
end

%% 10-9. 결과 저장
fprintf('\n▶ 결과 저장\n');

if exist('outputFileName', 'var') && ~isempty(outputFileName)
    saveFileName = outputFileName;
else
    saveFileName = sprintf('상관분석_하위항목분석_%s.xlsx', datestr(now, 'yyyymmdd_HHMM'));
end

try
    % 상관분석 결과 저장
    writetable(correlation_results, saveFileName, 'Sheet', '개별_상관분석결과');
    fprintf('✓ 상관분석 결과 저장: %d개 특성\n', height(correlation_results));
    
    % 선별된 특성 결과 저장
    writetable(selected_features_table, saveFileName, 'Sheet', '선별된_특성들');
    fprintf('✓ 선별된 특성 저장: %d개 특성\n', height(selected_features_table));
    
    % 모델 성능 비교 저장
    comparison_table = table();
    comparison_table.Model = model_names';
    comparison_table.R_squared = r2_values';
    comparison_table.RMSE = rmse_values';
    
    writetable(comparison_table, saveFileName, 'Sheet', '모델_성능비교');
    fprintf('✓ 모델 성능 비교 저장\n');
    
catch ME
    fprintf('✗ 결과 저장 실패: %s\n', ME.message);
    saveFileName = sprintf('상관분석_하위항목분석_%s.xlsx', datestr(now, 'yyyymmdd_HHMM'));
    try
        writetable(correlation_results, saveFileName, 'Sheet', '상관분석결과');
        writetable(selected_features_table, saveFileName, 'Sheet', '선별된특성');
        writetable(comparison_table, saveFileName, 'Sheet', '모델성능');
        fprintf('✓ 새 파일로 저장: %s\n', saveFileName);
    catch
        fprintf('✗ 저장 완전 실패\n');
    end
end

%% 10-10. 시각화
fprintf('\n▶ 결과 시각화\n');

figure('Name', '상관분석 기반 하위항목 분석', 'Position', [50, 50, 1400, 900]);

% 1. 상관계수 분포
subplot(2, 3, 1);
histogram(correlation_results.Correlation, 15);
xlabel('상관계수');
ylabel('빈도');
title('하위항목별 상관계수 분포', 'FontWeight', 'bold');
grid on;

% 2. 상위 특성 상관계수
subplot(2, 3, 2);
top_10_idx = 1:min(10, height(correlation_results));
barh(top_10_idx, correlation_results.Correlation(top_10_idx));
set(gca, 'YTick', top_10_idx, 'YTickLabel', correlation_results.Feature(top_10_idx));
title('상위 10개 특성 상관계수', 'FontWeight', 'bold');
xlabel('상관계수');
grid on;

% 3. p-value vs 상관계수 산점도
subplot(2, 3, 3);
scatter(correlation_results.Correlation, -log10(correlation_results.P_Value), 'filled');
xlabel('상관계수');
ylabel('-log10(p-value)');
title('상관계수 vs 유의성', 'FontWeight', 'bold');
% 유의성 선 추가
hold on;
yline(-log10(0.05), 'r--', 'p=0.05');
yline(-log10(0.01), 'r--', 'p=0.01');
hold off;
grid on;

% 4. 최고 모델 예측 vs 실제
subplot(2, 3, 4);
if ~isnan(best_r2)
    if strcmp(best_model_name, 'OLS') && ~isnan(model_results.ols.r2)
        Y_pred_best = [ones(size(X_test, 1), 1), X_test] * [model_results.ols.coeffs(1); model_results.ols.coeffs(2:end)];
    elseif strcmp(best_model_name, 'Ridge')
        Y_pred_best = X_test * model_results.ridge.coeffs + FitInfo_ridge.Intercept(FitInfo_ridge.IndexMinMSE);
    elseif strcmp(best_model_name, 'Elastic Net')
        Y_pred_best = X_test * model_results.elastic.coeffs + FitInfo_elastic.Intercept(FitInfo_elastic.IndexMinMSE);
    else
        Y_pred_best = Y_pred_weighted;
    end
    
    scatter(Y_test, Y_pred_best, 'filled');
    hold on;
    lim_min = min([Y_test; Y_pred_best]);
    lim_max = max([Y_test; Y_pred_best]);
    plot([lim_min, lim_max], [lim_min, lim_max], 'r--', 'LineWidth', 2);
    xlabel('실제 성과');
    ylabel('예측 성과');
    title(sprintf('%s 모델 (R²=%.3f)', best_model_name, best_r2), 'FontWeight', 'bold');
    grid on;
    hold off;
else
    text(0.5, 0.5, '예측 모델 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('예측 결과', 'FontWeight', 'bold');
end

% 5. 모델 성능 비교
subplot(2, 3, 5);
valid_models = ~isnan(r2_values);
if sum(valid_models) > 0
    bar(r2_values(valid_models));
    set(gca, 'XTickLabel', model_names(valid_models));
    ylabel('R²');
    title('모델별 성능 비교', 'FontWeight', 'bold');
    xtickangle(45);
    grid on;
end

% 6. 선별된 특성들의 상관계수
subplot(2, 3, 6);
if n_selected > 0
    barh(1:n_selected, selected_features_table.Correlation);
    set(gca, 'YTick', 1:n_selected, 'YTickLabel', selected_features_table.Feature);
    title('선별된 특성들의 상관계수', 'FontWeight', 'bold');
    xlabel('상관계수');
    grid on;
else
    text(0.5, 0.5, '선별된 특성 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('선별된 특성들', 'FontWeight', 'bold');
end

sgtitle('상관분석 기반 하위항목 분석 결과', 'FontSize', 14, 'FontWeight', 'bold');

%% 최종 요약
fprintf('\n[상관분석 기반 하위항목 분석 완료]\n');
fprintf('=====================================\n');

fprintf('분석 결과 요약:\n');
fprintf('  • 전체 하위항목: %d개\n', n_features);
fprintf('  • 분석 대상자: %d명\n', length(Y_clean));
fprintf('  • 선별된 특성: %d개\n', n_selected);

fprintf('\n상관분석 주요 발견:\n');
fprintf('  • 강한 상관 (|r|>0.3): %d개\n', strong_corr_count);
fprintf('  • 보통 상관 (|r|>0.2): %d개\n', moderate_corr_count);
fprintf('  • 약한 상관 (|r|>0.1): %d개\n', weak_corr_count);

if height(correlation_results) > 0
    fprintf('\n최고 상관계수 특성들:\n');
    for i = 1:min(5, height(correlation_results))
        fprintf('  %d. %s (r=%.3f, p=%.3f)\n', i, correlation_results.Feature{i}, ...
            correlation_results.Correlation(i), correlation_results.P_Value(i));
    end
end

if ~isnan(best_r2)
    fprintf('\n예측 모델 성능:\n');
    fprintf('  • 최고 모델: %s\n', best_model_name);
    fprintf('  • R²: %.4f\n', best_r2);
    fprintf('  • 사용된 특성: %d개\n', n_selected);
    
    if best_r2 > 0.3
        fprintf('\n✅ 선별된 특성들이 성과를 잘 예측합니다 (R² > 0.3).\n');
        fprintf('   이 특성들에 집중한 역량 개발을 권장합니다.\n');
    elseif best_r2 > 0.1
        fprintf('\n⚠ 선별된 특성들이 어느 정도 예측력을 보입니다 (R² > 0.1).\n');
        fprintf('   참고용으로 활용 가능하나 추가 검증이 필요합니다.\n');
    elseif best_r2 > 0
        fprintf('\n📊 선별된 특성들이 약한 예측력을 보입니다 (R² > 0).\n');
        fprintf('   개별 상관관계는 의미가 있으나 종합적 예측은 제한적입니다.\n');
    end
else
    fprintf('\n❌ 예측 모델 구축에 실패했습니다.\n');
    fprintf('   개별 상관관계만 참고하시기 바랍니다.\n');
end

fprintf('\n실용적 활용 방안:\n');
if strong_corr_count > 0
    fprintf('  • 강한 상관을 보인 특성들을 우선 집중 개발\n');
end
if moderate_corr_count > 0
    fprintf('  • 보통 상관 특성들을 보조적 개발 영역으로 활용\n');
end
fprintf('  • 상관계수가 낮은 특성들은 다른 성과 지표와의 관계 탐색\n');
fprintf('  • 개별 특성별 맞춤형 개선 프로그램 설계\n');

fprintf('\n저장된 파일:\n');
fprintf('  • %s\n', saveFileName);
fprintf('    - 시트1: 개별_상관분석결과 (모든 특성의 상관계수)\n');
fprintf('    - 시트2: 선별된_특성들 (유의미한 특성들)\n');
fprintf('    - 시트3: 모델_성능비교 (예측 모델 성능)\n');

if n_selected > 0 && ~isnan(best_r2)
    fprintf('\n✅ 상관분석 기반 특성 선별 및 예측 모델링 완료!\n');
    fprintf('   단계적 접근으로 유의미한 특성들을 식별했습니다.\n');
else
    fprintf('\n📋 상관분석 완료!\n');
    fprintf('   개별 특성들의 상관관계를 파악했습니다.\n');
    fprintf('   예측 모델링을 위해서는 추가 데이터나 다른 접근법이 필요할 수 있습니다.\n');
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

%% 10-8. 최고 성능 모델 선택 및 분석
fprintf('\n▶ 모델 성능 비교 및 최종 결과\n');

% 성능 비교
model_names = {'OLS', 'Ridge', 'Elastic Net', 'Weighted'};
r2_values = [model_results.ols.r2, model_results.ridge.r2, model_results.elastic.r2, model_results.weighted.r2];
rmse_values = [model_results.ols.rmse, model_results.ridge.rmse, model_results.elastic.rmse, model_results.weighted.rmse];

fprintf('모델 성능 비교:\n');
fprintf('%-12s %10s %10s\n', '모델', 'R²', 'RMSE');
fprintf('%s\n', repmat('-', 1, 35));

for i = 1:length(model_names)
    if ~isnan(r2_values(i))
        fprintf('%-12s %10.4f %10.4f\n', model_names{i}, r2_values(i), rmse_values(i));
    else
        fprintf('%-12s %10s %10s\n', model_names{i}, 'N/A', 'N/A');
    end
end

% 최고 성능 모델 선택
valid_r2 = r2_values(~isnan(r2_values));
if ~isempty(valid_r2)
    [best_r2, best_idx_temp] = max(valid_r2);
    valid_indices = find(~isnan(r2_values));
    best_idx = valid_indices(best_idx_temp);
    best_model_name = model_names{best_idx};
    
    fprintf('\n최고 성능 모델: %s (R² = %.4f)\n', best_model_name, best_r2);
else
    fprintf('\n모든 모델이 실패했습니다.\n');
    best_r2 = NaN;
    best_model_name = 'None';
end

%% 10-9. 결과 저장
fprintf('\n▶ 결과 저장\n');

if exist('outputFileName', 'var') && ~isempty(outputFileName)
    saveFileName = outputFileName;
else
    saveFileName = sprintf('상관분석_하위항목분석_%s.xlsx', datestr(now, 'yyyymmdd_HHMM'));
end

try
    % 상관분석 결과 저장
    writetable(correlation_results, saveFileName, 'Sheet', '개별_상관분석결과');
    fprintf('✓ 상관분석 결과 저장: %d개 특성\n', height(correlation_results));
    
    % 선별된 특성 결과 저장
    writetable(selected_features_table, saveFileName, 'Sheet', '선별된_특성들');
    fprintf('✓ 선별된 특성 저장: %d개 특성\n', height(selected_features_table));
    
    % 모델 성능 비교 저장
    comparison_table = table();
    comparison_table.Model = model_names';
    comparison_table.R_squared = r2_values';
    comparison_table.RMSE = rmse_values';
    
    writetable(comparison_table, saveFileName, 'Sheet', '모델_성능비교');
    fprintf('✓ 모델 성능 비교 저장\n');
    
catch ME
    fprintf('✗ 결과 저장 실패: %s\n', ME.message);
    saveFileName = sprintf('상관분석_하위항목분석_%s.xlsx', datestr(now, 'yyyymmdd_HHMM'));
    try
        writetable(correlation_results, saveFileName, 'Sheet', '상관분석결과');
        writetable(selected_features_table, saveFileName, 'Sheet', '선별된특성');
        writetable(comparison_table, saveFileName, 'Sheet', '모델성능');
        fprintf('✓ 새 파일로 저장: %s\n', saveFileName);
    catch
        fprintf('✗ 저장 완전 실패\n');
    end
end

%% 10-10. 시각화
fprintf('\n▶ 결과 시각화\n');

figure('Name', '상관분석 기반 하위항목 분석', 'Position', [50, 50, 1400, 900]);

% 1. 상관계수 분포
subplot(2, 3, 1);
histogram(correlation_results.Correlation, 15);
xlabel('상관계수');
ylabel('빈도');
title('하위항목별 상관계수 분포', 'FontWeight', 'bold');
grid on;

% 2. 상위 특성 상관계수
subplot(2, 3, 2);
top_10_idx = 1:min(10, height(correlation_results));
barh(top_10_idx, correlation_results.Correlation(top_10_idx));
set(gca, 'YTick', top_10_idx, 'YTickLabel', correlation_results.Feature(top_10_idx));
title('상위 10개 특성 상관계수', 'FontWeight', 'bold');
xlabel('상관계수');
grid on;

% 3. p-value vs 상관계수 산점도
subplot(2, 3, 3);
scatter(correlation_results.Correlation, -log10(correlation_results.P_Value), 'filled');
xlabel('상관계수');
ylabel('-log10(p-value)');
title('상관계수 vs 유의성', 'FontWeight', 'bold');
% 유의성 선 추가
hold on;
yline(-log10(0.05), 'r--', 'p=0.05');
yline(-log10(0.01), 'r--', 'p=0.01');
hold off;
grid on;

% 4. 최고 모델 예측 vs 실제
subplot(2, 3, 4);
if ~isnan(best_r2)
    if strcmp(best_model_name, 'OLS') && ~isnan(model_results.ols.r2)
        Y_pred_best = [ones(size(X_test, 1), 1), X_test] * [model_results.ols.coeffs(1); model_results.ols.coeffs(2:end)];
    elseif strcmp(best_model_name, 'Ridge')
        Y_pred_best = X_test * model_results.ridge.coeffs + FitInfo_ridge.Intercept(FitInfo_ridge.IndexMinMSE);
    elseif strcmp(best_model_name, 'Elastic Net')
        Y_pred_best = X_test * model_results.elastic.coeffs + FitInfo_elastic.Intercept(FitInfo_elastic.IndexMinMSE);
    else
        Y_pred_best = Y_pred_weighted;
    end
    
    scatter(Y_test, Y_pred_best, 'filled');
    hold on;
    lim_min = min([Y_test; Y_pred_best]);
    lim_max = max([Y_test; Y_pred_best]);
    plot([lim_min, lim_max], [lim_min, lim_max], 'r--', 'LineWidth', 2);
    xlabel('실제 성과');
    ylabel('예측 성과');
    title(sprintf('%s 모델 (R²=%.3f)', best_model_name, best_r2), 'FontWeight', 'bold');
    grid on;
    hold off;
else
    text(0.5, 0.5, '예측 모델 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('예측 결과', 'FontWeight', 'bold');
end

% 5. 모델 성능 비교
subplot(2, 3, 5);
valid_models = ~isnan(r2_values);
if sum(valid_models) > 0
    bar(r2_values(valid_models));
    set(gca, 'XTickLabel', model_names(valid_models));
    ylabel('R²');
    title('모델별 성능 비교', 'FontWeight', 'bold');
    xtickangle(45);
    grid on;
end

% 6. 선별된 특성들의 상관계수
subplot(2, 3, 6);
if n_selected > 0
    barh(1:n_selected, selected_features_table.Correlation);
    set(gca, 'YTick', 1:n_selected, 'YTickLabel', selected_features_table.Feature);
    title('선별된 특성들의 상관계수', 'FontWeight', 'bold');
    xlabel('상관계수');
    grid on;
else
    text(0.5, 0.5, '선별된 특성 없음', 'Units', 'normalized', 'HorizontalAlignment', 'center');
    title('선별된 특성들', 'FontWeight', 'bold');
end

sgtitle('상관분석 기반 하위항목 분석 결과', 'FontSize', 14, 'FontWeight', 'bold');

%% 최종 요약
fprintf('\n[상관분석 기반 하위항목 분석 완료]\n');
fprintf('=====================================\n');

fprintf('분석 결과 요약:\n');
fprintf('  • 전체 하위항목: %d개\n', n_features);
fprintf('  • 분석 대상자: %d명\n', length(Y_clean));
fprintf('  • 선별된 특성: %d개\n', n_selected);

fprintf('\n상관분석 주요 발견:\n');
fprintf('  • 강한 상관 (|r|>0.3): %d개\n', strong_corr_count);
fprintf('  • 보통 상관 (|r|>0.2): %d개\n', moderate_corr_count);
fprintf('  • 약한 상관 (|r|>0.1): %d개\n', weak_corr_count);

if height(correlation_results) > 0
    fprintf('\n최고 상관계수 특성들:\n');
    for i = 1:min(5, height(correlation_results))
        fprintf('  %d. %s (r=%.3f, p=%.3f)\n', i, correlation_results.Feature{i}, ...
            correlation_results.Correlation(i), correlation_results.P_Value(i));
    end
end

if ~isnan(best_r2)
    fprintf('\n예측 모델 성능:\n');
    fprintf('  • 최고 모델: %s\n', best_model_name);
    fprintf('  • R²: %.4f\n', best_r2);
    fprintf('  • 사용된 특성: %d개\n', n_selected);
    
    if best_r2 > 0.3
        fprintf('\n✅ 선별된 특성들이 성과를 잘 예측합니다 (R² > 0.3).\n');
        fprintf('   이 특성들에 집중한 역량 개발을 권장합니다.\n');
    elseif best_r2 > 0.1
        fprintf('\n⚠ 선별된 특성들이 어느 정도 예측력을 보입니다 (R² > 0.1).\n');
        fprintf('   참고용으로 활용 가능하나 추가 검증이 필요합니다.\n');
    elseif best_r2 > 0
        fprintf('\n📊 선별된 특성들이 약한 예측력을 보입니다 (R² > 0).\n');
        fprintf('   개별 상관관계는 의미가 있으나 종합적 예측은 제한적입니다.\n');
    end
else
    fprintf('\n❌ 예측 모델 구축에 실패했습니다.\n');
    fprintf('   개별 상관관계만 참고하시기 바랍니다.\n');
end

fprintf('\n실용적 활용 방안:\n');
if strong_corr_count > 0
    fprintf('  • 강한 상관을 보인 특성들을 우선 집중 개발\n');
end
if moderate_corr_count > 0
    fprintf('  • 보통 상관 특성들을 보조적 개발 영역으로 활용\n');
end
fprintf('  • 상관계수가 낮은 특성들은 다른 성과 지표와의 관계 탐색\n');
fprintf('  • 개별 특성별 맞춤형 개선 프로그램 설계\n');

fprintf('\n저장된 파일:\n');
fprintf('  • %s\n', saveFileName);
fprintf('    - 시트1: 개별_상관분석결과 (모든 특성의 상관계수)\n');
fprintf('    - 시트2: 선별된_특성들 (유의미한 특성들)\n');
fprintf('    - 시트3: 모델_성능비교 (예측 모델 성능)\n');

if n_selected > 0 && ~isnan(best_r2)
    fprintf('\n✅ 상관분석 기반 특성 선별 및 예측 모델링 완료!\n');
    fprintf('   단계적 접근으로 유의미한 특성들을 식별했습니다.\n');
else
    fprintf('\n📋 상관분석 완료!\n');
    fprintf('   개별 특성들의 상관관계를 파악했습니다.\n');
    fprintf('   예측 모델링을 위해서는 추가 데이터나 다른 접근법이 필요할 수 있습니다.\n');
end

