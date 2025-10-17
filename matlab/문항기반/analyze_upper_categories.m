%% 역량검사 상위항목 데이터 구조 분석
%
% 목적: 역량검사_상위항목 시트의 구조를 파악하고
%       성과점수와의 상관분석 및 중다회귀분석 준비
%

clear; clc; close all;

fprintf('========================================\n');
fprintf('역량검사 상위항목 데이터 구조 분석\n');
fprintf('========================================\n\n');

%% 1. 데이터 로드
fprintf('[1단계] 역량검사 상위항목 데이터 로드\n');
fprintf('----------------------------------------\n');

competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

try
    % 상위항목 데이터 로드
    upperCategoryData = readtable(competencyTestPath, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
    fprintf('✓ 역량검사_상위항목 데이터 로드 완료: %d명 × %d개 변수\n', height(upperCategoryData), width(upperCategoryData));

    % 컬럼명 출력
    fprintf('\n▶ 컬럼 구조:\n');
    colNames = upperCategoryData.Properties.VariableNames;
    for i = 1:length(colNames)
        colData = upperCategoryData{:, i};
        if isnumeric(colData)
            validData = colData(~isnan(colData));
            if ~isempty(validData)
                fprintf('  %2d. %-30s (숫자형): 범위 %.1f~%.1f, 평균 %.2f (N=%d)\n', ...
                    i, colNames{i}, min(validData), max(validData), mean(validData), length(validData));
            else
                fprintf('  %2d. %-30s (숫자형): 모든 값이 결측\n', i, colNames{i});
            end
        else
            uniqueVals = unique(colData);
            if length(uniqueVals) <= 10
                fprintf('  %2d. %-30s (범주형): %s\n', i, colNames{i}, strjoin(string(uniqueVals(1:min(5, end))), ', '));
            else
                fprintf('  %2d. %-30s (텍스트형): %d개 고유값\n', i, colNames{i}, length(uniqueVals));
            end
        end
    end

    % 종합점수 데이터도 로드 (비교용)
    competencyTotalData = readtable(competencyTestPath, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('\n✓ 역량검사_종합점수 데이터 로드 완료: %d명 × %d개 변수\n', height(competencyTotalData), width(competencyTotalData));

catch ME
    fprintf('✗ 데이터 로드 실패: %s\n', ME.message);
    return;
end

%% 2. ID 매칭 확인
fprintf('\n[2단계] ID 매칭 확인\n');
fprintf('----------------------------------------\n');

% ID 컬럼 찾기
upperIDCol = findIDColumn(upperCategoryData);
totalIDCol = findIDColumn(competencyTotalData);

if isempty(upperIDCol) || isempty(totalIDCol)
    fprintf('✗ ID 컬럼을 찾을 수 없습니다\n');
    return;
end

fprintf('✓ ID 컬럼 확인:\n');
fprintf('  - 상위항목: "%s"\n', upperCategoryData.Properties.VariableNames{upperIDCol});
fprintf('  - 종합점수: "%s"\n', competencyTotalData.Properties.VariableNames{totalIDCol});

% ID 매칭
upperIDs = extractAndStandardizeIDs(upperCategoryData{:, upperIDCol});
totalIDs = extractAndStandardizeIDs(competencyTotalData{:, totalIDCol});

[commonIDs, upperIdx, totalIdx] = intersect(upperIDs, totalIDs);
fprintf('  - 공통 ID: %d명\n', length(commonIDs));

%% 3. 상위항목 점수 컬럼 식별
fprintf('\n[3단계] 상위항목 점수 컬럼 식별\n');
fprintf('----------------------------------------\n');

% 숫자형 컬럼 중 점수로 보이는 것들 찾기
scoreColumns = {};
scoreColumnNames = {};

for i = 1:width(upperCategoryData)
    if i == upperIDCol
        continue; % ID 컬럼 제외
    end

    colName = colNames{i};
    colData = upperCategoryData{:, i};

    if isnumeric(colData)
        validData = colData(~isnan(colData));
        if ~isempty(validData) && var(validData) > 0
            % 점수 컬럼으로 판단 (분산이 0이 아닌 숫자형)
            scoreColumns{end+1} = i;
            scoreColumnNames{end+1} = colName;
        end
    end
end

fprintf('✓ 발견된 상위항목 점수 컬럼: %d개\n', length(scoreColumns));
for i = 1:length(scoreColumnNames)
    colIdx = scoreColumns{i};
    colData = upperCategoryData{:, colIdx};
    validData = colData(~isnan(colData));

    fprintf('  %d. %-30s: 범위 %.1f~%.1f, 평균 %.2f (N=%d)\n', ...
        i, scoreColumnNames{i}, min(validData), max(validData), mean(validData), length(validData));
end

%% 4. 매칭된 데이터셋 생성
fprintf('\n[4단계] 매칭된 데이터셋 생성\n');
fprintf('----------------------------------------\n');

if length(commonIDs) < 5
    fprintf('✗ 매칭된 데이터가 부족합니다 (%d명)\n', length(commonIDs));
    return;
end

% 매칭된 상위항목 데이터
matchedUpperData = upperCategoryData(upperIdx, :);
matchedTotalData = competencyTotalData(totalIdx, :);

% 상위항목 점수 행렬 생성
upperScoreMatrix = [];
upperScoreNames = {};

for i = 1:length(scoreColumns)
    colIdx = scoreColumns{i};
    colData = matchedUpperData{:, colIdx};
    upperScoreMatrix = [upperScoreMatrix, colData];
    upperScoreNames{end+1} = scoreColumnNames{i};
end

fprintf('✓ 매칭된 데이터셋 생성 완료:\n');
fprintf('  - 공통 ID: %d명\n', length(commonIDs));
fprintf('  - 상위항목 점수: %d개 범주\n', size(upperScoreMatrix, 2));

% 결측치 처리
validRows = sum(isnan(upperScoreMatrix), 2) < (size(upperScoreMatrix, 2) * 0.5);
cleanUpperScoreMatrix = upperScoreMatrix(validRows, :);
cleanCommonIDs = commonIDs(validRows);

fprintf('  - 결측치 처리 후: %d명\n', size(cleanUpperScoreMatrix, 1));

%% 5. 기초 통계 및 상관분석
fprintf('\n[5단계] 상위항목 간 상관분석\n');
fprintf('----------------------------------------\n');

if size(cleanUpperScoreMatrix, 1) < 3
    fprintf('✗ 상관분석을 위한 데이터가 부족합니다\n');
    return;
end

% 상위항목 간 상관계수 매트릭스
try
    upperCorrMatrix = corrcoef(cleanUpperScoreMatrix, 'Rows', 'pairwise');

    fprintf('✓ 상위항목 간 상관분석 완료\n');
    fprintf('\n▶ 상위항목 간 주요 상관계수:\n');

    % 높은 상관계수 출력
    [maxCorr, maxIdx] = max(abs(upperCorrMatrix - eye(size(upperCorrMatrix))), [], 'all');
    [row, col] = ind2sub(size(upperCorrMatrix), maxIdx);

    fprintf('  - 최고 상관: %s ↔ %s (r = %.3f)\n', ...
        upperScoreNames{row}, upperScoreNames{col}, upperCorrMatrix(row, col));

    % 상관계수 매트릭스 히트맵 생성 (선택적)
    if length(upperScoreNames) <= 10
        figure('Name', '상위항목 간 상관계수 매트릭스', 'Position', [100, 100, 800, 600]);
        imagesc(upperCorrMatrix);
        colorbar;
        colormap('RdBu');
        caxis([-1, 1]);

        set(gca, 'XTick', 1:length(upperScoreNames));
        set(gca, 'YTick', 1:length(upperScoreNames));
        set(gca, 'XTickLabel', upperScoreNames, 'XTickLabelRotation', 45);
        set(gca, 'YTickLabel', upperScoreNames);

        title('상위항목 간 상관계수 매트릭스', 'FontSize', 14, 'FontWeight', 'bold');

        % 상관계수 값 텍스트로 표시
        for i = 1:length(upperScoreNames)
            for j = 1:length(upperScoreNames)
                text(j, i, sprintf('%.2f', upperCorrMatrix(i,j)), ...
                    'HorizontalAlignment', 'center', 'FontSize', 10);
            end
        end
    end

catch ME
    fprintf('✗ 상관분석 실패: %s\n', ME.message);
end

%% 6. 결과 저장
fprintf('\n[6단계] 분석 결과 저장\n');
fprintf('----------------------------------------\n');

dateStr = datestr(now, 'yyyymmdd_HHMM');
outputFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\upper_category_analysis_%s.xlsx', dateStr);

try
    % 상위항목 점수 데이터 저장
    upperScoreTable = array2table(cleanUpperScoreMatrix, 'VariableNames', upperScoreNames);
    upperScoreTable.ID = cleanCommonIDs;
    upperScoreTable = upperScoreTable(:, [end, 1:end-1]); % ID를 첫 번째 컬럼으로

    writetable(upperScoreTable, outputFileName, 'Sheet', '상위항목점수데이터');

    % 상관계수 매트릭스 저장
    if exist('upperCorrMatrix', 'var')
        corrTable = array2table(upperCorrMatrix, ...
            'VariableNames', upperScoreNames, 'RowNames', upperScoreNames);
        writetable(corrTable, outputFileName, 'Sheet', '상위항목상관계수', 'WriteRowNames', true);
    end

    % 기초 통계 저장
    statsTable = table();
    for i = 1:length(upperScoreNames)
        colData = cleanUpperScoreMatrix(:, i);
        newRow = table();
        newRow.Category = {upperScoreNames{i}};
        newRow.Mean = mean(colData, 'omitnan');
        newRow.Std = std(colData, 'omitnan');
        newRow.Min = min(colData);
        newRow.Max = max(colData);
        newRow.N = sum(~isnan(colData));

        statsTable = [statsTable; newRow];
    end
    writetable(statsTable, outputFileName, 'Sheet', '기초통계');

    fprintf('✓ 분석 결과 저장 완료: %s\n', outputFileName);

    % MAT 파일로도 저장
    matFileName = sprintf('D:\\project\\HR데이터\\결과\\성과종합점수&역검\\upper_category_analysis_%s.mat', dateStr);
    save(matFileName, 'cleanUpperScoreMatrix', 'upperScoreNames', 'cleanCommonIDs', ...
         'upperCorrMatrix', 'matchedUpperData', 'matchedTotalData');

    fprintf('✓ MAT 파일 저장 완료: %s\n', matFileName);

catch ME
    fprintf('✗ 결과 저장 실패: %s\n', ME.message);
end

fprintf('\n========================================\n');
fprintf('상위항목 데이터 구조 분석 완료\n');
fprintf('========================================\n');

fprintf('\n📊 분석 요약:\n');
fprintf('  • 상위항목 수: %d개\n', length(upperScoreNames));
fprintf('  • 분석 대상자: %d명\n', size(cleanUpperScoreMatrix, 1));
fprintf('  • 저장된 파일: %s\n', outputFileName);

if length(upperScoreNames) > 0
    fprintf('\n📋 발견된 상위항목:\n');
    for i = 1:length(upperScoreNames)
        fprintf('  %d. %s\n', i, upperScoreNames{i});
    end
end

fprintf('\n✅ 다음 단계: 성과점수와의 상관분석 및 중다회귀분석 준비 완료\n');

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