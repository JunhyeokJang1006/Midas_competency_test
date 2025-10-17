%% 개선된 표준화 함수 테스트
%
% 목적: 개선된 표준화 함수가 올바르게 작동하는지 테스트
%

clear; clc; close all;
cd('D:\project\HR데이터\matlab')

fprintf('========================================\n');
fprintf('개선된 표준화 함수 테스트\n');
fprintf('========================================\n\n');

%% 1. 기존 데이터 로드
fprintf('[1단계] 기존 분석 결과 로드\n');
fprintf('----------------------------------------\n');

matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    matFileName = matFiles(idx).name;

    try
        loadedData = load(matFileName);
        if isfield(loadedData, 'allData')
            allData = loadedData.allData;
            fprintf('✓ 역량진단 데이터 로드 완료\n');
        else
            fprintf('✗ allData를 찾을 수 없습니다\n');
            return;
        end

        if isfield(loadedData, 'periods')
            periods = loadedData.periods;
        else
            periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
        end

    catch ME
        fprintf('✗ MAT 파일 로드 실패: %s\n', ME.message);
        return;
    end
else
    fprintf('✗ 분석 결과 파일을 찾을 수 없습니다\n');
    return;
end

%% 2. 테스트용 데이터 추출
fprintf('\n[2단계] 테스트용 데이터 추출\n');
fprintf('----------------------------------------\n');

% 첫 번째 period 데이터 사용
periodFields = fieldnames(allData);
testPeriod = periodFields{1};
testData = allData.(testPeriod).selfData;

% ID 컬럼 찾기
idCol = findIDColumn(testData);
if isempty(idCol)
    fprintf('✗ ID 컬럼을 찾을 수 없습니다\n');
    return;
end

% 문항 컬럼들 추출
[questionCols, questionData] = extractQuestionData(testData, idCol);
if isempty(questionCols)
    fprintf('✗ 문항 데이터를 찾을 수 없습니다\n');
    return;
end

fprintf('✓ 테스트 데이터 추출 완료\n');
fprintf('  - 문항 수: %d개\n', length(questionCols));
fprintf('  - 응답자 수: %d명\n', size(questionData, 1));

%% 3. 기존 표준화 vs 개선된 표준화 비교
fprintf('\n[3단계] 기존 vs 개선된 표준화 비교\n');
fprintf('----------------------------------------\n');

% 처음 5개 문항만 테스트
testQuestions = questionCols(1:min(5, length(questionCols)));
testQuestionData = questionData(:, 1:min(5, size(questionData, 2)));

fprintf('테스트 문항: %s\n', strjoin(testQuestions, ', '));

% 기존 표준화 (자동 탐지)
fprintf('\n▶ 기존 표준화 방법:\n');
try
    oldStandardized = standardizeQuestionScales_old(testQuestionData, testQuestions, 1);
    fprintf('✓ 기존 표준화 완료\n');
catch ME
    fprintf('✗ 기존 표준화 실패: %s\n', ME.message);
    oldStandardized = [];
end

% 개선된 표준화
fprintf('\n▶ 개선된 표준화 방법:\n');
try
    newStandardized = standardizeQuestionScales_improved(testQuestionData, testQuestions, 1);
    fprintf('✓ 개선된 표준화 완료\n');
catch ME
    fprintf('✗ 개선된 표준화 실패: %s\n', ME.message);
    newStandardized = [];
end

%% 4. 결과 비교 및 검증
fprintf('\n[4단계] 결과 비교 및 검증\n');
fprintf('----------------------------------------\n');

if ~isempty(oldStandardized) && ~isempty(newStandardized)
    for i = 1:length(testQuestions)
        qName = testQuestions{i};

        % 원본 데이터 통계
        originalData = testQuestionData(:, i);
        validOriginal = originalData(~isnan(originalData));

        % 기존 표준화 결과
        oldData = oldStandardized(:, i);
        validOld = oldData(~isnan(oldData));

        % 개선된 표준화 결과
        newData = newStandardized(:, i);
        validNew = newData(~isnan(newData));

        fprintf('\n%-6s 비교:\n', qName);
        fprintf('  원본 범위: [%.1f, %.1f]\n', min(validOriginal), max(validOriginal));
        fprintf('  기존 표준화: [%.3f, %.3f] (평균: %.3f)\n', ...
            min(validOld), max(validOld), mean(validOld));
        fprintf('  개선 표준화: [%.3f, %.3f] (평균: %.3f)\n', ...
            min(validNew), max(validNew), mean(validNew));

        % 범위 검증
        if any(validNew < -0.001) || any(validNew > 1.001)
            fprintf('  ❌ 개선된 표준화에서 범위 오류 발생!\n');
        else
            fprintf('  ✓ 개선된 표준화 범위 검증 통과\n');
        end

        % 차이 분석
        if length(validOld) == length(validNew)
            maxDiff = max(abs(validOld - validNew));
            fprintf('  최대 차이: %.3f\n', maxDiff);
        end
    end
end

fprintf('\n========================================\n');
fprintf('개선된 표준화 함수 테스트 완료\n');
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

function standardizedData = standardizeQuestionScales_old(questionData, questionNames, periodNum)
    % 기존 방식: 실제 데이터에서 자동 탐지
    standardizedData = questionData;

    fprintf('=== 기존 표준화 방법 (자동 탐지) ===\n');

    for i = 1:size(questionData, 2)
        questionName = questionNames{i};
        columnData = questionData(:, i);
        validData = columnData(~isnan(columnData));

        if isempty(validData)
            continue;
        end

        % 실제 데이터에서 최소/최대값 추정
        actualMin = min(validData);
        actualMax = max(validData);

        % 표준 리커트 척도로 보정
        if all(mod(validData, 1) == 0) && length(unique(validData)) <= 10
            if actualMax <= 4 && actualMin >= 1
                minScale = 1; maxScale = 4;
            elseif actualMax <= 5 && actualMin >= 1
                minScale = 1; maxScale = 5;
            elseif actualMax <= 7 && actualMin >= 1
                minScale = 1; maxScale = 7;
            elseif actualMax <= 10 && actualMin >= 1
                minScale = 1; maxScale = 10;
            else
                minScale = actualMin; maxScale = actualMax;
            end
        else
            minScale = actualMin; maxScale = actualMax;
        end

        fprintf('%-6s: 자동탐지 %.1f~%.1f → %d~%d\n', ...
            questionName, actualMin, actualMax, minScale, maxScale);

        % Min-Max 정규화
        if maxScale > minScale
            standardizedData(:, i) = (columnData - minScale) / (maxScale - minScale);
        else
            standardizedData(:, i) = 0.5 * ones(size(columnData));
            standardizedData(isnan(columnData), i) = NaN;
        end
    end
end

function standardizedData = standardizeQuestionScales_improved(questionData, questionNames, periodNum)
    % 개선된 방식: 사전 정의된 척도 매핑 사용
    standardizedData = questionData;

    fprintf('=== 개선된 표준화 방법 (사전 정의) ===\n');

    % 실제 데이터 분석 기반 척도 매핑
    scaleMapping = containers.Map();
    scaleMapping('Q1') = [1, 4];
    scaleMapping('Q4') = [0, 4];
    scaleMapping('Q5') = [1, 7];
    scaleMapping('Q6') = [1, 7];
    scaleMapping('Q7') = [1, 7];
    scaleMapping('Q8') = [1, 7];
    scaleMapping('Q9') = [1, 7];
    scaleMapping('Q10') = [1, 7];
    scaleMapping('Q11') = [1, 7];
    scaleMapping('Q12') = [1, 7];
    scaleMapping('Q13') = [1, 7];
    scaleMapping('Q14') = [1, 7];
    scaleMapping('Q15') = [1, 7];
    scaleMapping('Q16') = [1, 7];
    scaleMapping('Q17') = [1, 7];
    scaleMapping('Q18') = [1, 7];
    scaleMapping('Q19') = [1, 4];
    scaleMapping('Q20') = [1, 4];
    scaleMapping('Q21') = [1, 7];
    scaleMapping('Q22') = [1, 5];
    scaleMapping('Q23') = [1, 7];
    scaleMapping('Q24') = [1, 7];
    scaleMapping('Q25') = [0, 7];
    scaleMapping('Q26') = [1, 7];
    scaleMapping('Q27') = [0, 4];
    scaleMapping('Q31') = [1, 7];
    scaleMapping('Q32') = [1, 7];
    scaleMapping('Q33') = [1, 7];
    scaleMapping('Q34') = [1, 7];
    scaleMapping('Q35') = [1, 7];
    scaleMapping('Q36') = [1, 7];
    scaleMapping('Q37') = [1, 7];
    scaleMapping('Q38') = [1, 7];
    scaleMapping('Q39') = [1, 7];
    scaleMapping('Q40') = [1, 7];
    scaleMapping('Q41') = [0, 5];
    scaleMapping('Q42') = [1, 4];
    scaleMapping('Q44') = [0, 5];
    scaleMapping('Q45') = [1, 4];
    scaleMapping('Q51') = [0, 4];
    scaleMapping('Q52') = [1, 4];

    for i = 1:size(questionData, 2)
        questionName = questionNames{i};
        columnData = questionData(:, i);
        validData = columnData(~isnan(columnData));

        if isempty(validData)
            continue;
        end

        % 척도 정보 가져오기
        if isKey(scaleMapping, questionName)
            scaleInfo = scaleMapping(questionName);
            minScale = scaleInfo(1);
            maxScale = scaleInfo(2);
            fprintf('%-6s: 사전정의 %d~%d점 척도 사용\n', questionName, minScale, maxScale);
        else
            % 새로운 문항인 경우 자동 탐지
            actualMin = min(validData);
            actualMax = max(validData);

            if all(mod(validData, 1) == 0) && length(unique(validData)) <= 10
                if actualMax <= 4 && actualMin >= 1
                    minScale = 1; maxScale = 4;
                elseif actualMax <= 5 && actualMin >= 1
                    minScale = 1; maxScale = 5;
                elseif actualMax <= 7 && actualMin >= 1
                    minScale = 1; maxScale = 7;
                elseif actualMax <= 10 && actualMin >= 1
                    minScale = 1; maxScale = 10;
                else
                    minScale = actualMin; maxScale = actualMax;
                end
                fprintf('%-6s: 자동탐지 %d~%d점 척도\n', questionName, minScale, maxScale);
            else
                minScale = actualMin; maxScale = actualMax;
                fprintf('%-6s: 연속형 %.2f~%.2f 범위\n', questionName, minScale, maxScale);
            end
        end

        % Min-Max 표준화 적용
        if maxScale > minScale
            standardizedData(:, i) = (columnData - minScale) / (maxScale - minScale);
        else
            standardizedData(:, i) = 0.5 * ones(size(columnData));
            standardizedData(isnan(columnData), i) = NaN;
        end
    end
end