%% CSR 점수와 역량검사 결과 간 상관분석 및 모델링 시스템
%
% 목적: 기업사회책임(CSR) 점수와 개인 역량검사 결과 간의 관계 분석
% 데이터: 25년 상반기 CSR 문항(Q42-46)과 역량검사 상위항목 점수
% 기반: 문항기반 분석가이드 패턴 적용한 개선된 시스템
%
% 주요 개선사항:
% - 문항별 척도 분석 및 표준화 개선
% - Pairwise correlation 및 관대한 결측치 처리
% - 중다회귀분석 추가
% - 종합적인 시각화 시스템
% - 자동 백업 시스템
%
% 작성일: 2025년 9월 16일 (v2.0)
% 기반 코드: corr_item_vs_comp_score.m

clear; clc; close all;

%% 초기 설정
fprintf('========================================\n');
fprintf('📊 CSR vs 역량검사 상관분석 시작\n');
fprintf('========================================\n\n');

% 경로 설정 (corr_item_vs_comp_score.m과 동일)
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
competencyTestPath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
outputDir = 'D:\project\HR데이터\결과\CSR';
resultFileName = fullfile(outputDir, 'CSR_vs_competency_results.xlsx');

% 확장된 조직지표 + CSR 문항 정의 (25년 상반기 기준)
extendedQuestions = {
    'Q40', 'OrganizationalSynergy';    % 조직 시너지
    'Q41', 'Pride';                    % 자부심
    'Q42', 'C1_Communication_Relationship';   % CSR-소통관계
    'Q43', 'C2_Communication_Purpose';        % CSR-소통목적
    'Q44', 'C3_Strategy_CustomerValue';       % CSR-전략고객가치
    'Q45', 'C4_Strategy_Performance';         % CSR-전략성과
    'Q46', 'C5_Reflection_Organizational'     % CSR-성찰조직
};

% 개별 CSR 문항을 위한 리스트 (기존 호환성)
csrQuestions = {'Q42', 'Q43', 'Q44', 'Q45', 'Q46'};

% 문항별 척도 매핑 (실제 데이터 기반으로 확인됨)
extendedScaleMapping = containers.Map();
extendedScaleMapping('Q40') = [0, 4];  % 조직시너지 (실제 확인됨)
extendedScaleMapping('Q41') = [1, 5];  % 자부심 (추정)
extendedScaleMapping('Q42') = [2, 5];  % CSR C1 (실제 확인됨)
extendedScaleMapping('Q43') = [3, 5];  % CSR C2 (실제 확인됨)
extendedScaleMapping('Q44') = [2, 5];  % CSR C3 (실제 확인됨)
extendedScaleMapping('Q45') = [2, 5];  % CSR C4 (실제 확인됨)
extendedScaleMapping('Q46') = [2, 5];  % CSR C5 (실제 확인됨)

%% [1단계] 25년 상반기 데이터 로드
fprintf('[1단계] 25년 상반기 데이터 로드\n');
fprintf('----------------------------------------\n');

try
    % 25년 상반기 역량진단 데이터 로드
    csrDataFile = fullfile(dataPath, '25년_상반기_역량진단_응답데이터.xlsx');

    fprintf('▶ CSR 데이터 로드 중...\n');
    csrRawData = readtable(csrDataFile, 'Sheet', '하향 진단', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ CSR 원시 데이터: %d명\n', height(csrRawData));

    % 역량검사 데이터 로드 (corr_item_vs_comp_score.m과 동일한 방식)
    fprintf('▶ 역량검사 데이터 로드 중...\n');

    % 역량검사 종합점수 로드
    competencyTestData = readtable(competencyTestPath, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 역량검사 종합점수 데이터: %d명\n', height(competencyTestData));

    % 역량검사 상위항목 점수 로드
    upperCategoryData = readtable(competencyTestPath, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
    fprintf('  ✓ 역량검사 상위항목 데이터: %d명\n', height(upperCategoryData));

catch ME
    fprintf('✗ 데이터 로드 실패: %s\n', ME.message);
    return;
end

%% [2단계] 확장된 조직지표 + CSR 데이터 추출
fprintf('\n[2단계] 확장된 조직지표 + CSR 데이터 추출\n');
fprintf('----------------------------------------\n');

% ID 컬럼 찾기
csrIDCol = findIDColumn(csrRawData);
if isempty(csrIDCol)
    fprintf('✗ CSR 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
    return;
end

csrIDs = extractAndStandardizeIDs(csrRawData{:, csrIDCol});

% 확장된 문항 데이터 추출
fprintf('▶ 확장된 문항 검색 결과:\n');
extendedData = table();
extendedData.ID = csrIDs;

foundQuestions = {};
missingQuestions = {};
extendedMatrix = [];
validExtendedQuestions = {};

for i = 1:size(extendedQuestions, 1)
    qCode = extendedQuestions{i, 1};
    qName = extendedQuestions{i, 2};

    % 해당 문항 컬럼 찾기
    questionCol = [];
    colNames = csrRawData.Properties.VariableNames;

    for col = 1:width(csrRawData)
        colName = colNames{col};
        if strcmp(colName, qCode) || contains(colName, qCode)
            questionCol = col;
            break;
        end
    end

    if ~isempty(questionCol)
        qData = csrRawData{:, questionCol};

        % 셀 배열 처리
        if iscell(qData)
            try
                numData = cellfun(@(x) str2double(x), qData);
                if ~all(isnan(numData))
                    qData = numData;
                else
                    fprintf('  ✗ %s (%s) 셀 배열 변환 실패\n', qCode, qName);
                    extendedData.(qName) = NaN(height(extendedData), 1);
                    missingQuestions{end+1} = qCode;
                    continue;
                end
            catch
                fprintf('  ✗ %s (%s) 셀 배열 처리 오류\n', qCode, qName);
                extendedData.(qName) = NaN(height(extendedData), 1);
                missingQuestions{end+1} = qCode;
                continue;
            end
        end

        % 숫자 변환 시도
        if ~isnumeric(qData)
            try
                qData = str2double(qData);
            catch
                fprintf('  ✗ %s (%s) 숫자 변환 실패\n', qCode, qName);
                extendedData.(qName) = NaN(height(extendedData), 1);
                missingQuestions{end+1} = qCode;
                continue;
            end
        end

        % 유효 데이터 확인
        validData = qData(~isnan(qData));
        if length(validData) >= 5  % 최소 5명 이상
            extendedData.(qName) = qData;
            extendedMatrix = [extendedMatrix, qData];
            validExtendedQuestions{end+1} = qName;
            foundQuestions{end+1} = qCode;

            fprintf('  ✓ %s (%s): 범위 [%.1f, %.1f], 유효 %d명\n', qCode, qName, ...
                min(validData), max(validData), length(validData));
        else
            fprintf('  ✗ %s (%s) 유효 데이터 부족 (%d명)\n', qCode, qName, length(validData));
            extendedData.(qName) = NaN(height(extendedData), 1);
            missingQuestions{end+1} = qCode;
        end
    else
        fprintf('  ✗ %s (%s) 누락\n', qCode, qName);
        extendedData.(qName) = NaN(height(extendedData), 1);
        missingQuestions{end+1} = qCode;
    end
end

fprintf('\n▶ 문항 탐지 요약:\n');
fprintf('  발견된 문항: %d개 / %d개\n', length(foundQuestions), size(extendedQuestions, 1));
fprintf('  → 발견: %s\n', strjoin(foundQuestions, ', '));
if ~isempty(missingQuestions)
    fprintf('  → 누락: %s\n', strjoin(missingQuestions, ', '));
end

if isempty(validExtendedQuestions)
    fprintf('✗ 유효한 확장 문항이 없습니다\n');
    return;
end

%% [3단계] 확장된 문항별 척도 분석 및 표준화
fprintf('\n[3단계] 확장된 문항별 척도 분석 및 표준화\n');
fprintf('----------------------------------------\n');

% 실제 척도 분석
fprintf('▶ 문항별 실제 척도 분석:\n');
for q = 1:length(validExtendedQuestions)
    qName = validExtendedQuestions{q};
    qData = extendedMatrix(:, q);
    validData = qData(~isnan(qData));

    if ~isempty(validData)
        actualMin = min(validData);
        actualMax = max(validData);
        actualRange = actualMax - actualMin;
        uniqueValues = length(unique(validData));

        fprintf('  %s: 범위 [%.1f-%.1f], 폭 %.1f, 유니크값 %d개\n', ...
            qName, actualMin, actualMax, actualRange, uniqueValues);
    end
end

% 개선된 표준화 함수 적용
standardizedExtended = standardizeCSRQuestions(extendedMatrix, validExtendedQuestions, extendedScaleMapping);

% 표준화 검증
fprintf('\n▶ 표준화 결과 검증:\n');
for q = 1:size(standardizedExtended, 2)
    qName = validExtendedQuestions{q};
    stdData = standardizedExtended(:, q);
    validStdData = stdData(~isnan(stdData));

    if ~isempty(validStdData)
        minStd = min(validStdData);
        maxStd = max(validStdData);

        if minStd >= 0 && maxStd <= 1
            status = '✓';
        else
            status = '✗';
        end

        fprintf('  %s %s: 표준화 범위 [%.3f-%.3f]\n', status, qName, minStd, maxStd);
    end
end

%% [4단계] 영역별 점수 계산 (확장된 버전)
fprintf('\n[4단계] 영역별 점수 계산\n');
fprintf('----------------------------------------\n');

extendedScores = struct();

% 1. 조직 기반 지표 (Q40, Q41)
orgQuestions = {'OrganizationalSynergy', 'Pride'};
orgIndices = [];
for i = 1:length(orgQuestions)
    idx = find(strcmp(validExtendedQuestions, orgQuestions{i}), 1);
    if ~isempty(idx)
        orgIndices = [orgIndices, idx];
    end
end

if ~isempty(orgIndices)
    extendedScores.Organizational_Score = nanmean(standardizedExtended(:, orgIndices), 2);
    fprintf('  조직 기반 점수 (Q40,41): %d명 유효\n', sum(~isnan(extendedScores.Organizational_Score)));

    % 개별 점수도 저장
    for i = 1:length(orgIndices)
        qName = validExtendedQuestions{orgIndices(i)};
        extendedScores.(qName) = standardizedExtended(:, orgIndices(i));
    end
end

% 2. CSR 개별 문항 점수 (C1-C5)
csrQuestionNames = {'C1_Communication_Relationship', 'C2_Communication_Purpose', ...
                    'C3_Strategy_CustomerValue', 'C4_Strategy_Performance', ...
                    'C5_Reflection_Organizational'};

csrIndices = [];
for i = 1:length(csrQuestionNames)
    idx = find(strcmp(validExtendedQuestions, csrQuestionNames{i}), 1);
    if ~isempty(idx)
        csrIndices = [csrIndices, idx];
        % 개별 CSR 문항 점수 저장
        extendedScores.(csrQuestionNames{i}) = standardizedExtended(:, idx);
        fprintf('  %s: %d명 유효\n', csrQuestionNames{i}, sum(~isnan(standardizedExtended(:, idx))));
    end
end

% 3. CSR 카테고리별 점수
% Communication Score (C1, C2)
commIndices = [];
for qName = {'C1_Communication_Relationship', 'C2_Communication_Purpose'}
    idx = find(strcmp(validExtendedQuestions, qName{1}), 1);
    if ~isempty(idx)
        commIndices = [commIndices, idx];
    end
end
if ~isempty(commIndices)
    extendedScores.Communication_Score = nanmean(standardizedExtended(:, commIndices), 2);
    fprintf('  Communication 점수 (C1,C2): %d명 유효\n', sum(~isnan(extendedScores.Communication_Score)));
end

% Strategy Score (C3, C4)
stratIndices = [];
for qName = {'C3_Strategy_CustomerValue', 'C4_Strategy_Performance'}
    idx = find(strcmp(validExtendedQuestions, qName{1}), 1);
    if ~isempty(idx)
        stratIndices = [stratIndices, idx];
    end
end
if ~isempty(stratIndices)
    extendedScores.Strategy_Score = nanmean(standardizedExtended(:, stratIndices), 2);
    fprintf('  Strategy 점수 (C3,C4): %d명 유효\n', sum(~isnan(extendedScores.Strategy_Score)));
end

% Reflection Score (C5)
reflIdx = find(strcmp(validExtendedQuestions, 'C5_Reflection_Organizational'), 1);
if ~isempty(reflIdx)
    extendedScores.Reflection_Score = standardizedExtended(:, reflIdx);
    fprintf('  Reflection 점수 (C5): %d명 유효\n', sum(~isnan(extendedScores.Reflection_Score)));
end

% 4. 전체 CSR 점수 (C1-C5)
if ~isempty(csrIndices)
    extendedScores.Total_CSR_Score = nanmean(standardizedExtended(:, csrIndices), 2);
    fprintf('  전체 CSR 점수 (C1-C5): %d명 유효\n', sum(~isnan(extendedScores.Total_CSR_Score)));
end

% 5. 통합 리더십 점수 (Q40-Q46 전체)
if ~isempty(orgIndices) && ~isempty(csrIndices)
    allIndices = [orgIndices, csrIndices];
    extendedScores.Total_Leadership_Score = nanmean(standardizedExtended(:, allIndices), 2);
    fprintf('  통합 리더십 점수 (Q40-Q46): %d명 유효\n', sum(~isnan(extendedScores.Total_Leadership_Score)));
end

% 결과 요약
extendedCategories = fieldnames(extendedScores);
fprintf('\n▶ 계산된 점수 카테고리: %d개\n', length(extendedCategories));
for i = 1:length(extendedCategories)
    catName = extendedCategories{i};
    scores = extendedScores.(catName);
    validCount = sum(~isnan(scores));

    if validCount > 0
        fprintf('  %s: 평균 %.3f (±%.3f), 유효 %d명\n', catName, ...
            nanmean(scores), nanstd(scores), validCount);
    end
end

% 기존 호환성을 위한 별칭 생성
csrScores = extendedScores;
csrCategories = extendedCategories;

%% [5단계] 역량검사 상위항목 데이터 전처리
fprintf('\n[5단계] 역량검사 상위항목 데이터 전처리\n');
fprintf('----------------------------------------\n');

% 역량검사 ID 컬럼 찾기
compIDCol = findIDColumn(upperCategoryData);
if isempty(compIDCol)
    fprintf('✗ 역량검사 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
    return;
end

competencyIDs = extractAndStandardizeIDs(upperCategoryData{:, compIDCol});

% 상위항목 점수 컬럼 동적 탐지 (문항기반 분석가이드 패턴 적용)
fprintf('▶ 상위항목 점수 컬럼 동적 탐지 중...\n');

colNames = upperCategoryData.Properties.VariableNames;
competencyMatrix = [];
validCompetencyCategories = {};
competencyColumnIndices = [];

for i = 1:width(upperCategoryData)
    if i == compIDCol  % ID 컬럼 제외
        continue;
    end

    colName = colNames{i};
    colData = upperCategoryData{:, i};

    % 숫자형 데이터이고, 결측치가 아닌 값이 있으며, 분산이 0보다 큰 컬럼만 선택
    if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
        competencyMatrix = [competencyMatrix, colData];
        validCompetencyCategories{end+1} = colName;
        competencyColumnIndices = [competencyColumnIndices, i];

        % 통계 정보 출력
        validData = colData(~isnan(colData));
        fprintf('  %s: 평균 %.3f (±%.3f), 범위 [%.1f, %.1f], 유효 %d명\n', colName, ...
            nanmean(colData), nanstd(colData), min(validData), max(validData), length(validData));
    else
        % 제외된 컬럼에 대한 정보 (디버깅용)
        if ~isnumeric(colData)
            fprintf('  건너뜀: %s (숫자형 아님)\n', colName);
        elseif all(isnan(colData))
            fprintf('  건너뜀: %s (모든 값이 결측치)\n', colName);
        elseif var(colData, 'omitnan') == 0
            fprintf('  건너뜀: %s (분산이 0)\n', colName);
        end
    end
end

if isempty(validCompetencyCategories)
    fprintf('✗ 유효한 역량검사 카테고리가 없습니다\n');
    fprintf('  → 사용 가능한 컬럼들: %s\n', strjoin(colNames, ', '));
    return;
end

fprintf('  ✓ 발견된 상위항목: %d개\n', length(validCompetencyCategories));
fprintf('  → 상위항목 리스트: %s\n', strjoin(validCompetencyCategories, ', '));

%% [6단계] ID 매칭 및 데이터 정리
fprintf('\n[6단계] ID 매칭 및 데이터 정리\n');
fprintf('----------------------------------------\n');

% ID 매칭
[commonIDs, csrIdx, compIdx] = intersect(csrIDs, competencyIDs);

if length(commonIDs) < 5
    fprintf('✗ 매칭된 ID가 부족합니다 (%d개)\n', length(commonIDs));
    return;
end

fprintf('✓ ID 매칭 성공: %d명 (CSR: %d명, 역량검사: %d명)\n', ...
    length(commonIDs), length(csrIDs), length(competencyIDs));

% 매칭된 데이터 추출
matchedCSRData = struct();
for i = 1:length(csrCategories)
    catName = csrCategories{i};
    matchedCSRData.(catName) = csrScores.(catName)(csrIdx);
end

matchedCompetencyData = competencyMatrix(compIdx, :);

% 개선된 결측치 처리 (Pairwise + 관대한 기준)
fprintf('\n▶ 결측치 처리 방식:\n');

% 1) Listwise deletion을 위한 완전한 데이터 행 찾기
validRowsComplete = [];
for i = 1:length(commonIDs)
    csrValid = true;
    for j = 1:length(csrCategories)
        catName = csrCategories{j};
        if isnan(matchedCSRData.(catName)(i))
            csrValid = false;
            break;
        end
    end

    compValid = ~any(isnan(matchedCompetencyData(i, :)));

    if csrValid && compValid
        validRowsComplete = [validRowsComplete; i];
    end
end

% 2) 관대한 기준을 위한 부분 완전 데이터 행 찾기
missingCount = sum(isnan(matchedCompetencyData), 2);
maxMissing = floor(size(matchedCompetencyData, 2) * 0.3); % 30%까지 결측 허용
validRowsPartial = find(missingCount <= maxMissing);

% CSR 데이터도 각 카테고리별로 유효성 확인
csrValidRows = [];
for i = 1:length(commonIDs)
    csrMissingCount = 0;
    for j = 1:length(csrCategories)
        catName = csrCategories{j};
        if isnan(matchedCSRData.(catName)(i))
            csrMissingCount = csrMissingCount + 1;
        end
    end

    if csrMissingCount <= floor(length(csrCategories) * 0.3) % CSR도 30%까지 허용
        csrValidRows = [csrValidRows; i];
    end
end

% 최종 유효 행 결정
validRowsPartial = intersect(validRowsPartial, csrValidRows);

fprintf('  완전한 데이터 (Listwise): %d명\n', length(validRowsComplete));
fprintf('  부분 완전 데이터 (최대 30%% 결측): %d명\n', length(validRowsPartial));

% 분석에 사용할 행 결정 (완전한 데이터가 충분하면 사용, 아니면 부분 완전 사용)
if length(validRowsComplete) >= 20
    validRows = validRowsComplete;
    analysisType = 'Listwise (완전한 데이터)';
else
    validRows = validRowsPartial;
    analysisType = 'Pairwise (부분 완전 데이터)';
end

if length(validRows) < 10
    fprintf('✗ 분석에 필요한 최소 사례가 부족합니다 (%d개)\n', length(validRows));
    return;
end

fprintf('  최종 분석 방식: %s (%d명)\n', analysisType, length(validRows));

% 최종 분석 데이터
finalCSRData = struct();
for i = 1:length(csrCategories)
    catName = csrCategories{i};
    finalCSRData.(catName) = matchedCSRData.(catName)(validRows);
end

finalCompetencyData = matchedCompetencyData(validRows, :);
finalIDs = commonIDs(validRows);

fprintf('✓ 최종 분석 대상: %d명\n', length(validRows));

%% [7단계] CSR vs 역량검사 상관분석 (Pairwise correlation)
fprintf('\n[7단계] CSR vs 역량검사 상관분석\n');
fprintf('========================================\n');

correlationResults = struct();

for i = 1:length(csrCategories)
    csrCatName = csrCategories{i};
    csrData = finalCSRData.(csrCatName);

    fprintf('\n▶ %s vs 역량검사 상관분석:\n', csrCatName);

    categoryCorrelations = [];
    categoryPValues = [];
    sampleSizes = [];  % 각 분석의 실제 N값 저장

    for j = 1:length(validCompetencyCategories)
        compCatName = validCompetencyCategories{j};
        compData = finalCompetencyData(:, j);

        % Pairwise correlation (결측치 자동 제외)
        [r, p] = corr(csrData, compData, 'rows', 'pairwise');

        % 실제 사용된 샘플 크기 계산
        validPairs = ~isnan(csrData) & ~isnan(compData);
        actualN = sum(validPairs);

        categoryCorrelations = [categoryCorrelations; r];
        categoryPValues = [categoryPValues; p];
        sampleSizes = [sampleSizes; actualN];

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

        fprintf('  %s: r=%.3f (p=%.3f) %s [N=%d]\n', compCatName, r, p, sig, actualN);
    end

    % 결과 저장
    correlationResults.(csrCatName) = struct();
    correlationResults.(csrCatName).correlations = categoryCorrelations;
    correlationResults.(csrCatName).pValues = categoryPValues;
    correlationResults.(csrCatName).sampleSizes = sampleSizes;  % 샘플 크기 추가
    correlationResults.(csrCatName).competencyCategories = validCompetencyCategories;
end

%% [8단계] 유의한 상관관계 요약
fprintf('\n[8단계] 유의한 상관관계 요약\n');
fprintf('========================================\n');

allSignificantResults = {};
marginalResults = {};  % 0.05 < p < 0.1
strongCorrelations = {}; % |r| > 0.1 & p < 0.1 (유의 + marginal 포함)

for i = 1:length(csrCategories)
    csrCatName = csrCategories{i};
    correlations = correlationResults.(csrCatName).correlations;
    pValues = correlationResults.(csrCatName).pValues;

    % 유의한 상관관계 찾기 (p < 0.05)
    significantIdx = find(pValues < 0.05);
    % marginal 상관관계 찾기 (0.05 ≤ p < 0.1)
    marginalIdx = find(pValues >= 0.05 & pValues < 0.1);
    % 강한 상관관계 찾기 (|r| > 0.1 & p < 0.1)
    strongIdx = find(abs(correlations) > 0.1 & pValues < 0.1);

    if ~isempty(significantIdx)
        fprintf('\n▶ %s - 유의한 상관관계 (%d개):\n', csrCatName, length(significantIdx));

        for j = 1:length(significantIdx)
            idx = significantIdx(j);
            compCat = validCompetencyCategories{idx};
            r = correlations(idx);
            p = pValues(idx);

            if p < 0.001
                sig = '***';
            elseif p < 0.01
                sig = '**';
            elseif p < 0.05
                sig = '*';
            else
                sig = '';
            end

            fprintf('  %s ↔ %s: r=%.3f (p=%.3f) %s\n', csrCatName, compCat, r, p, sig);

            % 전체 결과에 추가
            allSignificantResults{end+1} = {csrCatName, compCat, r, p, sig, length(validRows)};

            % 강한 상관관계인 경우 별도 저장
            if abs(r) > 0.1
                strongCorrelations{end+1} = {csrCatName, compCat, r, p, sig, length(validRows)};
            end
        end
    end

    % marginal 상관관계 처리
    if ~isempty(marginalIdx)
        fprintf('\n▶ %s - Marginal 상관관계 (%d개):\n', csrCatName, length(marginalIdx));

        for j = 1:length(marginalIdx)
            idx = marginalIdx(j);
            compCat = validCompetencyCategories{idx};
            r = correlations(idx);
            p = pValues(idx);

            % marginal significance 표시
            if p < 0.1
                sig = '†';  % marginal significance 기호
            else
                sig = '';
            end

            fprintf('  %s ↔ %s: r=%.3f (p=%.3f) %s\n', csrCatName, compCat, r, p, sig);

            % marginal 결과에 추가
            marginalResults{end+1} = {csrCatName, compCat, r, p, sig, length(validRows)};

            % 강한 marginal 상관관계인 경우도 별도 저장
            if abs(r) > 0.1
                strongCorrelations{end+1} = {csrCatName, compCat, r, p, sig, length(validRows)};
            end
        end
    end

    % 유의하지도 marginal하지도 않은 경우
    if isempty(significantIdx) && isempty(marginalIdx)
        fprintf('\n▶ %s - 유의한 상관관계 없음\n', csrCatName);
    end
end

% 강한 상관관계 하이라이트 (유의 + marginal 포함)
if ~isempty(strongCorrelations)
    fprintf('\n🏆 강한 상관관계 (|r| > 0.1, p < 0.1):\n');
    for i = 1:length(strongCorrelations)
        result = strongCorrelations{i};
        fprintf('%d. %s ↔ %s: r=%.3f %s\n', i, result{1}, result{2}, result{3}, result{5});
    end
else
    fprintf('\n⚠ 강한 상관관계 (|r| > 0.1, p < 0.1)가 발견되지 않았습니다.\n');
end

% marginal 결과 요약
if ~isempty(marginalResults)
    fprintf('\n📊 Marginal 상관관계 요약 (0.05 ≤ p < 0.1):\n');
    for i = 1:length(marginalResults)
        result = marginalResults{i};
        fprintf('%d. %s ↔ %s: r=%.3f (p=%.3f) †\n', i, result{1}, result{2}, result{3}, result{4});
    end
end

%% [9단계] 중다회귀분석 (Total CSR 예측 모델)
fprintf('\n[9단계] 중다회귀분석\n');
fprintf('========================================\n');

% Total CSR 점수를 종속변수로 하는 중다회귀분석
if ismember('Total_CSR', csrCategories)
    fprintf('▶ Total CSR 점수 예측 모델 구축\n');

    % 종속변수
    dependentVar = finalCSRData.Total_CSR;

    % 독립변수 (역량검사 상위항목들)
    independentVars = finalCompetencyData;

    % 결측치 처리 (관대한 기준)
    missingCount = sum(isnan(independentVars), 2);
    maxMissing = floor(size(independentVars, 2) * 0.3);
    regressionValidRows = missingCount <= maxMissing & ~isnan(dependentVar);

    if sum(regressionValidRows) >= 15  % 최소 15명 이상
        Y = dependentVar(regressionValidRows);
        X = independentVars(regressionValidRows, :);

        % 결측치를 평균으로 대체
        for col = 1:size(X, 2)
            missingIdx = isnan(X(:, col));
            if any(missingIdx)
                X(missingIdx, col) = nanmean(X(:, col));
            end
        end

        % 중다회귀분석 수행
        [b, bint, r, rint, stats] = regress(Y, [ones(size(X,1), 1), X]);

        % 결과 저장
        regressionResults = struct();
        regressionResults.coefficients = b;
        regressionResults.coefficientIntervals = bint;
        regressionResults.residuals = r;
        regressionResults.R2 = stats(1);
        regressionResults.F = stats(2);
        regressionResults.pValue = stats(3);
        regressionResults.errorVariance = stats(4);
        regressionResults.sampleSize = length(Y);
        regressionResults.predictorNames = ['Intercept'; validCompetencyCategories'];

        % 예측값 계산
        Y_pred = [ones(size(X,1), 1), X] * b;
        regressionResults.predicted = Y_pred;
        regressionResults.actual = Y;

        % 성능 지표 계산
        mae = mean(abs(Y - Y_pred));
        rmse = sqrt(mean((Y - Y_pred).^2));
        regressionResults.MAE = mae;
        regressionResults.RMSE = rmse;

        % 결과 출력
        fprintf('\n  📊 회귀모델 성능:\n');
        fprintf('    R² = %.3f (설명변량 %.1f%%)\n', stats(1), stats(1)*100);
        fprintf('    F(%.0f,%.0f) = %.2f, p = %.3f\n', size(X,2), length(Y)-size(X,2)-1, stats(2), stats(3));
        fprintf('    MAE = %.3f, RMSE = %.3f\n', mae, rmse);
        fprintf('    분석 대상: %d명\n', length(Y));

        if stats(3) < 0.05
            fprintf('    ✓ 통계적으로 유의한 모델\n');
        else
            fprintf('    ⚠ 통계적으로 유의하지 않은 모델\n');
        end

        % 유의한 예측변수 찾기
        significantPredictors = {};
        for i = 2:length(b)  % intercept 제외
            tStat = b(i) / sqrt(stats(4) * ((bint(i,2) - bint(i,1)) / (2 * tinv(0.975, length(Y)-size(X,2)-1)))^(-2));
            pValue = 2 * (1 - tcdf(abs(tStat), length(Y)-size(X,2)-1));

            if pValue < 0.05
                significantPredictors{end+1} = validCompetencyCategories{i-1};
                fprintf('    • %s: β=%.3f (p=%.3f) *\n', validCompetencyCategories{i-1}, b(i), pValue);
            end
        end

        if isempty(significantPredictors)
            fprintf('    ⚠ 유의한 예측변수가 없음\n');
        end

    else
        fprintf('  ✗ 회귀분석에 필요한 최소 사례 부족 (%d명)\n', sum(regressionValidRows));
        regressionResults = [];
    end
else
    fprintf('  ⚠ Total_CSR 카테고리를 찾을 수 없어 회귀분석 생략\n');
    regressionResults = [];
end

%% [10단계] 시각화 생성
fprintf('\n[10단계] 시각화 생성\n');
fprintf('========================================\n');

try
    % 1. CSR 점수 분포 히스토그램
    figure('Name', 'CSR 점수 분포', 'Position', [100, 100, 1200, 800]);

    numPlots = length(csrCategories);
    numCols = ceil(sqrt(numPlots));
    numRows = ceil(numPlots / numCols);

    for i = 1:length(csrCategories)
        subplot(numRows, numCols, i);
        csrCatName = csrCategories{i};
        csrData = finalCSRData.(csrCatName);

        histogram(csrData, 20, 'FaceColor', [0.7, 0.7, 0.9], 'EdgeColor', 'black');
        title(sprintf('%s 분포', csrCatName), 'FontSize', 12, 'Interpreter', 'none');
        xlabel('점수', 'FontSize', 10);
        ylabel('빈도', 'FontSize', 10);
        grid on;

        % 통계 정보 추가
        meanVal = nanmean(csrData);
        stdVal = nanstd(csrData);
        text(0.7, 0.9, sprintf('평균: %.3f\n표준편차: %.3f', meanVal, stdVal), ...
             'Units', 'normalized', 'FontSize', 9, 'BackgroundColor', 'white');
    end

    sgtitle('CSR 카테고리별 점수 분포', 'FontSize', 14, 'FontWeight', 'bold');

    % 저장
    figFileName = fullfile(outputDir, 'CSR_score_distributions.png');
    saveas(gcf, figFileName);
    fprintf('  ✓ CSR 점수 분포 히스토그램 저장: %s\n', figFileName);

    % 2. 상관계수 히트맵
    if ~isempty(correlationResults) && length(csrCategories) > 1
        figure('Name', 'CSR-역량검사 상관계수 히트맵', 'Position', [150, 150, 1000, 600]);

        % 상관계수 매트릭스 구성
        corrMatrix = [];
        for i = 1:length(csrCategories)
            csrCatName = csrCategories{i};
            corrMatrix = [corrMatrix; correlationResults.(csrCatName).correlations'];
        end

        % 히트맵 생성
        imagesc(corrMatrix);
        colormap(jet);  % 안전한 기본 colormap 사용
        colorbar;
        caxis([-1, 1]);

        % 축 레이블 설정
        set(gca, 'XTick', 1:length(validCompetencyCategories), 'XTickLabel', validCompetencyCategories, ...
                 'YTick', 1:length(csrCategories), 'YTickLabel', csrCategories);
        xtickangle(45);

        title('CSR 카테고리와 역량검사 간 상관계수', 'FontSize', 14, 'FontWeight', 'bold');
        xlabel('역량검사 카테고리', 'FontSize', 12);
        ylabel('CSR 카테고리', 'FontSize', 12);

        % 상관계수 값 텍스트 표시
        for i = 1:size(corrMatrix, 1)
            for j = 1:size(corrMatrix, 2)
                text(j, i, sprintf('%.2f', corrMatrix(i,j)), ...
                     'HorizontalAlignment', 'center', 'FontSize', 8);
            end
        end

        figFileName = fullfile(outputDir, 'CSR_competency_correlation_heatmap.png');
        saveas(gcf, figFileName);
        fprintf('  ✓ 상관계수 히트맵 저장: %s\n', figFileName);
    end

    % 3. 회귀분석 결과 시각화 (Total CSR)
    if ~isempty(regressionResults)
        figure('Name', 'Total CSR 예측 모델', 'Position', [200, 200, 1200, 400]);

        % 실제값 vs 예측값 산점도
        subplot(1, 3, 1);
        scatter(regressionResults.actual, regressionResults.predicted, 50, 'filled', 'Alpha', 0.6);
        hold on;
        minVal = min([regressionResults.actual; regressionResults.predicted]);
        maxVal = max([regressionResults.actual; regressionResults.predicted]);
        plot([minVal, maxVal], [minVal, maxVal], 'r--', 'LineWidth', 2);
        xlabel('실제 Total CSR 점수');
        ylabel('예측 Total CSR 점수');
        title(sprintf('실제값 vs 예측값\n(R² = %.3f)', regressionResults.R2));
        grid on;

        % 잔차 플롯
        subplot(1, 3, 2);
        scatter(regressionResults.predicted, regressionResults.residuals, 50, 'filled', 'Alpha', 0.6);
        xlabel('예측값');
        ylabel('잔차');
        title('잔차 플롯');
        yline(0, 'r--', 'LineWidth', 2);
        grid on;

        % 계수 플롯 (intercept 제외)
        subplot(1, 3, 3);
        coeffs = regressionResults.coefficients(2:end);  % intercept 제외
        bar(coeffs);
        set(gca, 'XTick', 1:length(validCompetencyCategories), 'XTickLabel', validCompetencyCategories);
        xtickangle(45);
        xlabel('역량검사 카테고리');
        ylabel('회귀계수');
        title('회귀계수');
        grid on;

        sgtitle('Total CSR 예측 모델 분석 결과', 'FontSize', 14, 'FontWeight', 'bold');

        figFileName = fullfile(outputDir, 'CSR_regression_analysis.png');
        saveas(gcf, figFileName);
        fprintf('  ✓ 회귀분석 결과 시각화 저장: %s\n', figFileName);
    end

    fprintf('  ✓ 모든 시각화 완료\n');

catch ME
    fprintf('  ⚠ 시각화 생성 중 오류: %s\n', ME.message);
end

%% [11단계] 결과 저장 - 완전 개선 버전
fprintf('\n[11단계] 결과 저장\n');
fprintf('========================================\n');

% 출력 디렉토리 확인 및 생성
if ~exist(outputDir, 'dir')
    mkdir(outputDir);
    fprintf('✓ 출력 디렉토리 생성: %s\n', outputDir);
end

% 백업 디렉토리 설정
timestamp = datestr(now, 'yyyymmdd_HHMM');
backupDir = fullfile(outputDir, 'backup');
if ~exist(backupDir, 'dir')
    mkdir(backupDir);
end

% 기존 파일 백업
if exist(resultFileName, 'file')
    [~,name,ext] = fileparts(resultFileName);
    backupFileName = fullfile(backupDir, sprintf('%s_%s%s', name, timestamp, ext));
    try
        copyfile(resultFileName, backupFileName);
        fprintf('✓ 기존 파일 백업 완료: %s\n', backupFileName);
    catch
        fprintf('⚠ 백업 실패, 계속 진행합니다\n');
    end
end

% 저장 상태 추적
saveStatus = struct();
saveStatus.sheets = {};
saveStatus.success = [];
saveStatus.failed = {};

%% 11-1. 개별 CSR 카테고리 상관분석 결과 저장
fprintf('\n▶ [11-1] CSR 카테고리별 상관분석 결과 저장\n');
fprintf('----------------------------------------\n');

% 시트명 매핑 테이블 생성
sheetNameMapping = containers.Map();
sheetNameMapping('OrganizationalSynergy') = 'OrgSynergy';
sheetNameMapping('Pride') = 'Pride';
sheetNameMapping('C1_Communication_Relationship') = 'C1_CommRel';
sheetNameMapping('C2_Communication_Purpose') = 'C2_CommPur';
sheetNameMapping('C3_Strategy_CustomerValue') = 'C3_StratCust';
sheetNameMapping('C4_Strategy_Performance') = 'C4_StratPerf';
sheetNameMapping('C5_Reflection_Organizational') = 'C5_ReflOrg';
sheetNameMapping('Communication_Score') = 'Comm_Score';
sheetNameMapping('Strategy_Score') = 'Strat_Score';
sheetNameMapping('Reflection_Score') = 'Refl_Score';
sheetNameMapping('Organizational_Score') = 'Org_Score';
sheetNameMapping('Total_CSR_Score') = 'Total_CSR';
sheetNameMapping('Total_Leadership_Score') = 'Total_Leader';

% 각 CSR 카테고리별로 저장
for i = 1:length(csrCategories)
    csrCatName = csrCategories{i};
    
    % 안전한 시트명 생성
    if isKey(sheetNameMapping, csrCatName)
        sheetName = sheetNameMapping(csrCatName);
    else
        % 매핑이 없는 경우 안전한 이름 생성
        sheetName = sprintf('CSR_%02d', i);
    end
    
    % 31자 제한 확인
    if length(sheetName) > 31
        sheetName = sheetName(1:31);
    end
    
    try
        % 데이터 준비
        if isfield(correlationResults, csrCatName)
            corrData = correlationResults.(csrCatName).correlations;
            pvalData = correlationResults.(csrCatName).pValues;
            sampleData = correlationResults.(csrCatName).sampleSizes;
            compCategories = correlationResults.(csrCatName).competencyCategories;
            
            % 테이블 생성
            resultTable = table();
            resultTable.Competency = compCategories(:);
            resultTable.Correlation = corrData(:);
            resultTable.PValue = pvalData(:);
            resultTable.SampleSize = sampleData(:);
            
            % 유의성 표시 추가
            sigMarks = cell(length(pvalData), 1);
            for j = 1:length(pvalData)
                if pvalData(j) < 0.001
                    sigMarks{j} = '***';
                elseif pvalData(j) < 0.01
                    sigMarks{j} = '**';
                elseif pvalData(j) < 0.05
                    sigMarks{j} = '*';
                elseif pvalData(j) < 0.1
                    sigMarks{j} = '†';
                else
                    sigMarks{j} = '';
                end
            end
            resultTable.Significance = sigMarks;
            
            % Excel에 저장
            writetable(resultTable, resultFileName, 'Sheet', sheetName);
            
            saveStatus.sheets{end+1} = sheetName;
            saveStatus.success(end+1) = true;
            fprintf('  ✓ %s → "%s" 시트 저장 완료\n', csrCatName, sheetName);
        else
            fprintf('  ⚠ %s 데이터 없음, 건너뜀\n', csrCatName);
        end
        
    catch ME
        saveStatus.failed{end+1} = {csrCatName, ME.message};
        fprintf('  ✗ %s 저장 실패: %s\n', csrCatName, ME.message);
        
        % 실패 시 CSV 백업 시도
        try
            csvName = fullfile(outputDir, sprintf('%s_%s.csv', sheetName, timestamp));
            writetable(resultTable, csvName);
            fprintf('    → CSV 백업 성공: %s\n', csvName);
        catch
            fprintf('    → CSV 백업도 실패\n');
        end
    end
end

%% 11-2. 요약 테이블 저장
fprintf('\n▶ [11-2] 요약 통계 저장\n');
fprintf('----------------------------------------\n');

try
    summaryTable = createCSRSummaryTable(correlationResults, csrCategories, ...
                                         finalCSRData, validRows, validCompetencyCategories);
    
    writetable(summaryTable, resultFileName, 'Sheet', 'Summary');
    saveStatus.sheets{end+1} = 'Summary';
    fprintf('  ✓ 요약 통계 시트 저장 완료\n');
    
catch ME
    fprintf('  ✗ 요약 통계 저장 실패: %s\n', ME.message);
end

%% 11-3. 유의한 상관관계 저장
fprintf('\n▶ [11-3] 유의한 상관관계 저장\n');
fprintf('----------------------------------------\n');

% 유의한 상관관계 (p < 0.05)
if ~isempty(allSignificantResults)
    try
        sigTable = table();
        sigTable.CSR_Category = allSignificantResults(:, 1);
        sigTable.Competency_Category = allSignificantResults(:, 2);
        sigTable.Correlation = cell2mat(allSignificantResults(:, 3));
        sigTable.P_Value = cell2mat(allSignificantResults(:, 4));
        sigTable.Significance = allSignificantResults(:, 5);
        sigTable.Sample_Size = cell2mat(allSignificantResults(:, 6));
        
        % 상관계수 크기순으로 정렬
        [~, sortIdx] = sort(abs(sigTable.Correlation), 'descend');
        sigTable = sigTable(sortIdx, :);
        
        writetable(sigTable, resultFileName, 'Sheet', 'Significant');
        saveStatus.sheets{end+1} = 'Significant';
        fprintf('  ✓ 유의한 상관관계 시트 저장 완료 (%d개)\n', height(sigTable));
        
    catch ME
        fprintf('  ✗ 유의한 상관관계 저장 실패: %s\n', ME.message);
    end
end

% Marginal 상관관계 (0.05 ≤ p < 0.1)
if ~isempty(marginalResults)
    try
        margTable = table();
        margTable.CSR_Category = marginalResults(:, 1);
        margTable.Competency_Category = marginalResults(:, 2);
        margTable.Correlation = cell2mat(marginalResults(:, 3));
        margTable.P_Value = cell2mat(marginalResults(:, 4));
        margTable.Significance = marginalResults(:, 5);
        margTable.Sample_Size = cell2mat(marginalResults(:, 6));
        
        writetable(margTable, resultFileName, 'Sheet', 'Marginal');
        saveStatus.sheets{end+1} = 'Marginal';
        fprintf('  ✓ Marginal 상관관계 시트 저장 완료 (%d개)\n', height(margTable));
        
    catch ME
        fprintf('  ✗ Marginal 상관관계 저장 실패: %s\n', ME.message);
    end
end

% 강한 상관관계 (|r| > 0.1, p < 0.1)
if ~isempty(strongCorrelations)
    try
        strongTable = table();
        strongTable.CSR_Category = strongCorrelations(:, 1);
        strongTable.Competency_Category = strongCorrelations(:, 2);
        strongTable.Correlation = cell2mat(strongCorrelations(:, 3));
        strongTable.P_Value = cell2mat(strongCorrelations(:, 4));
        strongTable.Significance = strongCorrelations(:, 5);
        strongTable.Sample_Size = cell2mat(strongCorrelations(:, 6));
        
        writetable(strongTable, resultFileName, 'Sheet', 'Strong');
        saveStatus.sheets{end+1} = 'Strong';
        fprintf('  ✓ 강한 상관관계 시트 저장 완료 (%d개)\n', height(strongTable));
        
    catch ME
        fprintf('  ✗ 강한 상관관계 저장 실패: %s\n', ME.message);
    end
end

%% 11-4. 회귀분석 결과 저장 (있는 경우)
if exist('regressionResults', 'var') && ~isempty(regressionResults)
    fprintf('\n▶ [11-4] 회귀분석 결과 저장\n');
    fprintf('----------------------------------------\n');
    
    try
        % 회귀계수 테이블
        regCoeffTable = table();
        regCoeffTable.Predictor = regressionResults.predictorNames;
        regCoeffTable.Coefficient = regressionResults.coefficients;
        regCoeffTable.CI_Lower = regressionResults.coefficientIntervals(:,1);
        regCoeffTable.CI_Upper = regressionResults.coefficientIntervals(:,2);
        
        writetable(regCoeffTable, resultFileName, 'Sheet', 'Regression_Coef');
        saveStatus.sheets{end+1} = 'Regression_Coef';
        fprintf('  ✓ 회귀계수 시트 저장 완료\n');
        
        % 모델 성능 요약
        modelSummary = table();
        modelSummary.R_Squared = regressionResults.R2;
        modelSummary.F_Statistic = regressionResults.F;
        modelSummary.P_Value = regressionResults.pValue;
        modelSummary.MAE = regressionResults.MAE;
        modelSummary.RMSE = regressionResults.RMSE;
        modelSummary.Sample_Size = regressionResults.sampleSize;
        
        writetable(modelSummary, resultFileName, 'Sheet', 'Model_Performance');
        saveStatus.sheets{end+1} = 'Model_Performance';
        fprintf('  ✓ 모델 성능 시트 저장 완료\n');
        
    catch ME
        fprintf('  ✗ 회귀분석 결과 저장 실패: %s\n', ME.message);
    end
end

%% 11-5. 상관계수 매트릭스 저장
fprintf('\n▶ [11-5] 상관계수 매트릭스 저장\n');
fprintf('----------------------------------------\n');

try
    % 전체 상관계수 매트릭스 구성
    corrMatrix = [];
    rowNames = {};
    
    for i = 1:length(csrCategories)
        csrCatName = csrCategories{i};
        if isfield(correlationResults, csrCatName)
            corrMatrix = [corrMatrix; correlationResults.(csrCatName).correlations'];
            rowNames{end+1} = csrCatName;
        end
    end
    
    if ~isempty(corrMatrix)
        % 매트릭스를 테이블로 변환
        matrixTable = array2table(corrMatrix, ...
            'VariableNames', validCompetencyCategories, ...
            'RowNames', rowNames);
        
        writetable(matrixTable, resultFileName, 'Sheet', 'Corr_Matrix', ...
                   'WriteRowNames', true);
        saveStatus.sheets{end+1} = 'Corr_Matrix';
        fprintf('  ✓ 상관계수 매트릭스 저장 완료 (%dx%d)\n', ...
                size(corrMatrix, 1), size(corrMatrix, 2));
    end
    
catch ME
    fprintf('  ✗ 상관계수 매트릭스 저장 실패: %s\n', ME.message);
end

%% 11-6. MAT 파일 저장
fprintf('\n▶ [11-6] MATLAB 작업공간 저장\n');
fprintf('----------------------------------------\n');

matFileName = fullfile(outputDir, sprintf('CSR_workspace_%s.mat', timestamp));

try
    % 주요 변수들 저장
    save(matFileName, ...
         'correlationResults', 'csrCategories', 'csrScores', ...
         'finalCSRData', 'finalCompetencyData', 'finalIDs', ...
         'validCompetencyCategories', 'allSignificantResults', ...
         'marginalResults', 'strongCorrelations', ...
         'analysisType', '-v7.3');
    
    % 회귀분석 결과가 있으면 추가
    if exist('regressionResults', 'var') && ~isempty(regressionResults)
        save(matFileName, 'regressionResults', '-append');
    end
    
    fprintf('  ✓ MAT 파일 저장 완료: %s\n', matFileName);
    
catch ME
    fprintf('  ✗ MAT 파일 저장 실패: %s\n', ME.message);
end

%% 11-7. 저장 결과 요약
fprintf('\n▶ [11-7] 저장 결과 요약\n');
fprintf('========================================\n');

% 성공률 계산
totalAttempts = length(csrCategories) + 5; % 개별 시트 + 요약 시트들
successCount = sum(saveStatus.success);
successRate = (successCount / totalAttempts) * 100;

fprintf('📊 저장 통계:\n');
fprintf('  • 시도된 시트: %d개\n', totalAttempts);
fprintf('  • 성공한 시트: %d개\n', length(saveStatus.sheets));
fprintf('  • 성공률: %.1f%%\n', successRate);

if ~isempty(saveStatus.sheets)
    fprintf('\n✓ 저장된 시트 목록:\n');
    for i = 1:length(saveStatus.sheets)
        fprintf('  %2d. %s\n', i, saveStatus.sheets{i});
    end
end

if ~isempty(saveStatus.failed)
    fprintf('\n✗ 실패한 항목:\n');
    for i = 1:length(saveStatus.failed)
        fprintf('  - %s: %s\n', saveStatus.failed{i}{1}, saveStatus.failed{i}{2});
    end
end

fprintf('\n📁 최종 출력 파일:\n');
fprintf('  • Excel: %s\n', resultFileName);
fprintf('  • MAT: %s\n', matFileName);

% 파일 크기 확인
if exist(resultFileName, 'file')
    fileInfo = dir(resultFileName);
    fileSizeMB = fileInfo.bytes / (1024^2);
    fprintf('  • Excel 파일 크기: %.2f MB\n', fileSizeMB);
end

fprintf('\n========================================\n');
fprintf('✅ 결과 저장 완료!\n');
fprintf('========================================\n');
%% 분석 완료 요약
fprintf('\n========================================\n');
fprintf('🎉 CSR vs 역량검사 상관분석 완료\n');
fprintf('========================================\n');

fprintf('📊 분석 요약:\n');
fprintf('  • 분석 대상: %d명\n', length(validRows));
fprintf('  • CSR 카테고리: %d개 (%s)\n', length(csrCategories), strjoin(csrCategories, ', '));
fprintf('  • 역량검사 카테고리: %d개\n', length(validCompetencyCategories));
fprintf('  • 전체 상관분석: %d개\n', length(csrCategories) * length(validCompetencyCategories));

if ~isempty(allSignificantResults)
    fprintf('  • 유의한 상관관계 (p<0.05): %d개\n', length(allSignificantResults));
end

if ~isempty(marginalResults)
    fprintf('  • Marginal 상관관계 (0.05≤p<0.1): %d개\n', length(marginalResults));
end

if ~isempty(strongCorrelations)
    fprintf('  • 강한 상관관계 (|r|>0.1, p<0.1): %d개\n', length(strongCorrelations));
end

fprintf('\n📁 결과 파일:\n');
fprintf('  • Excel: %s\n', resultFileName);
fprintf('  • MAT: %s\n', fullfile(outputDir, 'CSR_vs_competency_workspace.mat'));

fprintf('\n✅ 모든 분석이 성공적으로 완료되었습니다!\n');

%% ===== 보조 함수들 =====

function summaryTable = createCSRSummaryTable(correlationResults, csrCategories, finalCSRData, validRows, validCompetencyCategories)
    % CSR 분석 요약 테이블 생성 함수 (참고 코드 패턴 적용)

    summaryData = {};
    for i = 1:length(csrCategories)
        csrCatName = csrCategories{i};

        % CSR 데이터 통계
        csrData = finalCSRData.(csrCatName);
        correlations = correlationResults.(csrCatName).correlations;
        pValues = correlationResults.(csrCatName).pValues;

        % 상관관계 분석
        numSignificant = sum(pValues < 0.05);
        numMarginal = sum(pValues >= 0.05 & pValues < 0.1);
        numStrong = sum(abs(correlations) > 0.1 & pValues < 0.1);
        maxCorr = max(abs(correlations));
        minCorr = min(abs(correlations));
        meanCorr = mean(abs(correlations));

        % 테이블 행 생성
        summaryData{end+1} = {csrCatName, length(validRows), nanmean(csrData), nanstd(csrData), ...
                              length(validCompetencyCategories), numSignificant, numMarginal, numStrong, ...
                              maxCorr, minCorr, meanCorr};
    end

    % 테이블 생성
    summaryTable = cell2table(summaryData, ...
        'VariableNames', {'CSR_Category', 'Sample_Size', 'CSR_Mean', 'CSR_Std', ...
        'Num_Competency_Categories', 'Num_Significant', 'Num_Marginal', 'Num_Strong', ...
        'Max_Correlation', 'Min_Correlation', 'Mean_Correlation'});
end

function standardizedData = standardizeCSRQuestions(questionData, questionNames, scaleMapping)
    % CSR 문항별 표준화 함수 (문항기반 분석가이드 패턴 적용)
    %
    % 입력:
    %   questionData: 문항 데이터 매트릭스 (행: 응답자, 열: 문항)
    %   questionNames: 문항 이름들
    %   scaleMapping: 문항별 척도 매핑 containers.Map 객체
    %
    % 출력:
    %   standardizedData: [0,1] 범위로 표준화된 데이터

    standardizedData = zeros(size(questionData));

    for q = 1:size(questionData, 2)
        qName = questionNames{q};
        qData = questionData(:, q);
        validData = qData(~isnan(qData));

        if isempty(validData)
            continue;
        end

        % 실제 데이터 범위 분석
        actualMin = min(validData);
        actualMax = max(validData);

        % 사전 정의된 척도 매핑 사용 (있는 경우)
        if isKey(scaleMapping, qName)
            expectedScale = scaleMapping(qName);
            expectedMin = expectedScale(1);
            expectedMax = expectedScale(2);

            % 실제 데이터가 예상 범위를 벗어나는 경우 경고
            if actualMin < expectedMin || actualMax > expectedMax
                fprintf('    경고: %s 실제 범위 [%.1f-%.1f]가 예상 범위 [%.1f-%.1f]를 벗어남\n', ...
                    qName, actualMin, actualMax, expectedMin, expectedMax);
            end

            % 예상 척도로 표준화
            minScale = expectedMin;
            maxScale = expectedMax;
        else
            % 실제 데이터 범위로 표준화
            minScale = actualMin;
            maxScale = actualMax;

            fprintf('    정보: %s 문항의 척도 매핑이 없어 실제 범위 [%.1f-%.1f] 사용\n', ...
                qName, actualMin, actualMax);
        end

        % [0,1] 범위로 표준화
        if maxScale > minScale
            standardizedData(:, q) = (qData - minScale) / (maxScale - minScale);
        else
            % 모든 값이 같은 경우
            standardizedData(:, q) = ones(size(qData)) * 0.5;
        end

        % 표준화 결과 검증
        validStdData = standardizedData(~isnan(qData), q);
        if ~isempty(validStdData)
            stdMin = min(validStdData);
            stdMax = max(validStdData);

            % [0,1] 범위를 벗어나는 경우 클리핑
            if stdMin < 0 || stdMax > 1
                fprintf('    경고: %s 표준화 결과 [%.3f-%.3f]가 [0,1] 범위를 벗어남, 클리핑 적용\n', ...
                    qName, stdMin, stdMax);
                standardizedData(:, q) = max(0, min(1, standardizedData(:, q)));
            end
        end
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