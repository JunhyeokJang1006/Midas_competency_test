%% 역량진단 데이터 문항별 척도 인식 표준화 전후 비교 분석
%
% 목적: 각 문항의 척도 정보를 자동 탐지하여 적절한 표준화 수행
%       1-5점, 1-7점, 1-10점 등 다양한 척도 고려
%
% 작성일: 2025년
% 저장위치: D:\project\HR데이터\matlab\문항기반\역진 데이터 확인

%% ==================== 메인 실행 스크립트 ====================
clear; clc; close all;
rng(42, 'twister'); % 재현성 보장

%% 1. 초기 설정 및 파라미터
fprintf('========================================\n');
fprintf('문항별 척도 인식 표준화 전후 비교 분석\n');
fprintf('========================================\n\n');

config = struct();
config.dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터\';
config.outputPath = 'D:\project\HR데이터\matlab\문항기반\역진 데이터 확인\결과\';
config.standardizeMethods = {'none', 'minmax', 'zscore', 'percentage', 'scale_adjusted'};
config.corrThreshold = 0.3;
config.pValueThreshold = 0.05;

% 데이터 파일 목록
config.dataFiles = {
    '23년_하반기_역량진단_응답데이터.xlsx';
    '24년_상반기_역량진단_응답데이터.xlsx';
    '24년_하반기_역량진단_응답데이터.xlsx';
    '25년_상반기_역량진단_응답데이터.xlsx';
    '23년_상반기_역량진단_응답데이터.xlsx'
};

% 결과 디렉토리 생성
if ~exist(config.outputPath, 'dir')
    mkdir(config.outputPath);
end

fprintf('[1단계] 초기 설정 완료\n');
fprintf('  - 출력 경로: %s\n', config.outputPath);
fprintf('  - 표준화 방법: %s\n', strjoin(config.standardizeMethods, ', '));
fprintf('  - 데이터 파일 수: %d개\n', length(config.dataFiles));

%% 2. 데이터 로드
fprintf('\n[2단계] 문항별 척도 인식 데이터 로드\n');
fprintf('----------------------------------------\n');

try
    [rawData, metadata] = loadCompetencyDataWithScales(config);
    fprintf('✓ 데이터 로드 완료\n');
    fprintf('  - 전체 데이터 크기: %dx%d\n', size(rawData.allScores));
    fprintf('  - Period 수: %d개\n', length(metadata.periods));
    fprintf('  - 인식된 문항 수: %d개\n', length(metadata.itemNames));
    fprintf('  - 척도 그룹 수: %d개\n', length(fieldnames(metadata.scaleGroups)));
catch ME
    fprintf('✗ 데이터 로드 실패: %s\n', ME.message);
    return;
end

%% 3. 표준화 전후 비교 분석
fprintf('\n[3단계] 척도별 표준화 전후 비교 분석\n');
fprintf('----------------------------------------\n');

standardizationResults = struct();
for i = 1:length(config.standardizeMethods)
    method = config.standardizeMethods{i};
    fprintf('  처리 중: %s 방법...\n', method);

    try
        standardizationResults.(method) = analyzeWithScaleAwareStandardization(rawData, metadata, method);
        fprintf('  ✓ %s 완료\n', method);
    catch ME
        fprintf('  ✗ %s 실패: %s\n', method, ME.message);
    end
end

% 표준화 영향 분석
fprintf('\n  표준화 영향 분석 중...\n');
try
    impactAnalysis = analyzeStandardizationImpact(standardizationResults);
    fprintf('  ✓ 영향 분석 완료\n');
catch ME
    fprintf('  ✗ 영향 분석 실패: %s\n', ME.message);
end

%% 4. 분포 분석
fprintf('\n[4단계] 척도별 분포 분석\n');
fprintf('----------------------------------------\n');

try
    distributionStats = performScaleAwareDistributionAnalysis(rawData, metadata);
    fprintf('✓ 분포 분석 완료\n');
    fprintf('  - 전체 평균: %.3f\n', distributionStats.overall.mean);
    fprintf('  - 전체 표준편차: %.3f\n', distributionStats.overall.std);
    fprintf('  - 이상치 개수: %d개\n', sum(distributionStats.overall.outliers));
catch ME
    fprintf('✗ 분포 분석 실패: %s\n', ME.message);
end

%% 5. 시각화 및 결과 저장
fprintf('\n[5단계] 시각화 및 결과 저장\n');
fprintf('----------------------------------------\n');

try
    createScaleAwareVisualizationsAndReports(standardizationResults, impactAnalysis, distributionStats, metadata, config);
    fprintf('✓ 결과 저장 완료\n');
catch ME
    fprintf('✗ 결과 저장 실패: %s\n', ME.message);
end

fprintf('\n========================================\n');
fprintf('척도 인식 분석 완료. 결과가 %s에 저장되었습니다.\n', config.outputPath);
fprintf('========================================\n');

%% ==================== 데이터 처리 함수 ====================

function [data, metadata] = loadCompetencyDataWithScales(config)
% LOADCOMPETENCYDATAWITHSCALES - 척도 정보 포함 데이터 로드
%   Input: config - 설정 구조체
%   Output: data - 원시 데이터, metadata - 척도 정보 포함 메타데이터

    arguments
        config struct
    end

    data = struct();
    metadata = struct();

    allScores = [];
    allIDs = {};
    periodLabels = {};
    allItemNames = {};
    allScaleInfo = struct();

    fprintf('  데이터 파일 처리 중:\n');

    for fileIdx = 1:length(config.dataFiles)
        filename = fullfile(config.dataPath, config.dataFiles{fileIdx});

        % 파일 존재 확인
        if ~exist(filename, 'file')
            fprintf('    ⚠️ 파일 없음: %s\n', config.dataFiles{fileIdx});
            continue;
        end

        try
            % 시트 목록 확인 후 자가 평정 시트 찾기
            [~, sheets] = xlsfinfo(filename);
            selfSheetName = '';

            % '자가' 포함 시트 중 Q 컬럼이 많은 시트 찾기
            for s = 1:length(sheets)
                if contains(sheets{s}, '자가', 'IgnoreCase', true)
                    try
                        testData = readtable(filename, 'Sheet', sheets{s}, 'VariableNamingRule', 'preserve');
                        colNames = testData.Properties.VariableNames;
                        qCols = colNames(contains(colNames, 'Q', 'IgnoreCase', true));

                        if length(qCols) > 10  % Q 컬럼이 10개 이상인 시트를 선택
                            selfSheetName = sheets{s};
                            break;
                        end
                    catch
                        continue;
                    end
                end
            end

            if isempty(selfSheetName)
                fprintf('    ⚠️ 자가 평정 시트 없음: %s\n', config.dataFiles{fileIdx});
                continue;
            end

            % 자가 평정 시트 읽기
            selfData = readtable(filename, 'Sheet', selfSheetName, 'VariableNamingRule', 'preserve');

            if height(selfData) == 0
                fprintf('    ⚠️ 빈 데이터: %s\n', config.dataFiles{fileIdx});
                continue;
            end

            % 문항 척도 탐지
            [itemData, scaleInfo, itemNames] = detectItemScalesAndExtract(selfData);

            if isempty(itemData)
                fprintf('    ⚠️ 유효한 문항 없음: %s\n', config.dataFiles{fileIdx});
                continue;
            end

            % 데이터 누적
            allScores = [allScores; itemData];

            % ID 정보 처리
            if ismember('ID', selfData.Properties.VariableNames)
                periodIDs = selfData.ID;
            else
                periodIDs = cellstr(string(1:height(selfData))');
            end
            allIDs = [allIDs; periodIDs];

            % Period 라벨
            [~, fileBaseName] = fileparts(config.dataFiles{fileIdx});
            periodName = extractBefore(fileBaseName, '_역량진단');
            periodLabels = [periodLabels; repmat({periodName}, height(selfData), 1)];

            % 척도 정보 병합
            if isempty(allItemNames)
                allItemNames = itemNames;
                allScaleInfo = scaleInfo;
            else
                % 동일한 문항 구조인지 확인
                if ~isequal(allItemNames, itemNames)
                    warning('파일별 문항 구조가 다릅니다: %s', config.dataFiles{fileIdx});
                end
            end

            fprintf('    ✓ %s: %d명, %d문항\n', periodName, height(selfData), length(itemNames));

        catch ME
            fprintf('    ✗ 처리 실패: %s (%s)\n', config.dataFiles{fileIdx}, ME.message);
        end
    end

    % 최종 데이터 구성
    data.allScores = allScores;
    data.allIDs = allIDs;
    data.periodLabels = periodLabels;

    % 메타데이터 구성
    metadata.itemNames = allItemNames;
    metadata.itemScales = allScaleInfo;
    metadata.periods = unique(periodLabels, 'stable');
    metadata.scaleGroups = groupItemsByScale(allScaleInfo);
    metadata.source = 'original_files';

    fprintf('  ✓ 전체 %d명의 데이터, %d개 문항 로드 완료\n', size(allScores, 1), length(allItemNames));
end

function [itemData, scaleInfo, itemNames] = detectItemScalesAndExtract(tableData)
% DETECTITEMSCALESANDEXTRACT - 문항 척도 탐지 및 데이터 추출

    arguments
        tableData table
    end

    % 문항 컬럼 찾기 (Q로 시작하고 숫자형)
    colNames = tableData.Properties.VariableNames;
    numericCols = varfun(@isnumeric, tableData, 'output', 'uniform');

    % ID 컬럼 제외
    idCols = contains(colNames, {'id', 'ID', 'idhr'}, 'IgnoreCase', true);

    % Q 관련 컬럼 중 숫자형이고 ID가 아닌 것들
    isQCol = contains(colNames, 'Q', 'IgnoreCase', true) | startsWith(colNames, 'Q');
    validCols = numericCols & isQCol & ~idCols;

    itemNames = colNames(validCols);
    itemData = tableData{:, validCols};

    % 각 문항의 척도 정보 탐지
    scaleInfo = struct();

    for i = 1:length(itemNames)
        itemName = itemNames{i};
        itemValues = itemData(:, i);
        validValues = itemValues(~isnan(itemValues));

        if ~isempty(validValues)
            scaleInfo.(itemName) = detectItemScale(validValues, itemName);
        else
            scaleInfo.(itemName) = struct('min', NaN, 'max', NaN, 'type', 'unknown');
        end
    end
end

function scaleResult = detectItemScale(values, itemName)
% DETECTITEMSCALE - 개별 문항의 척도 탐지

    arguments
        values (:,1) double
        itemName (1,:) char
    end

    scaleResult = struct();

    % 기본 통계
    minVal = min(values);
    maxVal = max(values);
    uniqueVals = unique(values);
    numUnique = length(uniqueVals);
    isInteger = all(values == floor(values));

    scaleResult.min = minVal;
    scaleResult.max = maxVal;
    scaleResult.uniqueCount = numUnique;
    scaleResult.isInteger = isInteger;
    scaleResult.uniqueValues = uniqueVals;

    % 척도 타입 판정
    if isInteger
        if minVal == 1
            if maxVal == 5 && numUnique <= 5
                scaleResult.type = '1-5점';
                scaleResult.theoreticalMin = 1;
                scaleResult.theoreticalMax = 5;
            elseif maxVal == 7 && numUnique <= 7
                scaleResult.type = '1-7점';
                scaleResult.theoreticalMin = 1;
                scaleResult.theoreticalMax = 7;
            elseif maxVal == 10 && numUnique <= 10
                scaleResult.type = '1-10점';
                scaleResult.theoreticalMin = 1;
                scaleResult.theoreticalMax = 10;
            else
                scaleResult.type = sprintf('1-%d점', maxVal);
                scaleResult.theoreticalMin = 1;
                scaleResult.theoreticalMax = maxVal;
            end
        elseif minVal == 0
            scaleResult.type = sprintf('0-%d점', maxVal);
            scaleResult.theoreticalMin = 0;
            scaleResult.theoreticalMax = maxVal;
        else
            scaleResult.type = sprintf('%.0f-%.0f점', minVal, maxVal);
            scaleResult.theoreticalMin = minVal;
            scaleResult.theoreticalMax = maxVal;
        end
    else
        scaleResult.type = '연속형';
        scaleResult.theoreticalMin = minVal;
        scaleResult.theoreticalMax = maxVal;
    end

    % 경고 사항 확인
    scaleResult.warnings = {};

    % 예상 범위와 실제 범위 불일치 확인
    if isInteger && minVal == 1 && (maxVal == 5 || maxVal == 7 || maxVal == 10)
        expectedRange = 1:maxVal;
        if ~all(ismember(expectedRange, uniqueVals))
            missingVals = expectedRange(~ismember(expectedRange, uniqueVals));
            scaleResult.warnings{end+1} = sprintf('누락된 척도값: %s', mat2str(missingVals));
        end
    end

    % 이상치 확인 (IQR 기준)
    Q1 = quantile(values, 0.25);
    Q3 = quantile(values, 0.75);
    IQR = Q3 - Q1;
    outliers = values < (Q1 - 1.5*IQR) | values > (Q3 + 1.5*IQR);
    if sum(outliers) > 0
        scaleResult.warnings{end+1} = sprintf('이상치 %d개 발견', sum(outliers));
    end
end

function scaleGroups = groupItemsByScale(scaleInfo)
% GROUPITEMSBYSCALE - 척도별 문항 그룹핑

    arguments
        scaleInfo struct
    end

    scaleGroups = struct();
    itemNames = fieldnames(scaleInfo);

    % 척도 타입별로 그룹핑
    scaleTypes = {};
    for i = 1:length(itemNames)
        scaleType = scaleInfo.(itemNames{i}).type;
        if ~ismember(scaleType, scaleTypes)
            scaleTypes{end+1} = scaleType;
        end
    end

    % 각 척도 타입별로 문항 목록 생성
    for i = 1:length(scaleTypes)
        scaleType = scaleTypes{i};
        groupName = sprintf('scale_%s', regexprep(scaleType, '[^a-zA-Z0-9]', '_'));

        itemsInGroup = {};
        for j = 1:length(itemNames)
            if strcmp(scaleInfo.(itemNames{j}).type, scaleType)
                itemsInGroup{end+1} = itemNames{j};
            end
        end

        scaleGroups.(groupName).type = scaleType;
        scaleGroups.(groupName).items = itemsInGroup;
        scaleGroups.(groupName).count = length(itemsInGroup);
    end
end

function result = analyzeWithScaleAwareStandardization(data, metadata, method)
% ANALYZEWITHSCALEAWARESTANDARDIZATION - 척도 인식 표준화 분석

    arguments
        data struct
        metadata struct
        method (1,:) char {mustBeMember(method, {'zscore', 'minmax', 'percentage', 'none', 'scale_adjusted'})} = 'zscore'
    end

    result = struct();
    result.method = method;
    result.originalData = data.allScores;  % 원본 저장
    result.itemNames = metadata.itemNames;
    result.itemScales = metadata.itemScales;

    if isempty(data.allScores)
        warning('빈 데이터입니다.');
        result.standardizedData = [];
        result.correlation = [];
        result.corrValue = [];
        result.pValue = [];
        result.validationResults = struct();
        return;
    end

    % 문항별 표준화 수행
    standardizedData = zeros(size(data.allScores));
    validationResults = struct();
    validationResults.outOfRange = [];
    validationResults.scaleMismatches = {};

    for i = 1:length(metadata.itemNames)
        itemName = metadata.itemNames{i};
        itemData = data.allScores(:, i);
        scaleInfo = metadata.itemScales.(itemName);

        % 문항별 표준화
        switch method
            case 'none'
                standardizedData(:, i) = itemData;

            case 'minmax'
                % 실제 데이터의 min/max 사용
                minVal = min(itemData, [], 'omitnan');
                maxVal = max(itemData, [], 'omitnan');
                if maxVal > minVal
                    standardizedData(:, i) = (itemData - minVal) ./ (maxVal - minVal);
                else
                    standardizedData(:, i) = itemData;
                end

            case 'zscore'
                % 문항별 z-score 표준화
                meanVal = mean(itemData, 'omitnan');
                stdVal = std(itemData, 'omitnan');
                if stdVal > 0
                    standardizedData(:, i) = (itemData - meanVal) ./ stdVal;
                else
                    standardizedData(:, i) = itemData - meanVal;
                end

            case 'percentage'
                % 실제 최대값 기준 백분율
                maxVal = max(itemData, [], 'omitnan');
                if maxVal > 0
                    standardizedData(:, i) = (itemData ./ maxVal) * 100;
                else
                    standardizedData(:, i) = itemData;
                end

            case 'scale_adjusted'
                % 이론적 척도 범위 기준 정규화
                if isfield(scaleInfo, 'theoreticalMin') && isfield(scaleInfo, 'theoreticalMax')
                    theoreticalMin = scaleInfo.theoreticalMin;
                    theoreticalMax = scaleInfo.theoreticalMax;
                    if theoreticalMax > theoreticalMin
                        standardizedData(:, i) = (itemData - theoreticalMin) ./ (theoreticalMax - theoreticalMin);

                        % 범위 검증 (0-1 범위를 벗어나는지 확인)
                        outOfRange = standardizedData(:, i) < -0.01 | standardizedData(:, i) > 1.01;  % 약간의 오차 허용
                        if sum(outOfRange) > 0
                            validationResults.outOfRange = [validationResults.outOfRange; ...
                                table({itemName}, sum(outOfRange), ...
                                'VariableNames', {'Item', 'OutOfRangeCount'})];
                        end
                    else
                        standardizedData(:, i) = itemData;
                        validationResults.scaleMismatches{end+1} = sprintf('%s: 이론적 범위 오류', itemName);
                    end
                else
                    % 이론적 범위가 없으면 실제 데이터 기준으로
                    minVal = min(itemData, [], 'omitnan');
                    maxVal = max(itemData, [], 'omitnan');
                    if maxVal > minVal
                        standardizedData(:, i) = (itemData - minVal) ./ (maxVal - minVal);
                    else
                        standardizedData(:, i) = itemData;
                    end
                    validationResults.scaleMismatches{end+1} = sprintf('%s: 이론적 범위 정보 없음', itemName);
                end
        end
    end

    result.standardizedData = standardizedData;
    result.validationResults = validationResults;

    % 상관분석 수행
    try
        validData = standardizedData;
        validData = validData(:, ~all(isnan(validData), 1)); % NaN 컬럼 제거

        if size(validData, 2) >= 2
            [corrMatrix, pValueMatrix] = corrcoef(validData, 'Rows', 'complete');
            result.correlation = corrMatrix;
            result.corrValue = corrMatrix;
            result.pValue = pValueMatrix;
        else
            result.correlation = [];
            result.corrValue = [];
            result.pValue = [];
        end
    catch ME
        warning('상관분석 실패: %s', ME.message);
        result.correlation = [];
        result.corrValue = [];
        result.pValue = [];
    end
end

function impact = analyzeStandardizationImpact(results)
% ANALYZESTANDARDIZATIONIMPACT - 표준화 전후 비교 메인 함수 (동일)

    arguments
        results struct
    end

    methods = fieldnames(results);
    impact = struct();

    % 원본(none) 대비 각 방법 비교
    if ~isfield(results, 'none') || isempty(results.none.corrValue)
        warning('원본 데이터 분석 결과가 없습니다.');
        return;
    end

    baseline = results.none;

    for i = 1:length(methods)
        if strcmp(methods{i}, 'none') || isempty(results.(methods{i}).corrValue)
            continue;
        end

        methodResult = results.(methods{i});

        % 상관계수 변화량 계산
        if isequal(size(baseline.corrValue), size(methodResult.corrValue))
            impact.(methods{i}).corrChange = methodResult.corrValue - baseline.corrValue;

            % p-value 변화 분석
            impact.(methods{i}).pValueChange = methodResult.pValue - baseline.pValue;

            % 유의한 상관관계 개수 변화
            sigBaseline = sum(baseline.pValue(:) < 0.05);
            sigMethod = sum(methodResult.pValue(:) < 0.05);
            impact.(methods{i}).sigCountChange = sigMethod - sigBaseline;

            % 상관계수 순위 변화
            impact.(methods{i}).rankChange = calculateRankChange(baseline.corrValue, methodResult.corrValue);

            % 평균 절대 변화량
            impact.(methods{i}).meanAbsChange = mean(abs(impact.(methods{i}).corrChange(:)), 'omitnan');

            % 검증 결과 포함
            if isfield(methodResult, 'validationResults')
                impact.(methods{i}).validationResults = methodResult.validationResults;
            end
        else
            warning('상관 매트릭스 크기가 다릅니다: %s', methods{i});
        end
    end
end

function rankChange = calculateRankChange(baseline, method)
% CALCULATERANKCHANGE - 상관계수 순위 변화 계산 (동일)

    % 상관계수를 벡터로 변환 (대각선 제외)
    n = size(baseline, 1);
    maskTril = tril(true(n), -1); % 하삼각 행렬

    baselineVec = baseline(maskTril);
    methodVec = method(maskTril);

    % 순위 계산
    [~, baselineRank] = sort(baselineVec, 'descend');
    [~, methodRank] = sort(methodVec, 'descend');

    % 순위 변화 계산 (Spearman 순위 상관계수)
    rankChange = corr(baselineRank, methodRank, 'Type', 'Spearman');
end

function stats = performScaleAwareDistributionAnalysis(data, metadata)
% PERFORMSCALEAWAREDISTRIBUTIONANALYSIS - 척도별 분포 분석

    arguments
        data struct
        metadata struct
    end

    stats = struct();

    if isempty(data.allScores)
        return;
    end

    % 전체 분포 분석
    stats.overall = calculateDescriptiveStats(data.allScores);
    stats.overall.outliers = detectOutliers(data.allScores, 'iqr');
    stats.overall.normality = performNormalityTest(data.allScores);

    % 척도별 분포 분석
    scaleGroups = metadata.scaleGroups;
    groupNames = fieldnames(scaleGroups);

    stats.byScaleGroup = struct();

    for i = 1:length(groupNames)
        groupName = groupNames{i};
        group = scaleGroups.(groupName);

        % 해당 그룹의 문항 인덱스 찾기
        itemIndices = [];
        for j = 1:length(group.items)
            itemIdx = find(strcmp(metadata.itemNames, group.items{j}), 1);
            if ~isempty(itemIdx)
                itemIndices = [itemIndices, itemIdx];
            end
        end

        if ~isempty(itemIndices)
            groupData = data.allScores(:, itemIndices);

            stats.byScaleGroup.(groupName) = struct();
            stats.byScaleGroup.(groupName).scaleType = group.type;
            stats.byScaleGroup.(groupName).itemCount = length(itemIndices);
            stats.byScaleGroup.(groupName).stats = calculateDescriptiveStats(groupData);
            stats.byScaleGroup.(groupName).outliers = detectOutliers(groupData, 'iqr');
        end
    end

    % 시점별 분석 (Period별 분석)
    if isfield(data, 'periodLabels') && ~isempty(data.periodLabels)
        stats.timeSeriesAnalysis = analyzeByPeriod(data.allScores, data.periodLabels);
    end
end

function stats = calculateDescriptiveStats(data)
% CALCULATEDESCRIPTIVESTATS - 기술통계 계산 (동일)

    arguments
        data (:,:) double
    end

    stats = struct();

    % 전체 데이터에 대한 통계 (모든 컬럼을 하나로 합친 벡터)
    allValues = data(:);
    allValues = allValues(~isnan(allValues)); % NaN 제거

    if isempty(allValues)
        warning('유효한 데이터가 없습니다.');
        stats.mean = NaN;
        stats.median = NaN;
        stats.std = NaN;
        stats.quartiles = [NaN, NaN, NaN];
        stats.skewness = NaN;
        stats.kurtosis = NaN;
        stats.iqr = NaN;
        return;
    end

    stats.mean = mean(allValues);
    stats.median = median(allValues);
    stats.std = std(allValues);
    stats.quartiles = quantile(allValues, [0.25, 0.5, 0.75]);
    stats.skewness = skewness(allValues);
    stats.kurtosis = kurtosis(allValues);
    stats.iqr = iqr(allValues);
end

function outliers = detectOutliers(data, method)
% DETECTOUTLIERS - 이상치 탐지 (동일)

    arguments
        data (:,:) double
        method (1,:) char {mustBeMember(method, {'iqr', 'zscore'})} = 'iqr'
    end

    outliers = false(size(data));

    switch method
        case 'iqr'
            for col = 1:size(data, 2)
                colData = data(:, col);
                validData = colData(~isnan(colData));

                if length(validData) >= 4
                    Q1 = quantile(validData, 0.25);
                    Q3 = quantile(validData, 0.75);
                    IQR = Q3 - Q1;

                    lowerBound = Q1 - 1.5 * IQR;
                    upperBound = Q3 + 1.5 * IQR;

                    outliers(:, col) = colData < lowerBound | colData > upperBound;
                end
            end

        case 'zscore'
            for col = 1:size(data, 2)
                colData = data(:, col);
                validData = colData(~isnan(colData));

                if length(validData) >= 2
                    z = abs((colData - mean(validData)) / std(validData));
                    outliers(:, col) = z > 3;
                end
            end
    end
end

function normality = performNormalityTest(data)
% PERFORMNORMALITYTEST - 정규성 검정 (동일)

    arguments
        data (:,:) double
    end

    normality = struct();

    % 전체 데이터를 하나의 벡터로 변환
    allValues = data(:);
    allValues = allValues(~isnan(allValues)); % NaN 제거

    if length(allValues) < 4
        normality.pValue = NaN;
        normality.isNormal = false;
        normality.testStat = NaN;
        return;
    end

    try
        % Kolmogorov-Smirnov test
        [~, pValue, ksStat] = kstest((allValues - mean(allValues)) / std(allValues));

        normality.pValue = pValue;
        normality.isNormal = pValue > 0.05;
        normality.testStat = ksStat;
        normality.testType = 'Kolmogorov-Smirnov';
    catch
        % Fallback: Shapiro-Wilk test (간단한 구현)
        normality.pValue = NaN;
        normality.isNormal = false;
        normality.testStat = NaN;
        normality.testType = 'None (계산 실패)';
    end
end

function timeAnalysis = analyzeByPeriod(data, periodLabels)
% ANALYZEBYPERIOD - 시점별 분석 (동일)

    arguments
        data (:,:) double
        periodLabels (:,1) cell
    end

    uniquePeriods = unique(periodLabels);
    timeAnalysis = struct();

    for p = 1:length(uniquePeriods)
        periodName = uniquePeriods{p};
        periodIdx = strcmp(periodLabels, periodName);

        if sum(periodIdx) > 0
            periodData = data(periodIdx, :);
            timeAnalysis.(periodName) = calculateDescriptiveStats(periodData);
            timeAnalysis.(periodName).sampleSize = sum(periodIdx);
        end
    end
end

function createScaleAwareVisualizationsAndReports(standardizationResults, impactAnalysis, distributionStats, metadata, config)
% CREATESCALEAWAREVISUALIZATIONSANDREPORTS - 척도별 시각화 및 보고서 생성

    arguments
        standardizationResults struct
        impactAnalysis struct
        distributionStats struct
        metadata struct
        config struct
    end

    try
        % 1. 척도별 종합 시각화 생성
        createScaleAwareMainVisualization(standardizationResults, impactAnalysis, distributionStats, metadata, config);

        % 2. Excel 결과 저장 (척도 정보 포함)
        exportScaleAwareToExcel(standardizationResults, impactAnalysis, distributionStats, metadata, config);

        % 3. 척도별 종합 보고서 생성
        createScaleAwareReport(distributionStats, metadata, config);

    catch ME
        warning('시각화 생성 중 오류: %s', ME.message);
    end
end

function createScaleAwareMainVisualization(standardizationResults, impactAnalysis, distributionStats, metadata, config)
% CREATESCALEAWAREMAINVISUALIZATION - 척도별 메인 시각화 생성

    % 메인 Figure 생성
    fig = figure('Position', [50, 50, 1600, 1000], 'Color', 'white');

    % 표준화 방법별 히스토그램 비교
    methods = fieldnames(standardizationResults);
    validMethods = {};

    for i = 1:length(methods)
        if ~isempty(standardizationResults.(methods{i}).standardizedData)
            validMethods{end+1} = methods{i};
        end
    end

    % 2x3 서브플롯으로 표준화 방법별 분포
    if length(validMethods) >= 5
        for i = 1:5
            subplot(3, 3, i);
            method = validMethods{i};
            data = standardizationResults.(method).standardizedData;

            if ~isempty(data)
                allValues = data(:);
                allValues = allValues(~isnan(allValues));

                if ~isempty(allValues)
                    histogram(allValues, 30, 'FaceColor', [0.2, 0.6, 0.8], 'EdgeColor', 'none', 'FaceAlpha', 0.7);
                    title(sprintf('%s 방법', upper(method)), 'FontSize', 12);
                    xlabel('표준화된 값');
                    ylabel('빈도');
                    grid on;

                    % 범위 정보 추가
                    if strcmp(method, 'scale_adjusted') || strcmp(method, 'minmax')
                        xlim([0, 1]);
                    end
                end
            end
        end
    end

    % 척도별 분포 비교 (6번째 서브플롯)
    subplot(3, 3, 6);
    if isfield(distributionStats, 'byScaleGroup')
        groupNames = fieldnames(distributionStats.byScaleGroup);
        groupMeans = [];
        groupLabels = {};

        for i = 1:length(groupNames)
            groupData = distributionStats.byScaleGroup.(groupNames{i});
            groupMeans(i) = groupData.stats.mean;
            groupLabels{i} = regexprep(groupNames{i}, '_', ' ');
        end

        bar(groupMeans, 'FaceColor', [0.4, 0.7, 0.4]);
        set(gca, 'XTickLabel', groupLabels);
        title('척도별 평균값');
        ylabel('평균');
        grid on;
        xtickangle(45);
    end

    % 상관계수 변화량 (7번째 서브플롯)
    subplot(3, 3, 7);
    if ~isempty(fieldnames(impactAnalysis))
        impactMethods = fieldnames(impactAnalysis);
        meanChanges = zeros(length(impactMethods), 1);

        for i = 1:length(impactMethods)
            if isfield(impactAnalysis.(impactMethods{i}), 'meanAbsChange')
                meanChanges(i) = impactAnalysis.(impactMethods{i}).meanAbsChange;
            end
        end

        bar(meanChanges, 'FaceColor', [0.8, 0.4, 0.2]);
        set(gca, 'XTickLabel', impactMethods);
        title('표준화 방법별 상관계수 변화');
        ylabel('평균 절대 변화량');
        grid on;
        xtickangle(45);
    end

    % 검증 결과 표시 (8번째 서브플롯)
    subplot(3, 3, 8);
    if isfield(standardizationResults, 'scale_adjusted')
        validationResults = standardizationResults.scale_adjusted.validationResults;

        % 범위 벗어난 항목 수 표시
        if ~isempty(validationResults.outOfRange)
            outOfRangeCounts = validationResults.outOfRange.OutOfRangeCount;
            bar(outOfRangeCounts, 'FaceColor', [0.8, 0.2, 0.2]);
            title('척도 조정 검증: 범위 초과');
            xlabel('문항');
            ylabel('범위 초과 건수');
            grid on;
        else
            text(0.5, 0.5, '모든 문항이 정상 범위 내', 'HorizontalAlignment', 'center', ...
                'FontSize', 14, 'Color', [0, 0.7, 0]);
            title('척도 조정 검증: 통과');
            axis off;
        end
    end

    % 전체 기술통계 (9번째 서브플롯)
    subplot(3, 3, 9);
    if ~isempty(fieldnames(distributionStats))
        stats = [distributionStats.overall.mean, distributionStats.overall.median, distributionStats.overall.std, ...
                distributionStats.overall.skewness, distributionStats.overall.kurtosis];
        statNames = {'평균', '중앙값', '표준편차', '왜도', '첨도'};

        bar(stats, 'FaceColor', [0.4, 0.7, 0.4]);
        set(gca, 'XTickLabel', statNames);
        title('전체 분포 통계');
        ylabel('값');
        grid on;
        xtickangle(45);
    end

    % 전체 제목 추가
    sgtitle('척도 인식 표준화 전후 비교 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');

    % 그림 저장
    saveas(fig, fullfile(config.outputPath, '척도인식_표준화영향분석_시각화_scale_adjusted.fig'));
    saveas(fig, fullfile(config.outputPath, '척도인식_표준화영향분석_시각화_scale_adjusted.png'), 'png');

    fprintf('  ✓ 시각화 저장: %s\n', fullfile(config.outputPath, '척도인식_표준화영향분석_시각화_scale_adjusted.png'));
end

function exportScaleAwareToExcel(standardizationResults, impactAnalysis, distributionStats, metadata, config)
% EXPORTSCALEAWARETOEXCEL - 척도 정보 포함 Excel 내보내기

    filename = fullfile(config.outputPath, '표준화방법별_비교결과_scale_adjusted.xlsx');

    try
        % Sheet 1: 문항별 척도 정보
        scaleInfoTable = createItemScaleInfoTable(metadata);
        if ~isempty(scaleInfoTable)
            writetable(scaleInfoTable, filename, 'Sheet', '문항별척도정보', 'WriteMode', 'overwritesheet');
        end

        % Sheet 2: 표준화 방법별 기술통계 비교
        statsComparison = createStatsComparisonTable(standardizationResults);
        if ~isempty(statsComparison)
            writetable(statsComparison, filename, 'Sheet', '기술통계비교', 'WriteMode', 'overwritesheet');
        end

        % Sheet 3: 영향 분석 결과
        if ~isempty(fieldnames(impactAnalysis))
            impactTable = createScaleAwareImpactAnalysisTable(impactAnalysis);
            if ~isempty(impactTable)
                writetable(impactTable, filename, 'Sheet', '영향분석', 'WriteMode', 'overwritesheet');
            end
        end

        % Sheet 4: 척도별 분포 통계
        if isfield(distributionStats, 'byScaleGroup')
            scaleGroupTable = createScaleGroupStatsTable(distributionStats.byScaleGroup);
            if ~isempty(scaleGroupTable)
                writetable(scaleGroupTable, filename, 'Sheet', '척도별분포통계', 'WriteMode', 'overwritesheet');
            end
        end

        % Sheet 5: 검증 결과
        if isfield(standardizationResults, 'scale_adjusted')
            validationTable = createValidationResultsTable(standardizationResults.scale_adjusted.validationResults);
            if ~isempty(validationTable)
                writetable(validationTable, filename, 'Sheet', '검증결과', 'WriteMode', 'overwritesheet');
            end
        end

        fprintf('  ✓ Excel 파일 저장: %s\n', filename);

    catch ME
        warning('Excel 저장 실패: %s', ME.message);
    end
end

function scaleInfoTable = createItemScaleInfoTable(metadata)
% CREATEITEMSCALEINFOTABLE - 문항별 척도 정보 테이블 생성

    itemNames = metadata.itemNames;
    scaleInfo = metadata.itemScales;

    items = {};
    scaleTypes = {};
    minVals = [];
    maxVals = [];
    theoreticalMins = [];
    theoreticalMaxs = [];
    warnings = {};

    for i = 1:length(itemNames)
        itemName = itemNames{i};
        info = scaleInfo.(itemName);

        items{end+1} = itemName;
        scaleTypes{end+1} = info.type;
        minVals(end+1) = info.min;
        maxVals(end+1) = info.max;

        if isfield(info, 'theoreticalMin')
            theoreticalMins(end+1) = info.theoreticalMin;
        else
            theoreticalMins(end+1) = NaN;
        end

        if isfield(info, 'theoreticalMax')
            theoreticalMaxs(end+1) = info.theoreticalMax;
        else
            theoreticalMaxs(end+1) = NaN;
        end

        if isfield(info, 'warnings') && ~isempty(info.warnings)
            warnings{end+1} = strjoin(info.warnings, '; ');
        else
            warnings{end+1} = '';
        end
    end

    scaleInfoTable = table(items', scaleTypes', minVals', maxVals', ...
        theoreticalMins', theoreticalMaxs', warnings', ...
        'VariableNames', {'문항', '척도타입', '실제최소값', '실제최대값', '이론적최소값', '이론적최대값', '경고사항'});
end

function statsTable = createStatsComparisonTable(standardizationResults)
% CREATESTATSCOMPARISONTABLE - 기술통계 비교 테이블 생성 (동일)

    methods = fieldnames(standardizationResults);

    methodNames = {};
    means = [];
    stds = [];
    mins = [];
    maxs = [];

    for i = 1:length(methods)
        result = standardizationResults.(methods{i});
        if ~isempty(result.standardizedData)
            allValues = result.standardizedData(:);
            allValues = allValues(~isnan(allValues));

            if ~isempty(allValues)
                methodNames{end+1} = methods{i};
                means(end+1) = mean(allValues);
                stds(end+1) = std(allValues);
                mins(end+1) = min(allValues);
                maxs(end+1) = max(allValues);
            end
        end
    end

    if ~isempty(methodNames)
        statsTable = table(methodNames', means', stds', mins', maxs', ...
            'VariableNames', {'표준화방법', '평균', '표준편차', '최솟값', '최댓값'});
    else
        statsTable = table();
    end
end

function impactTable = createScaleAwareImpactAnalysisTable(impactAnalysis)
% CREATESCALEAWAREIMPACTANALYSISTABLE - 검증 결과 포함 영향 분석 테이블

    methods = fieldnames(impactAnalysis);

    methodNames = {};
    meanAbsChanges = [];
    sigCountChanges = [];
    rankChanges = [];
    validationIssues = {};

    for i = 1:length(methods)
        impact = impactAnalysis.(methods{i});

        methodNames{end+1} = methods{i};

        if isfield(impact, 'meanAbsChange')
            meanAbsChanges(end+1) = impact.meanAbsChange;
        else
            meanAbsChanges(end+1) = NaN;
        end

        if isfield(impact, 'sigCountChange')
            sigCountChanges(end+1) = impact.sigCountChange;
        else
            sigCountChanges(end+1) = NaN;
        end

        if isfield(impact, 'rankChange')
            rankChanges(end+1) = impact.rankChange;
        else
            rankChanges(end+1) = NaN;
        end

        % 검증 결과 요약
        if isfield(impact, 'validationResults')
            validationResult = impact.validationResults;
            issues = {};

            if ~isempty(validationResult.outOfRange)
                issues{end+1} = sprintf('범위초과: %d건', height(validationResult.outOfRange));
            end

            if ~isempty(validationResult.scaleMismatches)
                issues{end+1} = sprintf('척도오류: %d건', length(validationResult.scaleMismatches));
            end

            if isempty(issues)
                validationIssues{end+1} = '정상';
            else
                validationIssues{end+1} = strjoin(issues, ', ');
            end
        else
            validationIssues{end+1} = '해당없음';
        end
    end

    if ~isempty(methodNames)
        impactTable = table(methodNames', meanAbsChanges', sigCountChanges', rankChanges', validationIssues', ...
            'VariableNames', {'표준화방법', '평균절대변화량', '유의상관개수변화', '순위상관계수', '검증결과'});
    else
        impactTable = table();
    end
end

function scaleGroupTable = createScaleGroupStatsTable(scaleGroupStats)
% CREATESCALEGROUPSTATSTABLE - 척도 그룹별 통계 테이블

    groupNames = fieldnames(scaleGroupStats);

    scaleTypes = {};
    itemCounts = [];
    groupMeans = [];
    groupStds = [];
    groupSkewness = [];

    for i = 1:length(groupNames)
        groupData = scaleGroupStats.(groupNames{i});

        scaleTypes{end+1} = groupData.scaleType;
        itemCounts(end+1) = groupData.itemCount;
        groupMeans(end+1) = groupData.stats.mean;
        groupStds(end+1) = groupData.stats.std;
        groupSkewness(end+1) = groupData.stats.skewness;
    end

    scaleGroupTable = table(scaleTypes', itemCounts', groupMeans', groupStds', groupSkewness', ...
        'VariableNames', {'척도타입', '문항수', '평균', '표준편차', '왜도'});
end

function validationTable = createValidationResultsTable(validationResults)
% CREATEVALIDATIONRESULTSTABLE - 검증 결과 테이블

    if isempty(validationResults.outOfRange) && isempty(validationResults.scaleMismatches)
        validationTable = table({'검증 통과'}, {'모든 항목이 정상 범위 내'}, ...
            'VariableNames', {'검증항목', '결과'});
        return;
    end

    validationItems = {};
    validationResults_cell = {};

    % 범위 초과 항목들
    if ~isempty(validationResults.outOfRange)
        for i = 1:height(validationResults.outOfRange)
            validationItems{end+1} = sprintf('범위 초과: %s', validationResults.outOfRange.Item{i});
            validationResults_cell{end+1} = sprintf('%d건 초과', validationResults.outOfRange.OutOfRangeCount(i));
        end
    end

    % 척도 불일치 항목들
    if ~isempty(validationResults.scaleMismatches)
        for i = 1:length(validationResults.scaleMismatches)
            validationItems{end+1} = '척도 불일치';
            validationResults_cell{end+1} = validationResults.scaleMismatches{i};
        end
    end

    if ~isempty(validationItems)
        validationTable = table(validationItems', validationResults_cell', ...
            'VariableNames', {'검증항목', '결과'});
    else
        validationTable = table();
    end
end

function createScaleAwareReport(stats, metadata, config)
% CREATESCALEAWAREREPORT - 척도별 종합 분석 보고서 생성

    reportFile = fullfile(config.outputPath, '척도인식_분포분석보고서_scale_adjusted.txt');

    try
        fid = fopen(reportFile, 'w', 'n', 'UTF-8');

        fprintf(fid, '========== 척도 인식 역량진단 데이터 분포 분석 보고서 ==========\n\n');
        fprintf(fid, '분석 일시: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

        % 척도 정보 요약
        fprintf(fid, '1. 척도 정보 요약\n');
        scaleGroups = metadata.scaleGroups;
        groupNames = fieldnames(scaleGroups);

        fprintf(fid, '   총 문항 수: %d개\n', length(metadata.itemNames));
        fprintf(fid, '   척도 그룹 수: %d개\n', length(groupNames));

        for i = 1:length(groupNames)
            group = scaleGroups.(groupNames{i});
            fprintf(fid, '   - %s: %d개 문항\n', group.type, group.count);
        end

        % 전체 기술통계
        if ~isempty(fieldnames(stats))
            fprintf(fid, '\n2. 전체 기술통계\n');
            fprintf(fid, '   - 평균: %.3f\n', stats.overall.mean);
            fprintf(fid, '   - 중앙값: %.3f\n', stats.overall.median);
            fprintf(fid, '   - 표준편차: %.3f\n', stats.overall.std);
            fprintf(fid, '   - 사분위수: [%.3f, %.3f, %.3f]\n', ...
                stats.overall.quartiles(1), stats.overall.quartiles(2), stats.overall.quartiles(3));
            fprintf(fid, '   - 왜도: %.3f', stats.overall.skewness);
            if stats.overall.skewness > 1
                fprintf(fid, ' (강한 우편향)');
            elseif stats.overall.skewness > 0.5
                fprintf(fid, ' (약한 우편향)');
            elseif stats.overall.skewness < -1
                fprintf(fid, ' (강한 좌편향)');
            elseif stats.overall.skewness < -0.5
                fprintf(fid, ' (약한 좌편향)');
            else
                fprintf(fid, ' (대칭)');
            end
            fprintf(fid, '\n');

            fprintf(fid, '   - 첨도: %.3f', stats.overall.kurtosis);
            if stats.overall.kurtosis > 3
                fprintf(fid, ' (뾰족한 분포)');
            elseif stats.overall.kurtosis < 3
                fprintf(fid, ' (평평한 분포)');
            else
                fprintf(fid, ' (정규분포 수준)');
            end
            fprintf(fid, '\n');

            fprintf(fid, '\n3. 이상치 분석\n');
            fprintf(fid, '   - 탐지된 이상치 개수: %d\n', sum(stats.overall.outliers(:)));
            if numel(stats.overall.outliers) > 0
                fprintf(fid, '   - 전체 대비 비율: %.2f%%\n', 100*sum(stats.overall.outliers(:))/numel(stats.overall.outliers));
            end

            if isfield(stats.overall, 'normality')
                fprintf(fid, '\n4. 정규성 검정\n');
                fprintf(fid, '   - 검정 방법: %s\n', stats.overall.normality.testType);
                fprintf(fid, '   - p-value: %.4f\n', stats.overall.normality.pValue);
                if stats.overall.normality.isNormal
                    fprintf(fid, '   - 정규성 만족 (α = 0.05 기준)\n');
                else
                    fprintf(fid, '   - 정규성 불만족 (α = 0.05 기준)\n');
                end
            end

            % 척도별 분석
            if isfield(stats, 'byScaleGroup')
                fprintf(fid, '\n5. 척도별 분석\n');
                scaleGroups = fieldnames(stats.byScaleGroup);
                for i = 1:length(scaleGroups)
                    groupData = stats.byScaleGroup.(scaleGroups{i});
                    fprintf(fid, '   [%s]\n', groupData.scaleType);
                    fprintf(fid, '     - 문항 수: %d개\n', groupData.itemCount);
                    fprintf(fid, '     - 평균: %.3f\n', groupData.stats.mean);
                    fprintf(fid, '     - 표준편차: %.3f\n', groupData.stats.std);
                    fprintf(fid, '     - 왜도: %.3f\n', groupData.stats.skewness);
                end
            end

            if isfield(stats, 'timeSeriesAnalysis')
                fprintf(fid, '\n6. 시점별 분석\n');
                periods = fieldnames(stats.timeSeriesAnalysis);
                for p = 1:length(periods)
                    periodStats = stats.timeSeriesAnalysis.(periods{p});
                    fprintf(fid, '   [%s]\n', periods{p});
                    fprintf(fid, '     - 표본 수: %d\n', periodStats.sampleSize);
                    fprintf(fid, '     - 평균: %.3f\n', periodStats.mean);
                    fprintf(fid, '     - 표준편차: %.3f\n', periodStats.std);
                end
            end
        else
            fprintf(fid, '분석할 데이터가 없습니다.\n');
        end

        fprintf(fid, '\n========== 보고서 종료 ==========\n');
        fclose(fid);

        fprintf('  ✓ 보고서 생성: %s\n', reportFile);

    catch ME
        if exist('fid', 'var') && fid > 0
            fclose(fid);
        end
        warning('보고서 생성 실패: %s', ME.message);
    end
end