%% 역량진단 데이터 표준화 전후 비교 분석
%
% 목적: 기존 corr_item_vs_comp_score.m 코드를 기반으로
%       표준화 전후 비교 분석 및 분포 분석 수행
%
% 작성일: 2025년
% 저장위치: D:\project\HR데이터\matlab\문항기반\역진 데이터 확인

%% ==================== 메인 실행 스크립트 ====================
clear; clc; close all;
rng(42, 'twister'); % 재현성 보장

%% 1. 초기 설정 및 파라미터
fprintf('========================================\n');
fprintf('역량진단 데이터 표준화 전후 비교 분석\n');
fprintf('========================================\n\n');

config = struct();
config.dataPath = 'D:\project\HR데이터\matlab\';
config.outputPath = 'D:\project\HR데이터\matlab\문항기반\역진 데이터 확인\결과\';
config.standardizeMethods = {'none', 'minmax', 'zscore', 'percentage'};
config.corrThreshold = 0.3;
config.pValueThreshold = 0.05;

% 결과 디렉토리 생성
if ~exist(config.outputPath, 'dir')
    mkdir(config.outputPath);
end

fprintf('[1단계] 초기 설정 완료\n');
fprintf('  - 출력 경로: %s\n', config.outputPath);
fprintf('  - 표준화 방법: %s\n', strjoin(config.standardizeMethods, ', '));

%% 2. 데이터 로드
fprintf('\n[2단계] 데이터 로드\n');
fprintf('----------------------------------------\n');

try
    [rawData, metadata] = loadCompetencyData(config.dataPath);
    fprintf('✓ 데이터 로드 완료\n');
    fprintf('  - 전체 데이터 크기: %dx%d\n', size(rawData.allScores));
    fprintf('  - Period 수: %d개\n', length(metadata.periods));
catch ME
    fprintf('✗ 데이터 로드 실패: %s\n', ME.message);
    return;
end

%% 3. 표준화 전후 비교 분석
fprintf('\n[3단계] 표준화 전후 비교 분석\n');
fprintf('----------------------------------------\n');

standardizationResults = struct();
for i = 1:length(config.standardizeMethods)
    method = config.standardizeMethods{i};
    fprintf('  처리 중: %s 방법...\n', method);

    try
        standardizationResults.(method) = analyzeWithStandardization(rawData, metadata, method);
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
fprintf('\n[4단계] 분포 분석\n');
fprintf('----------------------------------------\n');

try
    distributionStats = performDistributionAnalysis(rawData, metadata);
    fprintf('✓ 분포 분석 완료\n');
    fprintf('  - 평균: %.3f\n', distributionStats.mean);
    fprintf('  - 표준편차: %.3f\n', distributionStats.std);
    fprintf('  - 이상치 개수: %d개\n', sum(distributionStats.outliers));
catch ME
    fprintf('✗ 분포 분석 실패: %s\n', ME.message);
end

%% 5. 시각화 및 결과 저장
fprintf('\n[5단계] 시각화 및 결과 저장\n');
fprintf('----------------------------------------\n');

try
    createVisualizationsAndReports(standardizationResults, impactAnalysis, distributionStats, config);
    fprintf('✓ 결과 저장 완료\n');
catch ME
    fprintf('✗ 결과 저장 실패: %s\n', ME.message);
end

fprintf('\n========================================\n');
fprintf('분석 완료. 결과가 %s에 저장되었습니다.\n', config.outputPath);
fprintf('========================================\n');

%% ==================== 데이터 처리 함수 ====================

function [data, metadata] = loadCompetencyData(dataPath)
% LOADCOMPETENCYDATA - 역량진단 데이터 로드
%   Input: dataPath - 데이터 경로
%   Output: data - 원시 데이터, metadata - 메타정보

    arguments
        dataPath (1,:) char
    end

    data = struct();
    metadata = struct();

    % 기존 MAT 파일 찾기
    oldDir = pwd;
    cd(dataPath);

    try
        matFiles = dir('competency_correlation_workspace_*.mat');
        if ~isempty(matFiles)
            [~, idx] = max([matFiles.datenum]);
            matFileName = matFiles(idx).name;

            loadedData = load(matFileName);

            if isfield(loadedData, 'allData')
                allData = loadedData.allData;
                periods = fieldnames(allData);

                % 모든 Period 데이터 통합
                allScores = [];
                allIDs = {};
                periodLabels = {};

                for p = 1:length(periods)
                    periodData = allData.(periods{p});
                    if isfield(periodData, 'selfData') && ~isempty(periodData.selfData)
                        % ID 컬럼 제외하고 숫자형 컬럼만 추출
                        allColNames = periodData.selfData.Properties.VariableNames;
                        numericCols = varfun(@isnumeric, periodData.selfData, 'output', 'uniform');

                        % ID 관련 컬럼 제외
                        idCols = contains(allColNames, {'id', 'ID', 'idhr'}, 'IgnoreCase', true);
                        validCols = numericCols & ~idCols;

                        numericData = periodData.selfData{:, validCols};

                        allScores = [allScores; numericData];

                        % ID 정보 (있다면)
                        if ismember('ID', periodData.selfData.Properties.VariableNames)
                            periodIDs = periodData.selfData.ID;
                        else
                            periodIDs = cellstr(string(1:height(periodData.selfData))');
                        end
                        allIDs = [allIDs; periodIDs];
                        periodLabels = [periodLabels; repmat({periods{p}}, height(periodData.selfData), 1)];
                    end
                end

                data.allScores = allScores;
                data.allIDs = allIDs;
                data.periodLabels = periodLabels;
                metadata.periods = periods;
                metadata.source = 'MAT_file';

            else
                error('allData를 찾을 수 없습니다.');
            end
        else
            error('MAT 파일을 찾을 수 없습니다.');
        end

    catch ME
        error('데이터 로드 중 오류: %s', ME.message);
    finally
        cd(oldDir);
    end
end

function result = analyzeWithStandardization(data, metadata, method)
% ANALYZEWITHSTANDARDIZATION - 지정된 방법으로 표준화 후 분석
%   표준화 전 원본 데이터 저장
%   표준화 후 데이터와 함께 반환

    arguments
        data struct
        metadata struct
        method (1,:) char {mustBeMember(method, {'zscore', 'minmax', 'percentage', 'none'})} = 'zscore'
    end

    result = struct();
    result.method = method;
    result.originalData = data.allScores;  % 원본 저장

    if isempty(data.allScores)
        warning('빈 데이터입니다.');
        result.standardizedData = [];
        result.correlation = [];
        result.corrValue = [];
        result.pValue = [];
        return;
    end

    % 표준화 수행
    switch method
        case 'none'
            result.standardizedData = data.allScores;

        case 'minmax'
            minVals = min(data.allScores, [], 1, 'omitnan');
            maxVals = max(data.allScores, [], 1, 'omitnan');
            range = maxVals - minVals;
            range(range == 0) = 1; % 0으로 나누기 방지
            result.standardizedData = (data.allScores - minVals) ./ range;

        case 'zscore'
            meanVals = mean(data.allScores, 1, 'omitnan');
            stdVals = std(data.allScores, 0, 1, 'omitnan');
            stdVals(stdVals == 0) = 1; % 0으로 나누기 방지
            result.standardizedData = (data.allScores - meanVals) ./ stdVals;

        case 'percentage'
            maxVals = max(data.allScores, [], 1, 'omitnan');
            maxVals(maxVals == 0) = 1; % 0으로 나누기 방지
            result.standardizedData = (data.allScores ./ maxVals) * 100;
    end

    % 상관분석 수행
    try
        validData = result.standardizedData;
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
% ANALYZESTANDARDIZATIONIMPACT - 표준화 전후 비교 메인 함수
%   표준화 방법별 영향 종합 분석

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
        else
            warning('상관 매트릭스 크기가 다릅니다: %s', methods{i});
        end
    end
end

function rankChange = calculateRankChange(baseline, method)
% CALCULATERANKCHANGE - 상관계수 순위 변화 계산

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

function stats = performDistributionAnalysis(data, metadata)
% PERFORMDISTRIBUTIONANALYSIS - 상세 분포 분석

    arguments
        data struct
        metadata struct
    end

    if isempty(data.allScores)
        stats = struct();
        return;
    end

    % 기술통계 계산
    stats = calculateDescriptiveStats(data.allScores);

    % 이상치 탐지
    stats.outliers = detectOutliers(data.allScores, 'iqr');

    % 정규성 검정
    stats.normality = performNormalityTest(data.allScores);

    % 시점별 분석 (Period별 분석)
    if isfield(data, 'periodLabels') && ~isempty(data.periodLabels)
        stats.timeSeriesAnalysis = analyzeByPeriod(data.allScores, data.periodLabels);
    end
end

function stats = calculateDescriptiveStats(data)
% CALCULATEDESCRIPTIVESTATS - 기술통계 계산
%   평균, 중앙값, 표준편차, 사분위수, 왜도, 첨도

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
% DETECTOUTLIERS - 이상치 탐지
%   method: 'iqr', 'zscore'

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
% PERFORMNORMALITYTEST - 정규성 검정 (Kolmogorov-Smirnov test)

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
% ANALYZEBYPERIOD - 시점별 분석

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

function createVisualizationsAndReports(standardizationResults, impactAnalysis, distributionStats, config)
% CREATEVISUALIZATIONSANDREPORTS - 종합 시각화 및 보고서 생성

    arguments
        standardizationResults struct
        impactAnalysis struct
        distributionStats struct
        config struct
    end

    try
        % 1. 종합 시각화 생성
        createMainVisualization(standardizationResults, impactAnalysis, distributionStats, config);

        % 2. Excel 결과 저장
        exportToExcel(standardizationResults, impactAnalysis, distributionStats, config);

        % 3. 종합 보고서 생성
        createDistributionReport(distributionStats, config);

    catch ME
        warning('시각화 생성 중 오류: %s', ME.message);
    end
end

function createMainVisualization(standardizationResults, impactAnalysis, distributionStats, config)
% CREATEMAINVISUALIZATION - 메인 시각화 생성

    % 메인 Figure 생성
    fig = figure('Position', [100, 100, 1400, 900], 'Color', 'white');

    % 1. 표준화 방법별 히스토그램 비교 (2x2 subplot)
    methods = fieldnames(standardizationResults);
    validMethods = {};

    for i = 1:length(methods)
        if ~isempty(standardizationResults.(methods{i}).standardizedData)
            validMethods{end+1} = methods{i};
        end
    end

    if length(validMethods) >= 4
        for i = 1:4
            subplot(2, 3, i);
            method = validMethods{i};
            data = standardizationResults.(method).standardizedData;

            if ~isempty(data)
                allValues = data(:);
                allValues = allValues(~isnan(allValues));

                if ~isempty(allValues)
                    histogram(allValues, 30, 'FaceColor', [0.2, 0.6, 0.8], 'EdgeColor', 'none', 'FaceAlpha', 0.7);
                    title(sprintf('%s 분포', upper(method)), 'FontSize', 12);
                    xlabel('값');
                    ylabel('빈도');
                    grid on;
                end
            end
        end
    end

    % 5. 상관계수 변화량 막대그래프
    subplot(2, 3, 5);
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
        title('표준화 방법별 평균 상관계수 변화량');
        ylabel('평균 절대 변화량');
        grid on;
    end

    % 6. 기술통계 요약
    subplot(2, 3, 6);
    if ~isempty(fieldnames(distributionStats))
        stats = [distributionStats.mean, distributionStats.median, distributionStats.std, ...
                distributionStats.skewness, distributionStats.kurtosis];
        statNames = {'평균', '중앙값', '표준편차', '왜도', '첨도'};

        bar(stats, 'FaceColor', [0.4, 0.7, 0.4]);
        set(gca, 'XTickLabel', statNames);
        title('분포 기술통계');
        ylabel('값');
        grid on;
        xtickangle(45);
    end

    % 전체 제목 추가
    sgtitle('역량진단 데이터 표준화 전후 비교 분석 결과', 'FontSize', 16, 'FontWeight', 'bold');

    % 그림 저장
    saveas(fig, fullfile(config.outputPath, '표준화영향분석_시각화.fig'));
    saveas(fig, fullfile(config.outputPath, '표준화영향분석_시각화.png'), 'png');

    fprintf('  ✓ 시각화 저장: %s\n', fullfile(config.outputPath, '표준화영향분석_시각화.png'));
end

function exportToExcel(standardizationResults, impactAnalysis, distributionStats, config)
% EXPORTTOEXCEL - Excel로 결과 내보내기

    filename = fullfile(config.outputPath, '표준화방법별_비교결과.xlsx');

    try
        % Sheet 1: 표준화 방법별 기술통계 비교
        statsComparison = createStatsComparisonTable(standardizationResults);
        if ~isempty(statsComparison)
            writetable(statsComparison, filename, 'Sheet', '기술통계비교', 'WriteMode', 'overwritesheet');
        end

        % Sheet 2: 영향 분석 결과
        if ~isempty(fieldnames(impactAnalysis))
            impactTable = createImpactAnalysisTable(impactAnalysis);
            if ~isempty(impactTable)
                writetable(impactTable, filename, 'Sheet', '영향분석', 'WriteMode', 'overwritesheet');
            end
        end

        % Sheet 3: 분포 통계
        if ~isempty(fieldnames(distributionStats))
            statsTable = createDistributionStatsTable(distributionStats);
            if ~isempty(statsTable)
                writetable(statsTable, filename, 'Sheet', '분포통계', 'WriteMode', 'overwritesheet');
            end
        end

        fprintf('  ✓ Excel 파일 저장: %s\n', filename);

    catch ME
        warning('Excel 저장 실패: %s', ME.message);
    end
end

function statsTable = createStatsComparisonTable(standardizationResults)
% CREATESTATSCOMPARISONTABLE - 기술통계 비교 테이블 생성

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

function impactTable = createImpactAnalysisTable(impactAnalysis)
% CREATEIMPACTANALYSISTABLE - 영향 분석 테이블 생성

    methods = fieldnames(impactAnalysis);

    methodNames = {};
    meanAbsChanges = [];
    sigCountChanges = [];
    rankChanges = [];

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
    end

    if ~isempty(methodNames)
        impactTable = table(methodNames', meanAbsChanges', sigCountChanges', rankChanges', ...
            'VariableNames', {'표준화방법', '평균절대변화량', '유의상관개수변화', '순위상관계수'});
    else
        impactTable = table();
    end
end

function statsTable = createDistributionStatsTable(distributionStats)
% CREATEDISTRIBUTIONSTABLE - 분포 통계 테이블 생성

    if isempty(fieldnames(distributionStats))
        statsTable = table();
        return;
    end

    statNames = {'평균', '중앙값', '표준편차', '사분위수_Q1', '사분위수_Q2', '사분위수_Q3', ...
                 '왜도', '첨도', 'IQR', '이상치개수', '정규성_p값', '정규성_만족여부'};

    statValues = [distributionStats.mean, distributionStats.median, distributionStats.std, ...
                  distributionStats.quartiles(1), distributionStats.quartiles(2), distributionStats.quartiles(3), ...
                  distributionStats.skewness, distributionStats.kurtosis, distributionStats.iqr, ...
                  sum(distributionStats.outliers(:)), ...
                  distributionStats.normality.pValue, distributionStats.normality.isNormal];

    statsTable = table(statNames', statValues', ...
        'VariableNames', {'통계량', '값'});
end

function createDistributionReport(stats, config)
% CREATEDISTRIBUTIONREPORT - 종합 분포 분석 보고서 생성

    reportFile = fullfile(config.outputPath, '분포분석보고서.txt');

    try
        fid = fopen(reportFile, 'w', 'n', 'UTF-8');

        fprintf(fid, '========== 역량진단 데이터 분포 분석 보고서 ==========\n\n');
        fprintf(fid, '분석 일시: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

        if ~isempty(fieldnames(stats))
            fprintf(fid, '1. 기술통계\n');
            fprintf(fid, '   - 평균: %.3f\n', stats.mean);
            fprintf(fid, '   - 중앙값: %.3f\n', stats.median);
            fprintf(fid, '   - 표준편차: %.3f\n', stats.std);
            fprintf(fid, '   - 사분위수: [%.3f, %.3f, %.3f]\n', stats.quartiles(1), stats.quartiles(2), stats.quartiles(3));
            fprintf(fid, '   - 왜도: %.3f', stats.skewness);
            if stats.skewness > 1
                fprintf(fid, ' (강한 우편향)');
            elseif stats.skewness > 0.5
                fprintf(fid, ' (약한 우편향)');
            elseif stats.skewness < -1
                fprintf(fid, ' (강한 좌편향)');
            elseif stats.skewness < -0.5
                fprintf(fid, ' (약한 좌편향)');
            else
                fprintf(fid, ' (대칭)');
            end
            fprintf(fid, '\n');

            fprintf(fid, '   - 첨도: %.3f', stats.kurtosis);
            if stats.kurtosis > 3
                fprintf(fid, ' (뾰족한 분포)');
            elseif stats.kurtosis < 3
                fprintf(fid, ' (평평한 분포)');
            else
                fprintf(fid, ' (정규분포 수준)');
            end
            fprintf(fid, '\n');

            fprintf(fid, '\n2. 이상치 분석\n');
            fprintf(fid, '   - 탐지된 이상치 개수: %d\n', sum(stats.outliers(:)));
            if numel(stats.outliers) > 0
                fprintf(fid, '   - 전체 대비 비율: %.2f%%\n', 100*sum(stats.outliers(:))/numel(stats.outliers));
            end

            if isfield(stats, 'normality')
                fprintf(fid, '\n3. 정규성 검정\n');
                fprintf(fid, '   - 검정 방법: %s\n', stats.normality.testType);
                fprintf(fid, '   - p-value: %.4f\n', stats.normality.pValue);
                if stats.normality.isNormal
                    fprintf(fid, '   - 정규성 만족 (α = 0.05 기준)\n');
                else
                    fprintf(fid, '   - 정규성 불만족 (α = 0.05 기준)\n');
                end
            end

            if isfield(stats, 'timeSeriesAnalysis')
                fprintf(fid, '\n4. 시점별 분석\n');
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