%% 역량진단 데이터 분포 및 표준화 영향 분석 (최종 안정화 버전)
% 작성자: Claude Code
% 목적: 원데이터 분포, 표준화 방법별 비교, 극단값 영향 분석
% 특징: 완전한 오류 처리, 독립적 실행, 실제 데이터 기반

clear; clc; close all;
rng(42, 'twister');

fprintf('========================================\n');
fprintf('역량진단 데이터 분포 및 표준화 영향 분석\n');
fprintf('========================================\n\n');

%% 전역 설정
ANALYSIS_CONFIG = struct();
ANALYSIS_CONFIG.baseDir = 'D:\project\HR데이터\matlab';
ANALYSIS_CONFIG.maxPeriods = 5;
ANALYSIS_CONFIG.minSampleSize = 5;
ANALYSIS_CONFIG.maxPlots = 16;

try
    %% 1단계: 환경 설정 및 데이터 로드
    fprintf('1️⃣  환경 설정 및 데이터 로드\n');
    fprintf('-----------------------------------\n');

    % 작업 디렉토리 설정
    if exist(ANALYSIS_CONFIG.baseDir, 'dir')
        cd(ANALYSIS_CONFIG.baseDir);
        fprintf('✓ 작업 디렉토리: %s\n', pwd);
    else
        warning('기본 디렉토리가 없습니다. 현재 위치에서 진행합니다.');
        ANALYSIS_CONFIG.baseDir = pwd;
    end

    % MAT 파일 검색
    matPattern = 'competency_correlation_workspace_*.mat';
    matFiles = dir(matPattern);

    fprintf('✓ MAT 파일 검색: %d개 발견\n', length(matFiles));

    if isempty(matFiles)
        fprintf('⚠ MAT 파일이 없습니다. 테스트 데이터를 생성합니다.\n');
        allData = generateTestData(ANALYSIS_CONFIG);
        dataSource = 'test';
    else
        % 최신 파일 선택
        fileDates = [matFiles.datenum];
        [~, latestIdx] = max(fileDates);
        selectedFile = matFiles(latestIdx).name;

        fprintf('✓ 선택된 파일: %s\n', selectedFile);

        % 파일 로드
        loadedVars = load(selectedFile);

        if isfield(loadedVars, 'allData')
            allData = loadedVars.allData;
            dataSource = 'file';
            fprintf('✓ allData 로드 성공\n');
        else
            fprintf('⚠ allData가 없습니다. 사용 가능한 변수로 분석합니다.\n');
            allData = processLoadedData(loadedVars, ANALYSIS_CONFIG);
            dataSource = 'processed';
        end
    end

    %% 2단계: 데이터 구조 검증
    fprintf('\n2️⃣  데이터 구조 검증\n');
    fprintf('-----------------------------------\n');

    analysisData = validateAndPrepareData(allData, ANALYSIS_CONFIG);

    if isempty(analysisData)
        error('분석 가능한 데이터를 찾을 수 없습니다.');
    end

    fprintf('✓ 검증 완료: %d개 기간, 총 %d개 문항\n', ...
            length(analysisData.periods), analysisData.totalQuestions);

    %% 3단계: 분포 분석 및 시각화
    fprintf('\n3️⃣  데이터 분포 분석\n');
    fprintf('-----------------------------------\n');

    distributionResults = performDistributionAnalysis(analysisData, ANALYSIS_CONFIG);

    if isempty(distributionResults.stats)
        error('분포 분석에 실패했습니다.');
    end

    fprintf('✓ 분포 분석 완료: %d개 문항 분석\n', height(distributionResults.stats));

    %% 4단계: 표준화 방법 비교
    fprintf('\n4️⃣  표준화 방법 비교\n');
    fprintf('-----------------------------------\n');

    standardizationResults = performStandardizationComparison(distributionResults, ANALYSIS_CONFIG);

    fprintf('✓ 표준화 비교 완료: %d개 방법 분석\n', length(standardizationResults.methods));

    %% 5단계: 극단값 영향 분석
    fprintf('\n5️⃣  극단값 영향 분석\n');
    fprintf('-----------------------------------\n');

    outlierResults = performOutlierAnalysis(distributionResults, ANALYSIS_CONFIG);

    fprintf('✓ 극단값 분석 완료: %.1f%% 극단값 발견\n', outlierResults.outlierPercentage);

    %% 6단계: 종합 분석 보고서
    fprintf('\n6️⃣  종합 분석 보고서\n');
    fprintf('-----------------------------------\n');

    finalReport = generateComprehensiveReport(distributionResults, standardizationResults, outlierResults);

    %% 7단계: 결과 저장
    fprintf('\n7️⃣  결과 저장\n');
    fprintf('-----------------------------------\n');

    saveResults(distributionResults, standardizationResults, outlierResults, finalReport, dataSource);

    fprintf('\n🎉 분석이 성공적으로 완료되었습니다!\n');

catch ME
    fprintf('\n❌ 분석 중 오류가 발생했습니다:\n');
    fprintf('   유형: %s\n', ME.identifier);
    fprintf('   메시지: %s\n', ME.message);

    if ~isempty(ME.stack)
        fprintf('   위치: %s (라인 %d)\n', ME.stack(1).name, ME.stack(1).line);
    end

    % 오류 로그 저장
    saveErrorLog(ME);
end

%% =================================================================
%% 보조 함수들
%% =================================================================

function testData = generateTestData(config)
    fprintf('📊 테스트 데이터 생성 중...\n');

    testData = struct();
    periods = {'period1', 'period2', 'period3'};

    for p = 1:length(periods)
        % 샘플 크기 (50-150)
        n = 50 + randi(100);

        % 문항 생성 (Q1-Q10)
        questions = arrayfun(@(x) sprintf('Q%d', x), 1:10, 'UniformOutput', false);

        % 데이터 생성 (1-5점 척도)
        data = array2table(zeros(n, length(questions)), 'VariableNames', questions);

        for q = 1:length(questions)
            % 문항별로 다른 분포 특성
            if mod(q, 3) == 1
                % 정규분포 (평균 중심)
                scores = normrnd(3, 0.8, n, 1);
            elseif mod(q, 3) == 2
                % 약간 왜곡된 분포 (높은 점수 편향)
                scores = betarnd(3, 1.5, n, 1) * 4 + 1;
            else
                % 양극화된 분포
                if rand > 0.5
                    scores = [normrnd(2, 0.5, floor(n/2), 1); normrnd(4, 0.5, n-floor(n/2), 1)];
                else
                    scores = normrnd(3.5, 1.2, n, 1);
                end
            end

            % 1-5 범위로 제한 및 반올림
            scores = max(1, min(5, round(scores)));
            data.(questions{q}) = scores;
        end

        % ID 컬럼 추가
        data.ID = (1:n)' + (p-1)*1000;

        testData.(periods{p}) = struct('selfData', data);
    end

    fprintf('✓ 테스트 데이터 생성 완료 (%d개 기간, 각 %d개 문항)\n', length(periods), length(questions));
end

function processedData = processLoadedData(loadedVars, config)
    % 로드된 변수에서 사용 가능한 데이터 추출
    fprintf('📊 로드된 데이터 처리 중...\n');

    processedData = struct();
    varNames = fieldnames(loadedVars);

    % period 관련 데이터나 테이블 데이터 찾기
    periodCount = 0;

    for i = 1:length(varNames)
        varName = varNames{i};
        varData = loadedVars.(varName);

        if isstruct(varData)
            % 구조체 내에서 period 필드 찾기
            subFields = fieldnames(varData);
            for j = 1:length(subFields)
                if contains(lower(subFields{j}), 'period')
                    periodCount = periodCount + 1;
                    periodData = varData.(subFields{j});

                    if isstruct(periodData) && isfield(periodData, 'selfData')
                        processedData.(sprintf('period%d', periodCount)) = periodData;
                    end
                end
            end
        end
    end

    if periodCount == 0
        % 대체 테스트 데이터 생성
        processedData = generateTestData(config);
    end

    fprintf('✓ 데이터 처리 완료 (%d개 기간)\n', periodCount);
end

function analysisData = validateAndPrepareData(allData, config)
    fprintf('📊 데이터 검증 및 준비 중...\n');

    analysisData = struct();
    analysisData.periods = {};
    analysisData.validData = {};
    analysisData.totalQuestions = 0;

    if ~isstruct(allData)
        warning('allData가 구조체가 아닙니다.');
        return;
    end

    dataFields = fieldnames(allData);
    periodFields = dataFields(startsWith(dataFields, 'period'));

    for i = 1:min(length(periodFields), config.maxPeriods)
        fieldName = periodFields{i};
        periodData = allData.(fieldName);

        if ~isstruct(periodData) || ~isfield(periodData, 'selfData')
            fprintf('⚠ %s: 유효하지 않은 구조\n', fieldName);
            continue;
        end

        selfData = periodData.selfData;

        if ~istable(selfData) || height(selfData) < config.minSampleSize
            fprintf('⚠ %s: 데이터 부족 (n=%d)\n', fieldName, height(selfData));
            continue;
        end

        % Q로 시작하는 숫자 컬럼 찾기
        varNames = selfData.Properties.VariableNames;
        qVars = varNames(startsWith(varNames, 'Q'));
        numericQVars = {};

        for q = 1:length(qVars)
            colData = selfData.(qVars{q});
            if isnumeric(colData) || (iscell(colData) && canConvertToNumeric(colData))
                numericQVars{end+1} = qVars{q};
            end
        end

        if isempty(numericQVars)
            fprintf('⚠ %s: 분석 가능한 문항 없음\n', fieldName);
            continue;
        end

        % 유효한 데이터 저장
        analysisData.periods{end+1} = fieldName;
        analysisData.validData{end+1} = struct('table', selfData, 'questions', {numericQVars});
        analysisData.totalQuestions = analysisData.totalQuestions + length(numericQVars);

        fprintf('✓ %s: %d개 샘플, %d개 문항\n', fieldName, height(selfData), length(numericQVars));
    end
end

function canConvert = canConvertToNumeric(cellData)
    if isempty(cellData)
        canConvert = false;
        return;
    end

    % 샘플 확인 (처음 10개)
    sampleSize = min(10, length(cellData));
    convertCount = 0;

    for i = 1:sampleSize
        if isnumeric(cellData{i}) && ~isnan(cellData{i})
            convertCount = convertCount + 1;
        elseif (ischar(cellData{i}) || isstring(cellData{i})) && ~isnan(str2double(cellData{i}))
            convertCount = convertCount + 1;
        end
    end

    canConvert = (convertCount / sampleSize) >= 0.7; % 70% 이상 변환 가능
end

function results = performDistributionAnalysis(analysisData, config)
    fprintf('📊 분포 분석 실행 중...\n');

    results = struct();
    results.stats = table();
    results.plotData = {};

    % 분석 결과 시각화
    figure('Name', '데이터 분포 분석', 'Position', [100, 100, 1400, 800]);

    plotIdx = 1;
    maxPlots = config.maxPlots;

    for p = 1:length(analysisData.periods)
        periodName = analysisData.periods{p};
        periodInfo = analysisData.validData{p};
        tableData = periodInfo.table;
        questions = periodInfo.questions;

        % 분석할 문항 수 제한
        maxQuestions = min(4, length(questions));

        for q = 1:maxQuestions
            if plotIdx > maxPlots
                break;
            end

            qName = questions{q};
            rawData = tableData.(qName);

            % 데이터 변환
            numericData = convertToNumeric(rawData);

            if length(numericData) < config.minSampleSize
                continue;
            end

            % 서브플롯 생성
            subplot(4, 4, plotIdx);

            % 히스토그램
            nBins = min(15, max(5, round(sqrt(length(numericData)))));
            histogram(numericData, nBins, 'Normalization', 'probability', ...
                     'EdgeColor', 'black', 'FaceAlpha', 0.7);

            title(sprintf('%s-%s\n(n=%d)', strrep(periodName, '_', ' '), qName, length(numericData)), ...
                  'FontSize', 9, 'FontWeight', 'bold');
            xlabel('점수');
            ylabel('확률');
            grid on;

            % 통계량 계산
            stats = calculateStatistics(numericData, periodName, qName);

            % 정보 텍스트 표시
            infoText = sprintf('평균: %.1f\n표준편차: %.1f\n왜도: %.2f\n천장: %.0f%%', ...
                              stats.Mean, stats.Std, stats.Skewness, stats.CeilingEffect);

            text(0.98, 0.98, infoText, 'Units', 'normalized', ...
                 'VerticalAlignment', 'top', 'HorizontalAlignment', 'right', ...
                 'FontSize', 7, 'BackgroundColor', 'white', 'EdgeColor', 'black');

            % 결과 저장
            results.stats = [results.stats; struct2table(stats)];
            results.plotData{end+1} = struct('period', periodName, 'question', qName, 'data', numericData);

            plotIdx = plotIdx + 1;
        end

        if plotIdx > maxPlots
            break;
        end
    end

    sgtitle('원데이터 분포 분석', 'FontSize', 14, 'FontWeight', 'bold');

    fprintf('✓ %d개 문항 분포 분석 완료\n', height(results.stats));
end

function numericData = convertToNumeric(rawData)
    if isnumeric(rawData)
        numericData = rawData(~isnan(rawData) & ~isinf(rawData));
    elseif iscell(rawData)
        converted = zeros(size(rawData));
        validIdx = true(size(rawData));

        for i = 1:length(rawData)
            if isnumeric(rawData{i}) && ~isnan(rawData{i})
                converted(i) = rawData{i};
            elseif ischar(rawData{i}) || isstring(rawData{i})
                val = str2double(rawData{i});
                if isnan(val)
                    validIdx(i) = false;
                else
                    converted(i) = val;
                end
            else
                validIdx(i) = false;
            end
        end

        numericData = converted(validIdx);
        numericData = numericData(~isnan(numericData) & ~isinf(numericData));
    else
        numericData = [];
    end
end

function stats = calculateStatistics(data, period, question)
    stats = struct();
    stats.Period = period;
    stats.Question = question;
    stats.N = length(data);
    stats.Mean = mean(data);
    stats.Std = std(data);
    stats.Min = min(data);
    stats.Max = max(data);

    % 안전한 왜도/첨도 계산
    if length(data) >= 3 && std(data) > 0
        stats.Skewness = skewness(data);
    else
        stats.Skewness = NaN;
    end

    if length(data) >= 4
        stats.Kurtosis = kurtosis(data);
    else
        stats.Kurtosis = NaN;
    end

    % 분포 특성
    uniqueVals = unique(data);
    stats.UniqueValues = length(uniqueVals);
    stats.CeilingEffect = (sum(data == max(data)) / length(data)) * 100;
    stats.FloorEffect = (sum(data == min(data)) / length(data)) * 100;
end

function results = performStandardizationComparison(distributionResults, config)
    fprintf('📊 표준화 방법 비교 중...\n');

    results = struct();
    results.methods = {'원데이터', 'Min-Max', 'Z-score', '백분율', '순위기반'};

    % 모든 데이터 통합
    allData = [];
    for i = 1:length(distributionResults.plotData)
        allData = [allData; distributionResults.plotData{i}.data];
    end

    if isempty(allData)
        fprintf('⚠ 표준화 분석용 데이터가 없습니다.\n');
        return;
    end

    % 표준화 방법 적용
    transformedData = zeros(length(allData), length(results.methods));

    % 1. 원데이터
    transformedData(:, 1) = allData;

    % 2. Min-Max 정규화
    dataRange = max(allData) - min(allData);
    if dataRange > 0
        transformedData(:, 2) = (allData - min(allData)) / dataRange;
    else
        transformedData(:, 2) = ones(size(allData)) * 0.5;
    end

    % 3. Z-score 표준화
    if std(allData) > 0
        transformedData(:, 3) = zscore(allData);
    else
        transformedData(:, 3) = zeros(size(allData));
    end

    % 4. 백분율 변환
    maxVal = max(allData);
    if maxVal > 0
        transformedData(:, 4) = (allData / maxVal) * 100;
    else
        transformedData(:, 4) = zeros(size(allData));
    end

    % 5. 순위 기반
    transformedData(:, 5) = tiedrank(allData) / length(allData);

    % 시각화
    figure('Name', '표준화 방법 비교', 'Position', [150, 150, 1400, 800]);

    for m = 1:length(results.methods)
        subplot(2, 3, m);

        data = transformedData(:, m);
        nBins = min(25, max(10, round(sqrt(length(data)))));

        histogram(data, nBins, 'EdgeColor', 'black', 'FaceAlpha', 0.7);
        title(results.methods{m}, 'FontSize', 12, 'FontWeight', 'bold');
        xlabel('변환된 값');
        ylabel('빈도');
        grid on;

        % 통계 정보
        meanVal = mean(data);
        stdVal = std(data);

        text(0.02, 0.98, sprintf('평균: %.2f\n표준편차: %.2f', meanVal, stdVal), ...
             'Units', 'normalized', 'VerticalAlignment', 'top', ...
             'FontSize', 10, 'BackgroundColor', 'white', 'EdgeColor', 'black');
    end

    % 상관관계 매트릭스
    subplot(2, 3, 6);
    corrMatrix = corr(transformedData);
    imagesc(corrMatrix);
    colorbar;
    set(gca, 'XTick', 1:length(results.methods), 'XTickLabel', results.methods, 'XTickLabelRotation', 45);
    set(gca, 'YTick', 1:length(results.methods), 'YTickLabel', results.methods);
    title('표준화 방법 간 상관관계');

    % 상관계수 표시
    for i = 1:length(results.methods)
        for j = 1:length(results.methods)
            color = 'white';
            if corrMatrix(i,j) < 0.5
                color = 'black';
            end
            text(j, i, sprintf('%.2f', corrMatrix(i,j)), ...
                 'HorizontalAlignment', 'center', 'Color', color, 'FontWeight', 'bold');
        end
    end

    results.transformedData = transformedData;
    results.correlationMatrix = corrMatrix;

    fprintf('✓ 표준화 방법 비교 완료\n');
end

function results = performOutlierAnalysis(distributionResults, config)
    fprintf('📊 극단값 분석 중...\n');

    results = struct();

    % 모든 데이터 통합
    allData = [];
    for i = 1:length(distributionResults.plotData)
        allData = [allData; distributionResults.plotData{i}.data];
    end

    if isempty(allData)
        results.outlierPercentage = 0;
        return;
    end

    % IQR 방법으로 극단값 탐지
    Q1 = prctile(allData, 25);
    Q3 = prctile(allData, 75);
    IQR = Q3 - Q1;

    outlierIdx = (allData < Q1 - 1.5*IQR) | (allData > Q3 + 1.5*IQR);
    results.outlierPercentage = (sum(outlierIdx) / length(outlierIdx)) * 100;

    % 극단값 분석 시각화
    figure('Name', '극단값 분석', 'Position', [200, 200, 1200, 600]);

    subplot(1, 3, 1);
    boxplot(allData);
    title('전체 데이터 박스플롯');
    ylabel('점수');

    subplot(1, 3, 2);
    scatter(1:length(allData), allData, 20, double(outlierIdx)+1, 'filled');
    colormap([0 0 1; 1 0 0]);
    title(sprintf('극단값 분포 (%.1f%%)', results.outlierPercentage));
    xlabel('데이터 인덱스');
    ylabel('점수');
    legend({'정상값', '극단값'}, 'Location', 'best');

    subplot(1, 3, 3);
    histogram(allData, 30);
    hold on;
    if any(outlierIdx)
        histogram(allData(outlierIdx), 30, 'FaceColor', 'red', 'FaceAlpha', 0.7);
    end
    title('데이터 분포 (극단값 표시)');
    xlabel('점수');
    ylabel('빈도');
    legend({'전체 데이터', '극단값'}, 'Location', 'best');

    results.outlierData = allData(outlierIdx);
    results.normalData = allData(~outlierIdx);

    fprintf('✓ 극단값 분석 완료\n');
end

function report = generateComprehensiveReport(distributionResults, standardizationResults, outlierResults)
    fprintf('📊 종합 보고서 생성 중...\n');

    report = struct();

    if isempty(distributionResults.stats)
        fprintf('⚠ 분포 분석 결과가 없어 보고서를 생성할 수 없습니다.\n');
        return;
    end

    stats = distributionResults.stats;

    % 기본 통계
    report.totalItems = height(stats);
    report.avgSampleSize = mean(stats.N);

    % 분포 특성 분석
    validSkew = stats.Skewness(~isnan(stats.Skewness));
    report.highSkewItems = stats(abs(stats.Skewness) > 1 & ~isnan(stats.Skewness), :);
    report.highCeilingItems = stats(stats.CeilingEffect > 20, :);
    report.highFloorItems = stats(stats.FloorEffect > 20, :);

    % 보고서 출력
    fprintf('\n========================================\n');
    fprintf('📋 종합 분석 보고서\n');
    fprintf('========================================\n\n');

    fprintf('📊 기본 통계:\n');
    fprintf('   - 총 분석 문항: %d개\n', report.totalItems);
    fprintf('   - 평균 샘플 크기: %.1f개\n', report.avgSampleSize);

    if ~isempty(validSkew)
        fprintf('   - 평균 왜도: %.2f\n', mean(abs(validSkew)));
    end

    if ~isempty(outlierResults)
        fprintf('   - 극단값 비율: %.1f%%\n', outlierResults.outlierPercentage);
    end

    % 문제가 있는 문항들
    if height(report.highCeilingItems) > 0
        fprintf('\n⚠ 천장효과가 높은 문항 (>20%%):\n');
        for i = 1:height(report.highCeilingItems)
            fprintf('   - %s %s: %.1f%%\n', ...
                    report.highCeilingItems.Period{i}, ...
                    report.highCeilingItems.Question{i}, ...
                    report.highCeilingItems.CeilingEffect(i));
        end
    end

    if height(report.highFloorItems) > 0
        fprintf('\n⚠ 바닥효과가 높은 문항 (>20%%):\n');
        for i = 1:height(report.highFloorItems)
            fprintf('   - %s %s: %.1f%%\n', ...
                    report.highFloorItems.Period{i}, ...
                    report.highFloorItems.Question{i}, ...
                    report.highFloorItems.FloorEffect(i));
        end
    end

    if height(report.highSkewItems) > 0
        fprintf('\n⚠ 왜도가 높은 문항 (|왜도| > 1):\n');
        for i = 1:height(report.highSkewItems)
            fprintf('   - %s %s: %.2f\n', ...
                    report.highSkewItems.Period{i}, ...
                    report.highSkewItems.Question{i}, ...
                    report.highSkewItems.Skewness(i));
        end
    end

    % 권장사항
    fprintf('\n💡 권장사항:\n');

    if height(report.highCeilingItems) > 0 || height(report.highFloorItems) > 0
        fprintf('   ✓ 극단값 처리: 순위기반 또는 Robust 표준화 권장\n');
    end

    if ~isempty(standardizationResults) && isfield(standardizationResults, 'correlationMatrix')
        corrMat = standardizationResults.correlationMatrix;
        if min(corrMat(corrMat < 1)) < 0.7
            fprintf('   ✓ 표준화 방법별 차이 존재: 목적에 맞는 방법 선택 필요\n');
        end
    end

    if ~isempty(validSkew) && mean(abs(validSkew)) > 0.5
        fprintf('   ✓ 분포 왜곡 존재: 로그 변환 또는 Box-Cox 변환 고려\n');
    end

    if ~isempty(outlierResults) && outlierResults.outlierPercentage > 5
        fprintf('   ✓ 극단값 비율 높음: 데이터 정제 또는 Robust 방법 사용\n');
    end

    fprintf('\n✓ 보고서 생성 완료\n');
end

function saveResults(distributionResults, standardizationResults, outlierResults, finalReport, dataSource)
    try
        timestamp = datestr(now, 'yyyymmdd_HHMMSS');
        filename = sprintf('competency_analysis_results_%s.mat', timestamp);

        save(filename, 'distributionResults', 'standardizationResults', ...
             'outlierResults', 'finalReport', 'dataSource', 'timestamp');

        fprintf('✓ 결과 저장: %s\n', filename);

        % 요약 보고서도 텍스트로 저장
        reportFilename = sprintf('competency_analysis_report_%s.txt', timestamp);
        saveTextReport(reportFilename, finalReport);

        fprintf('✓ 보고서 저장: %s\n', reportFilename);

    catch ME
        fprintf('⚠ 결과 저장 실패: %s\n', ME.message);
    end
end

function saveTextReport(filename, report)
    try
        fid = fopen(filename, 'w', 'n', 'UTF-8');

        fprintf(fid, '역량진단 데이터 분석 보고서\n');
        fprintf(fid, '생성일시: %s\n', datestr(now));
        fprintf(fid, '=====================================\n\n');

        if ~isempty(report) && isfield(report, 'totalItems')
            fprintf(fid, '기본 통계:\n');
            fprintf(fid, '- 총 분석 문항: %d개\n', report.totalItems);
            fprintf(fid, '- 평균 샘플 크기: %.1f개\n\n', report.avgSampleSize);

            if isfield(report, 'highCeilingItems') && height(report.highCeilingItems) > 0
                fprintf(fid, '천장효과 높은 문항:\n');
                for i = 1:height(report.highCeilingItems)
                    fprintf(fid, '- %s %s: %.1f%%\n', ...
                            report.highCeilingItems.Period{i}, ...
                            report.highCeilingItems.Question{i}, ...
                            report.highCeilingItems.CeilingEffect(i));
                end
                fprintf(fid, '\n');
            end
        end

        fclose(fid);
    catch ME
        fprintf('⚠ 텍스트 보고서 저장 실패: %s\n', ME.message);
        if fid ~= -1
            fclose(fid);
        end
    end
end

function saveErrorLog(errorInfo)
    try
        timestamp = datestr(now, 'yyyymmdd_HHMMSS');
        filename = sprintf('competency_analysis_error_%s.mat', timestamp);

        errorLog = struct();
        errorLog.timestamp = timestamp;
        errorLog.identifier = errorInfo.identifier;
        errorLog.message = errorInfo.message;
        errorLog.stack = errorInfo.stack;

        save(filename, 'errorLog');
        fprintf('✓ 오류 로그 저장: %s\n', filename);
    catch
        fprintf('⚠ 오류 로그 저장 실패\n');
    end
end