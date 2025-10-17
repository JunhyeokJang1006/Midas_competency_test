%% 표준화 영향 분석 코드
% 목적: 원데이터 분포, 표준화 전후 비교, 극단값 영향 확인

clear; clc; close all;

%% 1. 데이터 로드 (기존 코드 활용)
cd('D:\project\HR데이터\matlab')

% 기존 MAT 파일 로드
matFiles = dir('competency_correlation_workspace_*.mat');
if ~isempty(matFiles)
    [~, idx] = max([matFiles.datenum]);
    load(matFiles(idx).name);
end

%% 2. 척도별 문항 분포 분석
figure('Name', '척도별 원데이터 분포', 'Position', [100, 100, 1400, 800]);

% 각 시점별 분석
periods = {'23년_상반기', '23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
performanceQuestions = struct();
performanceQuestions.period1 = {'Q3', 'Q4', 'Q5', 'Q22', 'Q23', 'Q45', 'Q46', 'Q51'};
performanceQuestions.period2 = {'Q4', 'Q21', 'Q23', 'Q25', 'Q32', 'Q33', 'Q34'};
performanceQuestions.period3 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};
performanceQuestions.period4 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};
performanceQuestions.period5 = {'Q4', 'Q22', 'Q25', 'Q27', 'Q31', 'Q32', 'Q33'};

plotIdx = 1;
distributionStats = table();

for p = 1:length(periods)
    if ~isfield(allData, sprintf('period%d', p))
        continue;
    end
    
    periodData = allData.(sprintf('period%d', p)).selfData;
    perfQuestions = performanceQuestions.(sprintf('period%d', p));
    
    % 성과 관련 문항 데이터 추출
    for q = 1:min(length(perfQuestions), 4) % 최대 4개 문항만 표시
        if plotIdx > 20
            break;
        end
        
        qName = perfQuestions{q};
        if ismember(qName, periodData.Properties.VariableNames)
            subplot(5, 4, plotIdx);
            
            % 원데이터 추출
            rawData = periodData.(qName);
            if iscell(rawData)
                rawData = cellfun(@str2double, rawData);
            end
            validData = rawData(~isnan(rawData));
            
            % 히스토그램과 박스플롯 동시 표시
            histogram(validData, 'Normalization', 'probability');
            title(sprintf('%s - %s', periods{p}, qName), 'FontSize', 10);
            
            % 통계량 계산
            stats = struct();
            stats.Period = periods{p};
            stats.Question = qName;
            stats.N = length(validData);
            stats.Mean = mean(validData);
            stats.Std = std(validData);
            stats.Min = min(validData);
            stats.Max = max(validData);
            stats.Skewness = skewness(validData);
            stats.Kurtosis = kurtosis(validData);
            
            % 천장/바닥 효과 감지
            uniqueVals = unique(validData);
            stats.UniqueValues = length(uniqueVals);
            stats.CeilingEffect = sum(validData == max(validData))/length(validData)*100; % %
            stats.FloorEffect = sum(validData == min(validData))/length(validData)*100; % %
            
            % 텍스트로 주요 통계 표시
            text(0.6, 0.9, sprintf('평균: %.2f\n왜도: %.2f\n천장: %.1f%%', ...
                stats.Mean, stats.Skewness, stats.CeilingEffect), ...
                'Units', 'normalized', 'FontSize', 8);
            
            % 테이블에 추가
            distributionStats = [distributionStats; struct2table(stats)];
            
            plotIdx = plotIdx + 1;
        end
    end
end

sgtitle('성과 문항 원데이터 분포 (표준화 전)', 'FontSize', 14, 'FontWeight', 'bold');

%% 3. 표준화 방법별 비교
figure('Name', '표준화 방법 비교', 'Position', [150, 150, 1400, 800]);

% 예시 데이터로 period3 사용
if isfield(allData, 'period3')
    periodData = allData.period3.selfData;
    perfQuestions = performanceQuestions.period3;
    
    % 모든 성과 문항 데이터 수집
    allRawScores = [];
    questionLabels = {};
    
    for q = 1:length(perfQuestions)
        qName = perfQuestions{q};
        if ismember(qName, periodData.Properties.VariableNames)
            rawData = periodData.(qName);
            if iscell(rawData)
                rawData = cellfun(@str2double, rawData);
            end
            validData = rawData(~isnan(rawData));
            allRawScores = [allRawScores; validData];
            questionLabels = [questionLabels; repmat({qName}, length(validData), 1)];
        end
    end
    
    % 다양한 표준화 방법 적용
    methods = {'원데이터', 'Min-Max', 'Z-score', '백분율', '순위기반'};
    transformedData = zeros(length(allRawScores), length(methods));
    
    % 1. 원데이터
    transformedData(:, 1) = allRawScores;
    
    % 2. Min-Max 스케일링
    transformedData(:, 2) = (allRawScores - min(allRawScores)) / (max(allRawScores) - min(allRawScores));
    
    % 3. Z-score 표준화
    transformedData(:, 3) = zscore(allRawScores);
    
    % 4. 백분율 변환 (각 문항의 최대값 기준)
    percentData = zeros(size(allRawScores));
    uniqueQuestions = unique(questionLabels);
    for q = 1:length(uniqueQuestions)
        idx = strcmp(questionLabels, uniqueQuestions{q});
        qData = allRawScores(idx);
        percentData(idx) = qData / max(qData) * 100;
    end
    transformedData(:, 4) = percentData;
    
    % 5. 순위 기반
    transformedData(:, 5) = tiedrank(allRawScores) / length(allRawScores);
    
    % 각 방법별 분포 비교
    for m = 1:length(methods)
        subplot(2, 3, m);
        histogram(transformedData(:, m), 30);
        title(methods{m}, 'FontSize', 12, 'FontWeight', 'bold');
        
        % 통계량 표시
        meanVal = mean(transformedData(:, m));
        stdVal = std(transformedData(:, m));
        skewVal = skewness(transformedData(:, m));
        
        text(0.6, 0.9, sprintf('평균: %.2f\n표준편차: %.2f\n왜도: %.2f', ...
            meanVal, stdVal, skewVal), ...
            'Units', 'normalized', 'FontSize', 10, 'BackgroundColor', 'white');
    end
    
    % 상관관계 매트릭스
    subplot(2, 3, 6);
    corrMatrix = corr(transformedData);
    imagesc(corrMatrix);
    colorbar;
    set(gca, 'XTick', 1:5, 'XTickLabel', methods, 'XTickLabelRotation', 45);
    set(gca, 'YTick', 1:5, 'YTickLabel', methods);
    title('표준화 방법 간 상관관계', 'FontSize', 12, 'FontWeight', 'bold');
    
    % 상관계수 값 표시
    for i = 1:5
        for j = 1:5
            text(j, i, sprintf('%.2f', corrMatrix(i,j)), ...
                'HorizontalAlignment', 'center', 'Color', 'w');
        end
    end
end

%% 4. 극단값 영향 분석
figure('Name', '극단값 영향 분석', 'Position', [200, 200, 1200, 700]);

if ~isempty(allRawScores)
    % 극단값 감지
    Q1 = prctile(allRawScores, 25);
    Q3 = prctile(allRawScores, 75);
    IQR = Q3 - Q1;
    outlierIdx = (allRawScores < Q1 - 1.5*IQR) | (allRawScores > Q3 + 1.5*IQR);
    
    subplot(2, 3, 1);
    boxplot(allRawScores);
    title('원데이터 박스플롯', 'FontSize', 12);
    ylabel('점수');
    
    subplot(2, 3, 2);
    scatter(1:length(allRawScores), allRawScores, 10, outlierIdx+1);
    colormap([0 0 1; 1 0 0]);
    title(sprintf('극단값 분포 (%.1f%%)', sum(outlierIdx)/length(outlierIdx)*100));
    xlabel('데이터 인덱스');
    ylabel('원점수');
    legend({'정상', '극단값'}, 'Location', 'best');
    
    % 표준화 전후 성과점수 비교
    subplot(2, 3, 3);
    
    % 원점수 평균
    rawMeans = [];
    standardizedMeans = [];
    
    for q = 1:length(uniqueQuestions)
        idx = strcmp(questionLabels, uniqueQuestions{q});
        rawMeans(q) = mean(allRawScores(idx));
        standardizedMeans(q) = mean(transformedData(idx, 2)); % Min-Max
    end
    
    scatter(rawMeans, standardizedMeans, 100, 'filled');
    xlabel('원점수 평균');
    ylabel('표준화 후 평균');
    title('문항별 평균 변화');
    
    % 각 문항 레이블 표시
    for q = 1:length(uniqueQuestions)
        text(rawMeans(q), standardizedMeans(q), uniqueQuestions{q}, ...
            'FontSize', 8, 'HorizontalAlignment', 'right');
    end
    
    % 평가자 편향 시뮬레이션
    subplot(2, 3, [4 5 6]);
    
    % 가상의 평가자 편향 추가
    biasedData = allRawScores;
    % 첫 50%는 관대한 평가자 (+0.5)
    biasedData(1:floor(length(biasedData)/2)) = biasedData(1:floor(length(biasedData)/2)) + 0.5;
    % 나머지는 엄격한 평가자 (-0.5)
    biasedData(floor(length(biasedData)/2)+1:end) = biasedData(floor(length(biasedData)/2)+1:end) - 0.5;
    
    % 표준화 방법별 편향 처리 비교
    x = 1:length(biasedData);
    plot(x, zscore(allRawScores), 'b-', 'DisplayName', '원데이터 Z-score');
    hold on;
    plot(x, zscore(biasedData), 'r-', 'DisplayName', '편향데이터 Z-score');
    plot(x, tiedrank(biasedData)/length(biasedData), 'g-', 'DisplayName', '편향데이터 순위기반');
    
    xlabel('데이터 인덱스');
    ylabel('표준화된 값');
    title('평가자 편향 시뮬레이션');
    legend('Location', 'best');
    grid on;
end

%% 5. 결과 요약 출력
fprintf('\n========================================\n');
fprintf('데이터 분포 및 표준화 영향 분석 결과\n');
fprintf('========================================\n\n');

% 천장/바닥 효과가 심한 문항 식별
highCeiling = distributionStats(distributionStats.CeilingEffect > 20, :);
highFloor = distributionStats(distributionStats.FloorEffect > 20, :);

if ~isempty(highCeiling)
    fprintf('📊 천장효과가 높은 문항 (>20%%):\n');
    for i = 1:height(highCeiling)
        fprintf('  - %s %s: %.1f%%\n', highCeiling.Period{i}, ...
            highCeiling.Question{i}, highCeiling.CeilingEffect(i));
    end
end

if ~isempty(highFloor)
    fprintf('\n📊 바닥효과가 높은 문항 (>20%%):\n');
    for i = 1:height(highFloor)
        fprintf('  - %s %s: %.1f%%\n', highFloor.Period{i}, ...
            highFloor.Question{i}, highFloor.FloorEffect(i));
    end
end

% 왜도가 높은 문항
highSkew = distributionStats(abs(distributionStats.Skewness) > 1, :);
if ~isempty(highSkew)
    fprintf('\n📊 왜도가 높은 문항 (|왜도| > 1):\n');
    for i = 1:height(highSkew)
        fprintf('  - %s %s: %.2f\n', highSkew.Period{i}, ...
            highSkew.Question{i}, highSkew.Skewness(i));
    end
end

fprintf('\n💡 권장사항:\n');
if any(distributionStats.CeilingEffect > 20) || any(distributionStats.FloorEffect > 20)
    fprintf('  - 극단값 문제 있음: 순위기반 또는 Robust 스케일링 고려\n');
end
if std([distributionStats.Max - distributionStats.Min]) > 2
    fprintf('  - 척도 차이 큼: 백분율 변환 또는 척도별 표준화 필요\n');
end
if mean(abs(distributionStats.Skewness)) > 0.5
    fprintf('  - 분포 왜곡: 로그 변환 또는 Box-Cox 변환 고려\n');
end

% 결과 저장
save('standardization_analysis_results.mat', 'distributionStats', 'transformedData');
fprintf('\n결과가 standardization_analysis_results.mat에 저장되었습니다.\n');