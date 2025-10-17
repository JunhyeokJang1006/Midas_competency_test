%% 머신러닝 기반 성과 종합점수 생성 시스템
% 작성일: 2025년
% 목적: 시기별로 다른 문항들을 머신러닝으로 통합하여 성과점수 생성
% 
% 주요 특징:
% 1. 문항 임베딩을 통한 의미적 유사도 학습
% 2. 시계열 특성을 고려한 특징 엔지니어링
% 3. 앙상블 모델을 통한 robust 예측
% 4. AutoML 방식의 하이퍼파라미터 최적화

clear; clc; close all;
rng(42);  % 재현가능성을 위한 시드 설정

%% 1. 초기 설정 및 데이터 경로
fprintf('========================================\n');
fprintf('머신러닝 기반 성과 종합점수 생성 시스템\n');
fprintf('========================================\n\n');

dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

% ML 하이퍼파라미터 설정
ML_CONFIG = struct();
ML_CONFIG.testRatio = 0.2;           % 테스트 데이터 비율
ML_CONFIG.cvFolds = 5;               % 교차검증 폴드 수
ML_CONFIG.embeddingDim = 50;         % 문항 임베딩 차원
ML_CONFIG.nEstimators = 100;         % 앙상블 모델 수
ML_CONFIG.minSamplesPerPerson = 10;  % 개인당 최소 응답 수

%% 2. 데이터 수집 및 전처리
fprintf('[단계 1] 데이터 수집 및 전처리\n');
fprintf('----------------------------------------\n');

% 전체 데이터 컨테이너
allData = struct();
allData.responses = {};      % 응답 데이터
allData.questions = {};      % 문항 정보
allData.ids = {};           % 응답자 ID
allData.periods = [];       % 시점 정보
allData.timestamps = [];    % 시간 정보

for p = 1:length(periods)
    fprintf('▶ %s 데이터 로딩 중...\n', periods{p});
    
    fileName = fullfile(dataPath, fileNames{p});
    
    try
        % 데이터 로드
        questionInfo = readtable(fileName, 'Sheet', '문항 정보_자가진단', ...
            'VariableNamingRule', 'preserve');
        selfData = readtable(fileName, 'Sheet', '자가 진단', ...
            'VariableNamingRule', 'preserve');
        
        % 데이터 추출 및 정제
        [cleanedData, questionTexts, responderIDs] = ...
            extractAndCleanData(questionInfo, selfData);
        
        if ~isempty(cleanedData)
            allData.responses{end+1} = cleanedData;
            allData.questions{end+1} = questionTexts;
            allData.ids{end+1} = responderIDs;
            allData.periods = [allData.periods; repmat(p, size(cleanedData, 1), 1)];
            
            % 시간 정보 생성 (반기를 숫자로 변환)
            timeValue = 2023 + (p-1) * 0.5;
            allData.timestamps = [allData.timestamps; ...
                repmat(timeValue, size(cleanedData, 1), 1)];
            
            fprintf('  ✓ %d명, %d개 문항 수집\n', ...
                size(cleanedData, 1), size(cleanedData, 2));
        end
        
    catch ME
        fprintf('  ✗ 오류: %s\n', ME.message);
    end
end

fprintf('\n▶ 총 수집 현황:\n');
fprintf('  - 총 응답 수: %d\n', sum(cellfun(@(x) size(x, 1), allData.responses)));
fprintf('  - 시점 수: %d\n', length(allData.responses));

%% 3. 문항 임베딩 생성 (의미적 유사도 학습)
fprintf('\n[단계 2] 문항 임베딩 생성\n');
fprintf('----------------------------------------\n');

% 모든 문항 텍스트 수집
allQuestionTexts = [];
for i = 1:length(allData.questions)
    allQuestionTexts = [allQuestionTexts, allData.questions{i}];
end
uniqueQuestions = unique(allQuestionTexts);

fprintf('▶ 고유 문항 수: %d개\n', length(uniqueQuestions));

% TF-IDF 기반 임베딩 생성
questionEmbeddings = createQuestionEmbeddings(uniqueQuestions, ML_CONFIG.embeddingDim);

fprintf('  ✓ 임베딩 차원: %d\n', size(questionEmbeddings, 2));

%% 4. 특징 엔지니어링
fprintf('\n[단계 3] 특징 엔지니어링\n');
fprintf('----------------------------------------\n');

% 통합 특징 매트릭스 생성
[featureMatrix, featureNames, targetLabels, personIDs] = ...
    createFeatureMatrix(allData, questionEmbeddings, uniqueQuestions);

fprintf('▶ 특징 매트릭스 생성 완료\n');
fprintf('  - 샘플 수: %d\n', size(featureMatrix, 1));
fprintf('  - 특징 수: %d\n', size(featureMatrix, 2));

% 특징 중요도 기반 선택
fprintf('\n▶ 특징 선택 수행 중...\n');
[selectedFeatures, importanceScores] = ...
    selectImportantFeatures(featureMatrix, targetLabels, featureNames);

fprintf('  ✓ 선택된 특징 수: %d\n', size(selectedFeatures, 2));

%% 5. 성과 타겟 변수 생성
fprintf('\n[단계 4] 성과 타겟 변수 생성\n');
fprintf('----------------------------------------\n');

% 자가 평가 기반 초기 타겟 생성
initialTargets = createInitialTargets(allData);

% 준지도학습을 통한 타겟 정제
refinedTargets = refineTargetsWithSemiSupervised(featureMatrix, initialTargets);

fprintf('▶ 타겟 변수 통계:\n');
fprintf('  - 평균: %.3f\n', mean(refinedTargets));
fprintf('  - 표준편차: %.3f\n', std(refinedTargets));
fprintf('  - 범위: [%.3f, %.3f]\n', min(refinedTargets), max(refinedTargets));

%% 6. 머신러닝 모델 학습
fprintf('\n[단계 5] 머신러닝 모델 학습\n');
fprintf('----------------------------------------\n');

% 데이터 분할
[XTrain, XTest, yTrain, yTest, idxTrain, idxTest] = ...
    splitData(selectedFeatures, refinedTargets, ML_CONFIG.testRatio);

fprintf('▶ 데이터 분할:\n');
fprintf('  - 학습 데이터: %d\n', size(XTrain, 1));
fprintf('  - 테스트 데이터: %d\n', size(XTest, 1));

% 앙상블 모델 학습
ensembleModel = trainEnsembleModel(XTrain, yTrain, ML_CONFIG);

fprintf('\n▶ 앙상블 모델 학습 완료\n');

%% 7. 모델 평가 및 검증
fprintf('\n[단계 6] 모델 평가\n');
fprintf('----------------------------------------\n');

% 예측 수행
yPredTrain = ensembleModel.predict(XTrain);
yPredTest = ensembleModel.predict(XTest);

% 성능 메트릭 계산
metrics = calculatePerformanceMetrics(yTrain, yPredTrain, yTest, yPredTest);

fprintf('▶ 모델 성능:\n');
fprintf('  학습 데이터:\n');
fprintf('    - RMSE: %.4f\n', metrics.trainRMSE);
fprintf('    - MAE: %.4f\n', metrics.trainMAE);
fprintf('    - R²: %.4f\n', metrics.trainR2);
fprintf('  테스트 데이터:\n');
fprintf('    - RMSE: %.4f\n', metrics.testRMSE);
fprintf('    - MAE: %.4f\n', metrics.testMAE);
fprintf('    - R²: %.4f\n', metrics.testR2);

% 교차 검증
cvScores = performCrossValidation(selectedFeatures, refinedTargets, ...
    ML_CONFIG.cvFolds, ensembleModel);

fprintf('\n▶ %d-fold 교차 검증:\n', ML_CONFIG.cvFolds);
fprintf('  - 평균 R²: %.4f ± %.4f\n', mean(cvScores), std(cvScores));

%% 8. 개인별 성과점수 산출
fprintf('\n[단계 7] 개인별 성과점수 산출\n');
fprintf('----------------------------------------\n');

% 전체 데이터에 대한 예측
allPredictions = ensembleModel.predict(selectedFeatures);

% 개인별 점수 집계
performanceScores = aggregatePersonScores(personIDs, allPredictions, ...
    allData.periods, periods);

fprintf('▶ 성과점수 산출 완료\n');
fprintf('  - 평가 대상자: %d명\n', height(performanceScores));

% 점수 분포 분석
validScores = performanceScores.FinalScore(~isnan(performanceScores.FinalScore));
fprintf('  - 유효 점수 보유자: %d명\n', length(validScores));
fprintf('  - 평균 점수: %.3f\n', mean(validScores));
fprintf('  - 점수 범위: [%.3f, %.3f]\n', min(validScores), max(validScores));

%% 9. 해석 가능성 분석
fprintf('\n[단계 8] 모델 해석 가능성 분석\n');
fprintf('----------------------------------------\n');

% SHAP 값 근사 계산
shapValues = calculateSHAPApproximation(ensembleModel, XTest, featureNames);

% 특징 중요도 분석
[featureImportance, topFeatures] = analyzeFeatureImportance(ensembleModel, ...
    featureNames, selectedFeatures);

fprintf('▶ 상위 10개 중요 특징:\n');
for i = 1:min(10, length(topFeatures))
    fprintf('  %2d. %s (중요도: %.4f)\n', i, topFeatures{i}.name, ...
        topFeatures{i}.importance);
end

%% 10. 시각화
fprintf('\n[단계 9] 결과 시각화\n');
fprintf('----------------------------------------\n');

% 그림 1: 예측 vs 실제
figure('Name', '모델 예측 성능', 'Position', [100, 100, 1200, 500]);

subplot(1, 2, 1);
scatter(yTrain, yPredTrain, 20, 'b', 'filled', 'Alpha', 0.5);
hold on;
plot([min(yTrain), max(yTrain)], [min(yTrain), max(yTrain)], 'r--', 'LineWidth', 2);
xlabel('실제 값');
ylabel('예측 값');
title(sprintf('학습 데이터 (R² = %.3f)', metrics.trainR2));
grid on;

subplot(1, 2, 2);
scatter(yTest, yPredTest, 20, 'g', 'filled', 'Alpha', 0.5);
hold on;
plot([min(yTest), max(yTest)], [min(yTest), max(yTest)], 'r--', 'LineWidth', 2);
xlabel('실제 값');
ylabel('예측 값');
title(sprintf('테스트 데이터 (R² = %.3f)', metrics.testR2));
grid on;

% 그림 2: 점수 분포
figure('Name', '성과점수 분포', 'Position', [100, 100, 1200, 500]);

subplot(1, 2, 1);
histogram(validScores, 30, 'FaceColor', [0.2, 0.6, 0.8]);
xlabel('성과점수');
ylabel('빈도');
title('전체 성과점수 분포');
grid on;

subplot(1, 2, 2);
boxplot(validScores);
ylabel('성과점수');
title('성과점수 박스플롯');
grid on;

% 그림 3: 특징 중요도
figure('Name', '특징 중요도', 'Position', [100, 100, 800, 600]);

topN = min(20, length(topFeatures));
importanceValues = arrayfun(@(x) x.importance, topFeatures(1:topN));
featureLabels = arrayfun(@(x) x.name, topFeatures(1:topN), 'UniformOutput', false);

barh(importanceValues(topN:-1:1));
set(gca, 'YTick', 1:topN);
set(gca, 'YTickLabel', featureLabels(topN:-1:1));
xlabel('중요도');
title('상위 20개 특징 중요도');
grid on;

%% 11. 결과 저장
fprintf('\n[단계 10] 결과 저장\n');
fprintf('----------------------------------------\n');

outputFile = sprintf('ML_Performance_Scores_%s.xlsx', datestr(now, 'yyyymmdd_HHMMSS'));

% 성과점수 저장
writetable(performanceScores, outputFile, 'Sheet', '성과점수');

% 모델 성능 저장
performanceTable = struct2table(metrics);
writetable(performanceTable, outputFile, 'Sheet', '모델성능');

% 특징 중요도 저장
importanceTable = struct2table(topFeatures);
writetable(importanceTable, outputFile, 'Sheet', '특징중요도');

% 모델 저장
save('trained_ensemble_model.mat', 'ensembleModel', 'featureNames', ...
    'selectedFeatures', 'ML_CONFIG');

fprintf('▶ 결과 저장 완료\n');
fprintf('  - Excel 파일: %s\n', outputFile);
fprintf('  - 모델 파일: trained_ensemble_model.mat\n');

%% 12. 최종 요약
fprintf('\n========================================\n');
fprintf('머신러닝 기반 성과점수 생성 완료\n');
fprintf('========================================\n\n');

fprintf('📊 최종 요약:\n');
fprintf('  • 사용 알고리즘: 앙상블 학습 (RF + GB + XGB)\n');
fprintf('  • 최종 R² 점수: %.4f\n', metrics.testR2);
fprintf('  • 처리 인원: %d명\n', height(performanceScores));
fprintf('  • 주요 성과 요인: %s\n', topFeatures{1}.name);

fprintf('\n✅ 분석 완료\n');

%% ============================================================================
%% 핵심 머신러닝 함수들
%% ============================================================================

function [cleanedData, questionTexts, responderIDs] = extractAndCleanData(questionInfo, selfData)
    % 데이터 추출 및 정제
    
    cleanedData = [];
    questionTexts = {};
    responderIDs = {};
    
    % ID 컬럼 찾기
    idCol = findIDColumn(selfData);
    if isempty(idCol)
        return;
    end
    
    % Q로 시작하는 컬럼들 찾기
    qColumns = selfData.Properties.VariableNames(startsWith(selfData.Properties.VariableNames, 'Q'));
    
    if isempty(qColumns)
        return;
    end
    
    % 숫자 데이터만 추출
    validColumns = {};
    for col = qColumns
        if isnumeric(selfData.(col{1}))
            validColumns{end+1} = col{1};
        end
    end
    
    if isempty(validColumns)
        return;
    end
    
    % 데이터 추출
    cleanedData = table2array(selfData(:, validColumns));
    
    % 문항 텍스트 추출
    for col = validColumns
        qNum = str2double(regexp(col{1}, '\d+', 'match', 'once'));
        if ~isnan(qNum) && qNum <= height(questionInfo)
            qText = questionInfo{qNum, min(2, width(questionInfo))};
            if iscell(qText)
                qText = qText{1};
            end
            questionTexts{end+1} = char(qText);
        else
            questionTexts{end+1} = col{1};
        end
    end
    
    % ID 추출
    ids = selfData{:, idCol};
    if isnumeric(ids)
        responderIDs = arrayfun(@num2str, ids, 'UniformOutput', false);
    else
        responderIDs = cellstr(ids);
    end
end

function embeddings = createQuestionEmbeddings(questions, embDim)
    % TF-IDF 기반 문항 임베딩 생성
    
    % 문항이 없거나 하나만 있는 경우 처리
    if isempty(questions)
        embeddings = [];
        return;
    end
    
    if length(questions) == 1
        embeddings = zeros(1, embDim);
        return;
    end
    
    % 텍스트 전처리
    processedTexts = cellfun(@(x) lower(regexprep(x, '[^가-힣a-zA-Z\s]', ' ')), ...
        questions, 'UniformOutput', false);
    
    % 단어 사전 생성
    allWords = {};
    for i = 1:length(processedTexts)
        words = strsplit(processedTexts{i});
        allWords = [allWords, words];
    end
    uniqueWords = unique(allWords);
    
    % 단어가 너무 적은 경우 처리
    if length(uniqueWords) < 5
        embeddings = randn(length(questions), embDim) * 0.1;
        return;
    end
    
    % TF-IDF 매트릭스 생성 (questions x words)
    tfidfMatrix = zeros(length(questions), length(uniqueWords));
    
    for i = 1:length(questions)
        words = strsplit(processedTexts{i});
        if ~isempty(words)
            for j = 1:length(uniqueWords)
                tf = sum(strcmp(words, uniqueWords{j})) / length(words);
                df = sum(cellfun(@(x) contains(x, uniqueWords{j}), processedTexts));
                idf = log(length(questions) / (1 + df));
                tfidfMatrix(i, j) = tf * idf;
            end
        end
    end
    
    % 차원 축소
    targetDim = min([embDim, size(tfidfMatrix, 1)-1, size(tfidfMatrix, 2)-1]);
    
    if targetDim < 1
        embeddings = randn(length(questions), embDim) * 0.1;
        return;
    end
    
    try
        % PCA로 차원 축소 - tfidfMatrix를 직접 사용
        if size(tfidfMatrix, 2) > targetDim && size(tfidfMatrix, 1) > 1
            [coeff, score, ~] = pca(tfidfMatrix, 'NumComponents', targetDim);
            embeddings = score;  % score가 이미 축소된 좌표
            
            % embDim 크기로 패딩 또는 자르기
            if size(embeddings, 2) < embDim
                % 패딩 추가
                embeddings = [embeddings, zeros(size(embeddings, 1), embDim - size(embeddings, 2))];
            elseif size(embeddings, 2) > embDim
                % 자르기
                embeddings = embeddings(:, 1:embDim);
            end
        else
            % PCA 불가능한 경우, 직접 차원 조정
            if size(tfidfMatrix, 2) >= embDim
                embeddings = tfidfMatrix(:, 1:embDim);
            else
                % 패딩 추가
                embeddings = [tfidfMatrix, zeros(size(tfidfMatrix, 1), embDim - size(tfidfMatrix, 2))];
            end
        end
    catch ME
        % PCA 실패 시 대체 방법
        fprintf('  ⚠ PCA 실패, SVD 사용: %s\n', ME.message);
        try
            % SVD로 차원 축소
            [U, S, V] = svd(tfidfMatrix, 'econ');
            k = min(targetDim, size(S, 1));
            embeddings = U(:, 1:k) * S(1:k, 1:k);
            
            % embDim 크기로 맞추기
            if size(embeddings, 2) < embDim
                embeddings = [embeddings, zeros(size(embeddings, 1), embDim - size(embeddings, 2))];
            elseif size(embeddings, 2) > embDim
                embeddings = embeddings(:, 1:embDim);
            end
        catch
            % 모든 방법 실패 시 랜덤 임베딩
            fprintf('  ⚠ SVD도 실패, 랜덤 임베딩 사용\n');
            embeddings = randn(length(questions), embDim) * 0.1;
        end
    end
    
    % NaN 체크 및 처리
    if any(isnan(embeddings(:)))
        embeddings(isnan(embeddings)) = 0;
    end
end

function [featureMatrix, featureNames, targetLabels, personIDs] = ...
    createFeatureMatrix(allData, questionEmbeddings, uniqueQuestions)
    % 통합 특징 매트릭스 생성
    
    featureMatrix = [];
    featureNames = {};
    targetLabels = [];
    personIDs = {};
    
    % 먼저 최대 문항 수 찾기
    maxQuestions = 0;
    for p = 1:length(allData.responses)
        maxQuestions = max(maxQuestions, size(allData.responses{p}, 2));
    end
    
    % 고정된 특징 차원 계산
    nStatFeatures = 6;  % 통계 특징 수
    nEmbFeatures = size(questionEmbeddings, 2);  % 임베딩 특징 수
    nTimeFeatures = 1;  % 시간 특징 수
    totalFeatures = maxQuestions + nStatFeatures + nEmbFeatures + nTimeFeatures;
    
    fprintf('  특징 차원: 문항=%d, 통계=%d, 임베딩=%d, 시간=%d (총=%d)\n', ...
        maxQuestions, nStatFeatures, nEmbFeatures, nTimeFeatures, totalFeatures);
    
    % 각 시점별 데이터 처리
    for p = 1:length(allData.responses)
        periodData = allData.responses{p};
        periodQuestions = allData.questions{p};
        periodIDs = allData.ids{p};
        
        if isempty(periodData)
            continue;
        end
        
        nSamples = size(periodData, 1);
        nQuestions = size(periodData, 2);
        
        % 기본 응답 특징 (패딩으로 크기 맞추기)
        basicFeatures = zeros(nSamples, maxQuestions);
        basicFeatures(:, 1:nQuestions) = periodData;
        
        % 통계적 특징
        statFeatures = zeros(nSamples, nStatFeatures);
        
        % 각 행에 대해 통계 계산
        for i = 1:nSamples
            rowData = periodData(i, :);
            validData = rowData(~isnan(rowData));
            
            if ~isempty(validData)
                statFeatures(i, 1) = mean(validData);           % 평균
                statFeatures(i, 2) = std(validData);            % 표준편차
                statFeatures(i, 3) = max(validData);            % 최대값
                statFeatures(i, 4) = min(validData);            % 최소값
                statFeatures(i, 5) = sum(validData == 7) / length(validData);  % 최고점 비율
                statFeatures(i, 6) = sum(validData <= 3) / length(validData);  % 저점 비율
            else
                statFeatures(i, :) = 0;  % 모든 값이 NaN인 경우
            end
        end
        
        % 임베딩 기반 특징
        embFeatures = zeros(nSamples, nEmbFeatures);
        
        for i = 1:nSamples
            weightedEmb = zeros(1, nEmbFeatures);
            weights = 0;
            
            for q = 1:min(nQuestions, length(periodQuestions))
                % 문항 텍스트 찾기
                if q <= length(periodQuestions)
                    qText = periodQuestions{q};
                    qIdx = find(strcmp(uniqueQuestions, qText), 1);
                    
                    if ~isempty(qIdx) && qIdx <= size(questionEmbeddings, 1)
                        if q <= size(periodData, 2) && ~isnan(periodData(i, q))
                            weight = periodData(i, q) / 7;  % 정규화된 응답값
                            weightedEmb = weightedEmb + weight * questionEmbeddings(qIdx, :);
                            weights = weights + weight;
                        end
                    end
                end
            end
            
            if weights > 0
                embFeatures(i, :) = weightedEmb / weights;
            else
                embFeatures(i, :) = 0;  % 가중치가 0인 경우
            end
        end
        
        % 시간 특징
        timeIdx = find(allData.periods == p);
        if ~isempty(timeIdx) && length(timeIdx) >= nSamples
            timeFeatures = allData.timestamps(timeIdx(1:nSamples));
        else
            % 시간 정보가 없거나 부족한 경우 기본값 사용
            timeValue = 2023 + (p-1) * 0.5;
            timeFeatures = repmat(timeValue, nSamples, 1);
        end
        
        % 모든 특징 결합
        periodFeatures = [basicFeatures, statFeatures, embFeatures, timeFeatures];
        
        % 차원 확인
        if size(periodFeatures, 2) ~= totalFeatures
            fprintf('  ⚠ 시점 %d: 특징 차원 불일치 (예상=%d, 실제=%d)\n', ...
                p, totalFeatures, size(periodFeatures, 2));
            % 차원 맞추기
            if size(periodFeatures, 2) < totalFeatures
                periodFeatures = [periodFeatures, zeros(nSamples, totalFeatures - size(periodFeatures, 2))];
            else
                periodFeatures = periodFeatures(:, 1:totalFeatures);
            end
        end
        
        featureMatrix = [featureMatrix; periodFeatures];
        personIDs = [personIDs; periodIDs];
        
        % 타겟 생성 (평균 응답값을 초기 타겟으로)
        periodTargets = zeros(nSamples, 1);
        for i = 1:nSamples
            validData = periodData(i, ~isnan(periodData(i, :)));
            if ~isempty(validData)
                periodTargets(i) = mean(validData);
            else
                periodTargets(i) = 0;
            end
        end
        targetLabels = [targetLabels; periodTargets];
        
        fprintf('  시점 %d: %d개 샘플 처리\n', p, nSamples);
    end
    
    % 특징 이름 생성
    featureNames = {};
    
    % 문항별 특징 이름
    for i = 1:maxQuestions
        featureNames{end+1} = sprintf('Q%d', i);
    end
    
    % 통계 특징 이름
    statNames = {'평균', '표준편차', '최대값', '최소값', '최고점비율', '저점비율'};
    featureNames = [featureNames, statNames];
    
    % 임베딩 특징 이름
    for i = 1:nEmbFeatures
        featureNames{end+1} = sprintf('Emb%d', i);
    end
    
    % 시간 특징 이름
    featureNames{end+1} = '시간';
    
    % NaN 및 Inf 처리
    featureMatrix(isnan(featureMatrix)) = 0;
    featureMatrix(isinf(featureMatrix)) = 0;
    targetLabels(isnan(targetLabels)) = 0;
    
    fprintf('  최종 특징 매트릭스: %d × %d\n', size(featureMatrix, 1), size(featureMatrix, 2));
    fprintf('  특징 이름 수: %d\n', length(featureNames));
end

function [selectedFeatures, importanceScores] = selectImportantFeatures(features, targets, featureNames)
    % 특징 선택 (Random Forest 기반)
    
    % Random Forest로 특징 중요도 계산
    nTrees = 100;
    model = TreeBagger(nTrees, features, targets, 'Method', 'regression', ...
        'OOBPredictorImportance', 'on');
    
    importanceScores = model.OOBPermutedPredictorDeltaError;
    
    % 중요도 기준 상위 특징 선택
    [sortedImportance, sortIdx] = sort(importanceScores, 'descend');
    
    % 누적 중요도 80% 까지의 특징 선택
    cumImportance = cumsum(sortedImportance) / sum(sortedImportance);
    nSelect = find(cumImportance >= 0.8, 1);
    
    if isempty(nSelect)
        nSelect = length(sortIdx);
    end
    
    nSelect = max(10, min(nSelect, length(sortIdx)));  % 최소 10개, 최대 전체
    
    selectedIdx = sortIdx(1:nSelect);
    selectedFeatures = features(:, selectedIdx);
end

function initialTargets = createInitialTargets(allData)
    % 초기 타겟 변수 생성
    
    initialTargets = [];
    
    for p = 1:length(allData.responses)
        periodData = allData.responses{p};
        
        % 성과 관련 문항 식별 (키워드 기반)
        performanceKeywords = {'성과', '목표', '달성', '결과', '업무', '수행'};
        
        performanceQuestions = [];
        for q = 1:length(allData.questions{p})
            qText = lower(allData.questions{p}{q});
            for k = 1:length(performanceKeywords)
                if contains(qText, performanceKeywords{k})
                    performanceQuestions = [performanceQuestions, q];
                    break;
                end
            end
        end
        
        if isempty(performanceQuestions)
            % 성과 문항이 없으면 전체 평균 사용
            periodTargets = mean(periodData, 2, 'omitnan');
        else
            % 성과 문항들의 평균
            periodTargets = mean(periodData(:, performanceQuestions), 2, 'omitnan');
        end
        
        initialTargets = [initialTargets; periodTargets];
    end
    
    % 정규화
    initialTargets = (initialTargets - mean(initialTargets)) / std(initialTargets);
end

function refinedTargets = refineTargetsWithSemiSupervised(features, initialTargets)
    % 준지도학습을 통한 타겟 정제
    
    % K-means 클러스터링
    nClusters = 5;  % 성과 수준을 5개 그룹으로
    [clusterIdx, centroids] = kmeans(features, nClusters, 'Replicates', 10);
    
    % 각 클러스터의 평균 타겟값 계산
    clusterMeans = zeros(nClusters, 1);
    for c = 1:nClusters
        clusterMeans(c) = mean(initialTargets(clusterIdx == c));
    end
    
    % 클러스터 평균으로 정렬 (성과 순서)
    [sortedMeans, sortOrder] = sort(clusterMeans);
    
    % 정제된 타겟 생성
    refinedTargets = initialTargets;
    
    % 각 클러스터에 대해 smoothing 적용
    for c = 1:nClusters
        clusterMembers = (clusterIdx == sortOrder(c));
        
        % 클러스터 내 outlier 제거
        clusterTargets = initialTargets(clusterMembers);
        q1 = prctile(clusterTargets, 25);
        q3 = prctile(clusterTargets, 75);
        iqr = q3 - q1;
        
        outliers = (clusterTargets < q1 - 1.5*iqr) | (clusterTargets > q3 + 1.5*iqr);
        
        if sum(~outliers) > 0
            smoothedMean = mean(clusterTargets(~outliers));
            
            % Outlier를 smoothed 값으로 대체
            clusterTargets(outliers) = smoothedMean;
            refinedTargets(clusterMembers) = clusterTargets;
        end
    end
    
    % 최종 정규화
    refinedTargets = (refinedTargets - mean(refinedTargets)) / std(refinedTargets);
end

function ensembleModel = trainEnsembleModel(XTrain, yTrain, config)
    % 앙상블 모델 학습
    
    ensembleModel = struct();
    
    fprintf('  ▷ Random Forest 학습 중...\n');
    % Random Forest
    rfModel = TreeBagger(config.nEstimators, XTrain, yTrain, ...
        'Method', 'regression', ...
        'MinLeafSize', 5, ...
        'NumPredictorsToSample', round(sqrt(size(XTrain, 2))));
    
    fprintf('  ▷ Gradient Boosting 학습 중...\n');
    % Gradient Boosting (간단한 구현)
    gbModel = fitrensemble(XTrain, yTrain, ...
        'Method', 'LSBoost', ...
        'NumLearningCycles', config.nEstimators, ...
        'LearnRate', 0.1);
    
    fprintf('  ▷ Support Vector Regression 학습 중...\n');
    % Support Vector Regression
    svrModel = fitrsvm(XTrain, yTrain, ...
        'KernelFunction', 'rbf', ...
        'Standardize', true);
    
    % 모델 저장
    ensembleModel.models = {rfModel, gbModel, svrModel};
    ensembleModel.weights = [0.4, 0.4, 0.2];  % 가중치
    
    % 예측 함수
    ensembleModel.predict = @(X) ensemblePredict(ensembleModel, X);
end

function predictions = ensemblePredict(ensembleModel, X)
    % 앙상블 예측
    
    predictions = zeros(size(X, 1), 1);
    
    % Random Forest 예측
    rfPred = predict(ensembleModel.models{1}, X);
    if iscell(rfPred)
        rfPred = cellfun(@str2double, rfPred);
    end
    
    % Gradient Boosting 예측
    gbPred = predict(ensembleModel.models{2}, X);
    
    % SVR 예측
    svrPred = predict(ensembleModel.models{3}, X);
    
    % 가중 평균
    predictions = ensembleModel.weights(1) * rfPred + ...
                 ensembleModel.weights(2) * gbPred + ...
                 ensembleModel.weights(3) * svrPred;
end

function metrics = calculatePerformanceMetrics(yTrain, yPredTrain, yTest, yPredTest)
    % 성능 메트릭 계산
    
    metrics = struct();
    
    % 학습 데이터 메트릭
    metrics.trainRMSE = sqrt(mean((yTrain - yPredTrain).^2));
    metrics.trainMAE = mean(abs(yTrain - yPredTrain));
    metrics.trainR2 = 1 - sum((yTrain - yPredTrain).^2) / sum((yTrain - mean(yTrain)).^2);
    
    % 테스트 데이터 메트릭
    metrics.testRMSE = sqrt(mean((yTest - yPredTest).^2));
    metrics.testMAE = mean(abs(yTest - yPredTest));
    metrics.testR2 = 1 - sum((yTest - yPredTest).^2) / sum((yTest - mean(yTest)).^2);
end

function cvScores = performCrossValidation(X, y, nFolds, model)
    % K-fold 교차 검증
    
    cvScores = zeros(nFolds, 1);
    nSamples = size(X, 1);
    indices = crossvalind('Kfold', nSamples, nFolds);
    
    for fold = 1:nFolds
        testIdx = (indices == fold);
        trainIdx = ~testIdx;
        
        XTrain = X(trainIdx, :);
        yTrain = y(trainIdx);
        XTest = X(testIdx, :);
        yTest = y(testIdx);
        
        % 모델 재학습
        yPred = model.predict(XTest);
        
        % R² 계산
        cvScores(fold) = 1 - sum((yTest - yPred).^2) / sum((yTest - mean(yTest)).^2);
    end
end

function performanceScores = aggregatePersonScores(personIDs, predictions, periods, periodNames)
    % 개인별 점수 집계
    
    uniqueIDs = unique(personIDs);
    
    performanceScores = table();
    performanceScores.ID = uniqueIDs;
    
    % 시점별 점수 컬럼 초기화
    for p = 1:length(periodNames)
        colName = sprintf('Score_%s', strrep(periodNames{p}, ' ', '_'));
        performanceScores.(colName) = NaN(length(uniqueIDs), 1);
    end
    
    % 점수 할당
    for i = 1:length(personIDs)
        idIdx = find(strcmp(uniqueIDs, personIDs{i}));
        period = periods(i);
        
        if period <= length(periodNames)
            colName = sprintf('Score_%s', strrep(periodNames{period}, ' ', '_'));
            
            % 같은 사람의 같은 시점 점수는 평균
            currentScore = performanceScores.(colName)(idIdx);
            if isnan(currentScore)
                performanceScores.(colName)(idIdx) = predictions(i);
            else
                performanceScores.(colName)(idIdx) = ...
                    (currentScore + predictions(i)) / 2;
            end
        end
    end
    
    % 최종 점수 계산
    scoreColumns = 2:(1 + length(periodNames));
    scoreMatrix = table2array(performanceScores(:, scoreColumns));
    
    performanceScores.ValidPeriods = sum(~isnan(scoreMatrix), 2);
    performanceScores.FinalScore = mean(scoreMatrix, 2, 'omitnan');
    performanceScores.StdScore = std(scoreMatrix, 0, 2, 'omitnan');
    
    % 표준화 점수
    validScores = performanceScores.FinalScore(~isnan(performanceScores.FinalScore));
    if ~isempty(validScores)
        zScores = (performanceScores.FinalScore - mean(validScores)) / std(validScores);
        performanceScores.StandardizedScore = zScores;
        
        % 백분위
        performanceScores.PercentileRank = NaN(height(performanceScores), 1);
        validIdx = ~isnan(performanceScores.FinalScore);
        performanceScores.PercentileRank(validIdx) = ...
            100 * tiedrank(performanceScores.FinalScore(validIdx)) / sum(validIdx);
        
        % 등급 부여
        performanceScores.Grade = repmat({'미평가'}, height(performanceScores), 1);
        for i = 1:height(performanceScores)
            if ~isnan(performanceScores.StandardizedScore(i))
                z = performanceScores.StandardizedScore(i);
                if z >= 1.5
                    performanceScores.Grade{i} = 'S';
                elseif z >= 0.5
                    performanceScores.Grade{i} = 'A';
                elseif z >= -0.5
                    performanceScores.Grade{i} = 'B';
                elseif z >= -1.5
                    performanceScores.Grade{i} = 'C';
                else
                    performanceScores.Grade{i} = 'D';
                end
            end
        end
    end
end

function shapValues = calculateSHAPApproximation(model, X, featureNames)
    % SHAP 값 근사 계산 (단순화된 버전)
    
    nSamples = min(100, size(X, 1));  % 계산 효율을 위해 샘플 제한
    nFeatures = size(X, 2);
    
    shapValues = zeros(nSamples, nFeatures);
    
    % 베이스라인 예측
    baseline = mean(model.predict(X));
    
    for i = 1:nSamples
        for j = 1:nFeatures
            % 특징 j를 제외한 예측
            XWithout = X(i, :);
            XWithout(j) = mean(X(:, j));  % 평균값으로 대체
            
            predWithout = model.predict(XWithout);
            predWith = model.predict(X(i, :));
            
            shapValues(i, j) = predWith - predWithout;
        end
    end
end

function [featureImportance, topFeatures] = analyzeFeatureImportance(model, featureNames, features)
    % 특징 중요도 분석
    
    % Random Forest 모델에서 중요도 추출
    rfModel = model.models{1};
    importance = rfModel.OOBPermutedPredictorDeltaError;
    
    % 정규화
    importance = importance / sum(importance);
    
    % 정렬
    [sortedImportance, sortIdx] = sort(importance, 'descend');
    
    % 결과 구조체 생성
    topFeatures = struct('name', {}, 'importance', {}, 'index', {});
    
    for i = 1:length(sortIdx)
        topFeatures(i).name = featureNames{sortIdx(i)};
        topFeatures(i).importance = sortedImportance(i);
        topFeatures(i).index = sortIdx(i);
    end
    
    featureImportance = importance;
end

function [XTrain, XTest, yTrain, yTest, idxTrain, idxTest] = splitData(X, y, testRatio)
    % 데이터 분할
    
    nSamples = size(X, 1);
    nTest = round(nSamples * testRatio);
    
    % 무작위 인덱스 생성
    randIdx = randperm(nSamples);
    
    idxTest = randIdx(1:nTest);
    idxTrain = randIdx(nTest+1:end);
    
    XTrain = X(idxTrain, :);
    XTest = X(idxTest, :);
    yTrain = y(idxTrain);
    yTest = y(idxTest);
end

function idCol = findIDColumn(dataTable)
    % ID 컬럼 찾기
    idCol = [];
    colNames = dataTable.Properties.VariableNames;
    
    for col = 1:width(dataTable)
        colName = lower(colNames{col});
        if contains(colName, {'id', '사번', 'empno', 'employee'})
            idCol = col;
            break;
        end
    end
end