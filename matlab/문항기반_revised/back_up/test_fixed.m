%% 수정된 설정 출력 테스트
clear; clc;

% 설정 구조체 생성
analysisConfig = struct();

% Cook's Distance 이상치 제거 설정
analysisConfig.outlierRemoval.enabled = true;
analysisConfig.outlierRemoval.cooksDThreshold = 0;
analysisConfig.outlierRemoval.method = 'cooks';
analysisConfig.outlierRemoval.maxIterations = 3;
analysisConfig.outlierRemoval.minSampleSize = 30;
analysisConfig.outlierRemoval.reportDetails = true;

% Partial Correlation 분석 설정
analysisConfig.partialCorr.enabled = true;
analysisConfig.partialCorr.controlVariable = 'age';
analysisConfig.partialCorr.method = 'Spearman';
analysisConfig.partialCorr.significanceLevel = 0.05;
analysisConfig.partialCorr.reportDetails = true;
analysisConfig.partialCorr.generatePlots = true;
analysisConfig.partialCorr.saveResults = true;

fprintf('========================================\n');
fprintf('Period별 문항-종합점수 상관 매트릭스 생성\n');
fprintf(' * 엑셀 메탃데이터 기반 100점 척도 변환\n');
fprintf(' * 메타데이터 파일: question_scale_metadata_with_23_rebuilt.xlsx\n');

% Cook's Distance 이상치 제거 설정 출력
if analysisConfig.outlierRemoval.enabled
    fprintf(' * Cook Distance 기반 이상치 제거: 활성화\n');
    if analysisConfig.outlierRemoval.cooksDThreshold == 0
        thresholdText = '4/n (동적)';
    else
        thresholdText = sprintf('%.4f', analysisConfig.outlierRemoval.cooksDThreshold);
    end
    fprintf('   - 임계값: %s, 최대 반복: %d회\n', thresholdText, analysisConfig.outlierRemoval.maxIterations);
else
    fprintf(' * Cook Distance 기반 이상치 제거: 비활성화\n');
end

% Partial Correlation 분석 설정 출력
if analysisConfig.partialCorr.enabled
    fprintf(' * Partial Correlation 분석: 활성화\n');
    fprintf('   - 통제변수: %s, 방법: %s\n', analysisConfig.partialCorr.controlVariable, analysisConfig.partialCorr.method);
    if analysisConfig.partialCorr.generatePlots
        plotStatus = 'ON';
    else
        plotStatus = 'OFF';
    end
    fprintf('   - 유의수준: %.3f, 시각화: %s\n', analysisConfig.partialCorr.significanceLevel, plotStatus);
else
    fprintf(' * Partial Correlation 분석: 비활성화\n');
end
fprintf('========================================\n\n');

% 토글 테스트
disp('=== 토글 테스트 ===')
analysisConfig.partialCorr.generatePlots = false;

if analysisConfig.partialCorr.generatePlots
    plotStatus = 'ON';
else
    plotStatus = 'OFF';
end
fprintf('시각화 OFF로 변경: %s\n', plotStatus);

analysisConfig.partialCorr.generatePlots = true;
if analysisConfig.partialCorr.generatePlots
    plotStatus = 'ON';
else
    plotStatus = 'OFF';
end
fprintf('시각화 ON으로 변경: %s\n', plotStatus);

disp('✓ 삼항 연산자 수정 완료 - 정상 작동!')