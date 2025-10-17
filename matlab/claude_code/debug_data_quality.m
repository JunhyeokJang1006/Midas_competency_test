%% 데이터 품질 진단 스크립트
% 문제의 근본 원인을 단계별로 찾아보는 스크립트

clear; clc; close all;

% 데이터 경로 설정
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

fprintf('========================================\n');
fprintf('데이터 품질 진단 시작\n');
fprintf('========================================\n\n');

for p = 1:length(periods)
    fprintf('▶ %s 진단 시작\n', periods{p});
    fprintf('----------------------------------------\n');
    
    fileName = fullfile(dataPath, fileNames{p});
    
    try
        % 1단계: 기본 데이터 로드
        fprintf('1. 기본 데이터 로드...\n');
        masterIDs = readtable(fileName, 'Sheet', '기준인원 검토', 'VariableNamingRule', 'preserve');
        selfData = readtable(fileName, 'Sheet', '하향 진단', 'VariableNamingRule', 'preserve');
        questionInfo = readtable(fileName, 'Sheet', '문항 정보_타인진단', 'VariableNamingRule', 'preserve');
        
        fprintf('  ✓ 마스터ID: %d명\n', height(masterIDs));
        fprintf('  ✓ 하향진단: %d명\n', height(selfData));
        fprintf('  ✓ 문항정보: %d개\n', height(questionInfo));
        
        % 2단계: 문항 데이터 추출
        fprintf('\n2. 문항 데이터 추출...\n');
        
        % ID 컬럼 찾기
        idCol = [];
        colNames = selfData.Properties.VariableNames;
        for col = 1:width(selfData)
            colName = lower(colNames{col});
            colData = selfData{:, col};
            if contains(colName, {'id', '사번', 'empno', 'employee'}) && ...
                    ((isnumeric(colData) && ~all(isnan(colData))) || ...
                    (iscell(colData) && ~all(cellfun(@isempty, colData))))
                idCol = col;
                break;
            end
        end
        
        if isempty(idCol)
            fprintf('  ✗ ID 컬럼을 찾을 수 없습니다.\n');
            continue;
        end
        
        % Q로 시작하는 문항 컬럼 추출
        questionCols = {};
        for col = 1:width(selfData)
            colName = colNames{col};
            colData = selfData{:, col};
            if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
                questionCols{end+1} = colName;
            end
        end
        
        fprintf('  ✓ 발견된 문항: %d개\n', length(questionCols));
        
        if length(questionCols) < 3
            fprintf('  ✗ 문항이 너무 적습니다.\n');
            continue;
        end
        
        % 3단계: 응답 데이터 추출 및 기본 통계
        fprintf('\n3. 응답 데이터 분석...\n');
        
        responseData = table2array(selfData(:, questionCols));
        responseIDs = selfData{:, idCol};
        
        fprintf('  - 데이터 크기: %d x %d\n', size(responseData, 1), size(responseData, 2));
        
        % 결측치 분석
        missingCount = sum(isnan(responseData(:)));
        totalCount = numel(responseData);
        missingRate = (missingCount / totalCount) * 100;
        
        fprintf('  - 결측치: %d개 (%.2f%%)\n', missingCount, missingRate);
        
        % 4단계: 데이터 품질 검사
        fprintf('\n4. 데이터 품질 검사...\n');
        
        % 상수 열 검사
        columnVariances = var(responseData, 0, 1, 'omitnan');
        constantColumns = columnVariances < 1e-10;
        constantCount = sum(constantColumns);
        
        fprintf('  - 상수 열: %d개\n', constantCount);
        
        % 저분산 열 검사
        lowVarianceColumns = columnVariances < 0.1;
        lowVarianceCount = sum(lowVarianceColumns);
        
        fprintf('  - 저분산 열: %d개\n', lowVarianceCount);
        
        % 5단계: 상관행렬 분석
        fprintf('\n5. 상관행렬 분석...\n');
        
        try
            % 유효한 데이터만으로 상관행렬 계산
            validData = responseData;
            
            % 상수 열 제거
            if any(constantColumns)
                validData(:, constantColumns) = [];
                fprintf('  - 상수 열 제거 후: %d개 문항\n', size(validData, 2));
            end
            
            if size(validData, 2) < 3
                fprintf('  ✗ 상수 열 제거 후 문항이 너무 적습니다.\n');
                continue;
            end
            
            % 상관행렬 계산
            R = corrcoef(validData, 'Rows', 'pairwise');
            
            % 행렬식과 조건수
            det_R = det(R);
            cond_R = cond(R);
            
            fprintf('  - 상관행렬 크기: %d x %d\n', size(R, 1), size(R, 2));
            fprintf('  - 행렬식: %.2e\n', det_R);
            fprintf('  - 조건수: %.2e\n', cond_R);
            
            % 특이값 분석
            [~, S, ~] = svd(R);
            singular_values = diag(S);
            fprintf('  - 특이값 범위: %.2e ~ %.2e\n', min(singular_values), max(singular_values));
            fprintf('  - 특이값 비율: %.2e\n', max(singular_values)/min(singular_values));
            
            % 품질 판정
            if (det_R > 1e-10) && (cond_R < 1e10)
                quality = 'GOOD';
            elseif (det_R > 1e-15) && (cond_R < 1e12)
                quality = 'CAUTION';
            else
                quality = 'POOR';
            end
            
            fprintf('  - 품질 판정: %s\n', quality);
            
            % 6단계: 상세 진단
            fprintf('\n6. 상세 진단...\n');
            
            % 상관행렬의 특성 분석
            R_values = R(:);
            R_values = R_values(R_values ~= 1); % 대각선 제외
            
            fprintf('  - 상관계수 범위: %.3f ~ %.3f\n', min(R_values), max(R_values));
            fprintf('  - 상관계수 평균: %.3f\n', mean(R_values));
            fprintf('  - 상관계수 표준편차: %.3f\n', std(R_values));
            
            % 높은 상관을 보이는 문항 쌍
            [row, col] = find(abs(R) > 0.9 & R ~= 1);
            highCorrPairs = length(row);
            fprintf('  - 높은 상관 쌍 (|r| > 0.9): %d개\n', highCorrPairs);
            
            % 7단계: 문제 원인 진단
            fprintf('\n7. 문제 원인 진단...\n');
            
            if strcmp(quality, 'POOR')
                fprintf('  ✗ POOR 품질의 원인:\n');
                
                if det_R <= 1e-15
                    fprintf('    → 행렬식이 매우 작음 (%.2e)\n', det_R);
                end
                
                if cond_R >= 1e12
                    fprintf('    → 조건수가 매우 큼 (%.2e)\n', cond_R);
                end
                
                if highCorrPairs > size(validData, 2) * 0.1
                    fprintf('    → 다중공선성 문제 (%d개 높은 상관 쌍)\n', highCorrPairs);
                end
                
                if constantCount > 0
                    fprintf('    → 상수 열 존재 (%d개)\n', constantCount);
                end
                
                if lowVarianceCount > size(validData, 2) * 0.3
                    fprintf('    → 저분산 문항 과다 (%d개)\n', lowVarianceCount);
                end
            else
                fprintf('  ✓ 데이터 품질 양호\n');
            end
            
        catch ME
            fprintf('  ✗ 상관행렬 계산 실패: %s\n', ME.message);
        end
        
    catch ME
        fprintf('  ✗ 데이터 로드 실패: %s\n', ME.message);
    end
    
    fprintf('\n');
end

fprintf('========================================\n');
fprintf('진단 완료\n');
fprintf('========================================\n');
