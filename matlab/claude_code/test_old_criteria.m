%% 기존 기준으로 테스트해보기
% 기존에는 아마도 더 관대한 기준을 사용했을 것

clear; clc; close all;

% 데이터 경로 설정
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods = {'23년_하반기', '24년_상반기', '24년_하반기', '25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

fprintf('========================================\n');
fprintf('기존 기준으로 테스트\n');
fprintf('========================================\n\n');

for p = 1:length(periods)
    fprintf('▶ %s 테스트\n', periods{p});
    fprintf('----------------------------------------\n');
    
    fileName = fullfile(dataPath, fileNames{p});
    
    try
        % 데이터 로드
        selfData = readtable(fileName, 'Sheet', '하향 진단', 'VariableNamingRule', 'preserve');
        
        % 문항 데이터 추출
        colNames = selfData.Properties.VariableNames;
        questionCols = {};
        for col = 1:width(selfData)
            colName = colNames{col};
            colData = selfData{:, col};
            if isnumeric(colData) && (startsWith(colName, 'Q') || startsWith(colName, 'q'))
                questionCols{end+1} = colName;
            end
        end
        
        responseData = table2array(selfData(:, questionCols));
        
        % 상관행렬 계산
        R = corrcoef(responseData, 'Rows', 'pairwise');
        det_R = det(R);
        cond_R = cond(R);
        
        fprintf('  - 행렬식: %.2e\n', det_R);
        fprintf('  - 조건수: %.2e\n', cond_R);
        
        % 현재 기준 (너무 엄격)
        fprintf('\n  [현재 기준 - 너무 엄격]\n');
        if (det_R > 1e-10) && (cond_R < 1e10)
            quality_current = 'GOOD';
        elseif (det_R > 1e-15) && (cond_R < 1e12)
            quality_current = 'CAUTION';
        else
            quality_current = 'POOR';
        end
        fprintf('    → %s\n', quality_current);
        
        % 제안 1: 더 관대한 기준
        fprintf('\n  [제안 1: 더 관대한 기준]\n');
        if (det_R > 1e-20) && (cond_R < 1e15)
            quality_relaxed1 = 'GOOD';
        elseif (det_R > 1e-25) && (cond_R < 1e18)
            quality_relaxed1 = 'CAUTION';
        else
            quality_relaxed1 = 'POOR';
        end
        fprintf('    → %s\n', quality_relaxed1);
        
        % 제안 2: 매우 관대한 기준
        fprintf('\n  [제안 2: 매우 관대한 기준]\n');
        if (det_R > 1e-30) && (cond_R < 1e20)
            quality_relaxed2 = 'GOOD';
        elseif (det_R > 1e-40) && (cond_R < 1e25)
            quality_relaxed2 = 'CAUTION';
        else
            quality_relaxed2 = 'POOR';
        end
        fprintf('    → %s\n', quality_relaxed2);
        
        % 제안 3: 조건수만 고려 (행렬식 무시)
        fprintf('\n  [제안 3: 조건수만 고려]\n');
        if cond_R < 1e10
            quality_cond_only = 'GOOD';
        elseif cond_R < 1e15
            quality_cond_only = 'CAUTION';
        else
            quality_cond_only = 'POOR';
        end
        fprintf('    → %s\n', quality_cond_only);
        
        % 제안 4: 완전히 관대한 기준
        fprintf('\n  [제안 4: 완전히 관대한 기준]\n');
        if cond_R < 1e20
            quality_very_relaxed = 'GOOD';
        else
            quality_very_relaxed = 'CAUTION';
        end
        fprintf('    → %s\n', quality_very_relaxed);
        
    catch ME
        fprintf('  ✗ 오류: %s\n', ME.message);
    end
    
    fprintf('\n');
end

fprintf('========================================\n');
fprintf('결론: 현재 기준이 너무 엄격합니다!\n');
fprintf('기존에는 더 관대한 기준을 사용했을 것입니다.\n');
fprintf('========================================\n');
