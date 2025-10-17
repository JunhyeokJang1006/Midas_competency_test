function rebuild_23_second_half_metadata()
    inputPath = "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23.xlsx";
    dataFile = "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년_하반기_역량진단_응답데이터.xlsx";
    sourcePeriod = "24년 상반기";
    targetPeriod = "23년 하반기";

    try
        baseTable = readtable(inputPath);
    catch ME
        error("메타데이터 파일 로드 실패: %s", ME.message);
    end

    % 제거: 기존 23년 하반기 행
    baseTable(baseTable.Period == targetPeriod, :) = [];

    % 질문 ID 추출
    questionIDs = collectQuestionIDs(dataFile);
    uniqueIDs = unique(questionIDs, 'stable');

    % 추가할 행 초기화
    newRows = baseTable(1:0, :);

    for i = 1:numel(uniqueIDs)
        qid = uniqueIDs(i);
        matchIdx = find(baseTable.QuestionID == qid & baseTable.Period == sourcePeriod, 1);
        if isempty(matchIdx)
            matchIdx = find(baseTable.QuestionID == qid, 1);
        end

        if ~isempty(matchIdx)
            row = baseTable(matchIdx, :);
            row.Period = string(targetPeriod);
            row.SourceFile = string(dataFile);
            newRows = [newRows; row]; %#ok<AGROW>
        else
            blankRow = table(string(targetPeriod), qid, string(""), string(""), NaN, NaN, NaN, string(""), string(""), string(dataFile), ...
                'VariableNames', {'Period','QuestionID','Category','QuestionText','MinScale','MaxScale','OptionCount','OptionValues','OptionMap','SourceFile'});
            newRows = [newRows; blankRow]; %#ok<AGROW>
        end
    end

    % 병합 및 정렬
    mergedTable = [baseTable; newRows];
    if ismember('Period', mergedTable.Properties.VariableNames) && ismember('QuestionID', mergedTable.Properties.VariableNames)
        mergedTable = sortrows(mergedTable, {'Period','QuestionID'});
    end

    writetable(mergedTable, inputPath, 'Sheet', 'QuestionScales');
    fprintf('✓ 23년 하반기 문항 %d개 갱신 완료 (동일 QuestionID는 24년 상반기 데이터 기반).\n', numel(uniqueIDs));
end

function questionIDs = collectQuestionIDs(filePath)
    [~, sheets] = xlsfinfo(filePath);
    if isempty(sheets)
        error('시트를 찾을 수 없습니다: %s', filePath);
    end
    bestSheet = "";
    bestCount = -1;
    for i = 1:numel(sheets)
        try
            tbl = readtable(filePath, 'Sheet', sheets{i}, 'VariableNamingRule', 'preserve');
        catch
            continue;
        end
        varNames = tbl.Properties.VariableNames;
        mask = startsWith(varNames, 'Q') | startsWith(varNames, 'q');
        count = sum(mask);
        if count > bestCount
            bestCount = count;
            bestSheet = sheets{i};
        end
    end
    if strlength(bestSheet) == 0
        error('문항 시트를 찾을 수 없습니다: %s', filePath);
    end
    tbl = readtable(filePath, 'Sheet', bestSheet, 'VariableNamingRule', 'preserve');
    varNames = tbl.Properties.VariableNames;
    mask = startsWith(varNames, 'Q') | startsWith(varNames, 'q');
    questionIDs = string(varNames(mask));
end

rebuild_23_second_half_metadata();
