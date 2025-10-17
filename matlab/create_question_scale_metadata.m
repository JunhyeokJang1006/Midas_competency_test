files = {
    struct("period","24년 상반기", "path","D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx"),
    struct("period","24년 하반기", "path","D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_하반기_역량진단_응답데이터.xlsx"),
    struct("period","25년 상반기", "path","D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/25년_상반기_역량진단_응답데이터.xlsx")
};

rows = [];
for i = 1:numel(files)
    info = files{i};
    if ~isfile(info.path)
        warning("파일을 찾을 수 없습니다: %s", info.path);
        continue;
    end

    sheetName = selectQuestionSheet(info.path);
    if sheetName == ""
        warning("문항 응답 시트를 찾을 수 없습니다: %s", info.path);
        continue;
    end

    try
        selfData = readtable(info.path, "Sheet", sheetName, "VariableNamingRule", "preserve");
    catch ME
        warning("응답 시트 로드 실패 (%s): %s", info.path, ME.message);
        continue;
    end

    questionInfo = [];
    try
        questionInfo = readtable(info.path, "Sheet", "문항 정보_타인진단", "VariableNamingRule", "preserve");
    catch
    end

    questionMap = buildQuestionInfoMap(questionInfo);

    varNames = selfData.Properties.VariableNames;
    for v = 1:numel(varNames)
        name = varNames{v};
        column = selfData{:, v};
        if ~isnumeric(column)
            continue;
        end
        if ~(startsWith(name,"Q") || startsWith(name,"q"))
            continue;
        end

        valid = column(~isnan(column));
        if isempty(valid)
            continue;
        end

        uniqueVals = unique(valid);
        uniqueVals = uniqueVals(:)';
        minVal = min(uniqueVals);
        maxVal = max(uniqueVals);
        formattedVals = arrayfun(@(x) formatNumeric(x), uniqueVals, "UniformOutput", false);
        uniqueStr = strjoin(string(formattedVals), ", ");

        [questionText, category] = lookupQuestionInfo(questionMap, name);

        newRow = struct();
        newRow.Period = string(info.period);
        newRow.QuestionID = string(name);
        newRow.Category = string(category);
        newRow.QuestionText = string(questionText);
        newRow.MinObserved = minVal;
        newRow.MaxObserved = maxVal;
        newRow.UniqueCount = numel(uniqueVals);
        newRow.UniqueValues = uniqueStr;
        newRow.SourceFile = string(info.path);
        newRow.ResponseSheet = string(sheetName);

        if isempty(rows)
            rows = newRow;
        else
            rows(end+1) = newRow; %#ok<AGROW>
        end
    end
end

if isempty(rows)
    warning("추출된 문항 정보가 없습니다.");
    return;
end

metadataTable = struct2table(rows);

outputPath = "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata.xlsx";
try
    writetable(metadataTable, outputPath, "Sheet", "QuestionScales");
    fprintf("✓ 문항 척도 메타데이터 저장 완료: %s\n", outputPath);
catch ME
    warning("메타데이터 파일 저장 실패: %s", ME.message);
end

function sheetName = selectQuestionSheet(filePath)
    sheetName = "";
    [~, sheets] = xlsfinfo(filePath);
    if isempty(sheets)
        return;
    end

    bestIdx = 0;
    bestCount = -1;

    for i = 1:numel(sheets)
        candidate = sheets{i};
        try
            tbl = readtable(filePath, "Sheet", candidate, "VariableNamingRule", "preserve", "NumHeaderLines", 0);
            varNames = tbl.Properties.VariableNames;
            qCount = sum(startsWith(varNames, "Q") | startsWith(varNames, "q"));
            if qCount > bestCount
                bestCount = qCount;
                bestIdx = i;
            end
        catch
            continue;
        end
    end

    if bestIdx > 0
        sheetName = sheets{bestIdx};
    end
end

function questionMap = buildQuestionInfoMap(questionInfo)
    questionMap = containers.Map("KeyType", "char", "ValueType", "any");
    if isempty(questionInfo)
        return;
    end

    numRows = height(questionInfo);
    numCols = width(questionInfo);
    for r = 1:numRows
        rawCode = questionInfo{r, 1};
        code = sanitizeCellValue(rawCode);
        if strlength(code) == 0
            continue;
        end

        questionText = "";
        category = "";
        if numCols >= 2
            category = sanitizeCellValue(questionInfo{r, 2});
        end
        if numCols >= 3
            questionText = sanitizeCellValue(questionInfo{r, 3});
        end

        questionMap(char(code)) = struct("text", char(questionText), "category", char(category));
    end
end

function [text, category] = lookupQuestionInfo(questionMap, questionID)
    text = "";
    category = "";
    if isempty(questionMap)
        return;
    end
    if isKey(questionMap, questionID)
        info = questionMap(questionID);
        text = info.text;
        category = info.category;
    else
        lowerKeys = keys(questionMap);
        idx = find(strcmpi(lowerKeys, questionID), 1);
        if ~isempty(idx)
            info = questionMap(lowerKeys{idx});
            text = info.text;
            category = info.category;
        end
    end
end

function value = sanitizeCellValue(cellValue)
    if iscell(cellValue)
        if isempty(cellValue)
            value = "";
            return;
        end
        cellValue = cellValue{1};
    end

    if isnumeric(cellValue)
        if isscalar(cellValue)
            if isnan(cellValue)
                value = "";
            else
                value = string(cellValue);
            end
        else
            value = string(cellValue(1));
        end
    elseif isstring(cellValue)
        value = cellValue;
    elseif ischar(cellValue)
        value = string(cellValue);
    else
        value = string(cellValue);
    end
    value = strtrim(value);
end

function out = formatNumeric(x)
    if abs(x - round(x)) < 1e-9
        out = sprintf("%d", round(x));
    else
        out = sprintf("%.4f", x);
    end
end
