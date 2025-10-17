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

    try
        questionInfo = readtable(info.path, "Sheet", "문항 정보_타인진단", "VariableNamingRule", "preserve", "ReadVariableNames", false);
    catch ME
        warning("문항 정보_타인진단 시트 로드 실패 (%s): %s", info.path, ME.message);
        continue;
    end

    data = table2cell(questionInfo);
    numRows = size(data, 1);
    numCols = size(data, 2);

    for r = 1:numRows
        qid = sanitizeCellValue(data{r, 1});
        if strlength(qid) == 0 || ~(startsWith(qid, "Q") || startsWith(qid, "q"))
            continue;
        end

        category = sanitizeCellValue(getCell(data, r, 2));
        questionText = sanitizeCellValue(getCell(data, r, 3));

        optionValues = [];
        optionPairs = strings(0,1);

        for c = 4:numCols
            optionRaw = sanitizeCellValue(getCell(data, r, c));
            if strlength(optionRaw) == 0
                continue;
            end

            [value, label, ok] = parseOption(optionRaw);
            if ok
                optionValues(end+1) = value; %#ok<AGROW>
                optionPairs(end+1,1) = sprintf("%d:%s", value, label);
            else
                optionPairs(end+1,1) = optionRaw;
            end
        end

        minScale = NaN;
        maxScale = NaN;
        if ~isempty(optionValues)
            minScale = min(optionValues);
            maxScale = max(optionValues);
        end

        newRow = struct();
        newRow.Period = string(info.period);
        newRow.QuestionID = string(qid);
        newRow.Category = string(category);
        newRow.QuestionText = string(questionText);
        newRow.MinScale = minScale;
        newRow.MaxScale = maxScale;
        newRow.OptionCount = numel(optionPairs);
        newRow.OptionValues = string(strjoin(arrayfun(@num2str, optionValues, 'UniformOutput', false), ', '));
        newRow.OptionMap = string(strjoin(optionPairs, ' | '));
        newRow.SourceFile = string(info.path);

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

outputPath = "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_from_questioninfo.xlsx";
try
    writetable(metadataTable, outputPath, "Sheet", "QuestionScales");
    fprintf("✓ 문항 척도 메타데이터 저장 완료: %s\n", outputPath);
catch ME
    warning("메타데이터 파일 저장 실패: %s", ME.message);
end

function value = sanitizeCellValue(cellValue)
    if nargin == 0 || isempty(cellValue)
        value = "";
        return;
    end
    if iscell(cellValue)
        if isempty(cellValue)
            value = "";
            return;
        end
        cellValue = cellValue{1};
    end

    if ismissing(cellValue)
        value = "";
    elseif isstring(cellValue)
        value = strtrim(cellValue);
    elseif ischar(cellValue)
        value = string(strtrim(cellValue));
    elseif isnumeric(cellValue)
        if isscalar(cellValue) && ~isnan(cellValue)
            value = string(cellValue);
        else
            value = "";
        end
    else
        value = string(cellValue);
    end
end

function cellValue = getCell(data, r, c)
    if r > size(data,1) || c > size(data,2)
        cellValue = "";
    else
        cellValue = data{r,c};
    end
end

function [value, label, ok] = parseOption(optionStr)
    if strlength(optionStr) == 0
        value = NaN;
        label = "";
        ok = false;
        return;
    end

    optionStr = strip(optionStr);
    pattern = "^\(([-+]?\d+)\)\s*(.*)$";
    tokens = regexp(optionStr, pattern, 'tokens', 'once');
    if isempty(tokens)
        value = NaN;
        label = optionStr;
        ok = false;
        return;
    end

    value = str2double(tokens{1});
    label = string(strtrim(tokens{2}));
    ok = ~isnan(value);
end
