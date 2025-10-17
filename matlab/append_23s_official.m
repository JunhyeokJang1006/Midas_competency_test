officialFile = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
existingPath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_from_questioninfo.xlsx';
outPath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23.xlsx';

if isfile(existingPath)
    existingTable = readtable(existingPath);
else
    existingTable = table();
end

officialTable = readtable(officialFile, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');

rows = table();

for r = 1:height(officialTable)
    qid = sanitizeCellValue(officialTable{r, 3});
    if strlength(qid) == 0 || ~(startsWith(qid, "Q") || startsWith(qid, "q"))
        continue;
    end

    category = sanitizeCellValue(officialTable{r, 4});
    questionText = sanitizeCellValue(officialTable{r, 5});
    optionRaw = officialTable{r, 6};
    if iscell(optionRaw)
        if isempty(optionRaw)
            continue;
        end
        optionRaw = optionRaw{1};
    end
    if ismissing(optionRaw)
        continue;
    end
    optionChar = char(optionRaw);

    tokens = regexp(optionChar, '\((\d+)\)\s*([^()]+)', 'tokens');
    if isempty(tokens)
        continue;
    end

    values = zeros(numel(tokens), 1);
    optionPairs = strings(numel(tokens), 1);
    for t = 1:numel(tokens)
        val = str2double(tokens{t}{1});
        label = strtrim(tokens{t}{2});
        values(t) = val;
        optionPairs(t) = sprintf('%d:%s', val, label);
    end

    minScale = min(values);
    maxScale = max(values);
    optionValuesStr = strjoin(arrayfun(@num2str, values', 'UniformOutput', false), ', ');
    optionMapStr = strjoin(optionPairs, ' | ');

    newRow = table("23년 상반기", string(qid), string(category), string(questionText), minScale, maxScale, numel(values), string(optionValuesStr), string(optionMapStr), string(officialFile), ...
        'VariableNames', {'Period','QuestionID','Category','QuestionText','MinScale','MaxScale','OptionCount','OptionValues','OptionMap','SourceFile'});

    rows = [rows; newRow]; %#ok<AGROW>
end

if isempty(rows)
    error('추출된 문항이 없습니다.');
end

if isempty(existingTable)
    mergedTable = rows;
else
    mergedTable = [existingTable; rows];
end

writetable(mergedTable, outPath, 'Sheet', 'QuestionScales');
fprintf('✓ 23년 상반기 문항 척도 포함 메타데이터 저장 완료: %s\n', outPath);

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
