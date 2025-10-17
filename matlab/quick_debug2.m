officialFile = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
existingPath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_from_questioninfo.xlsx';

officialTable = readtable(officialFile, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');

for r = 1:5
    try
        qid = sanitizeCellValue(officialTable{r,3});
        optionRaw = officialTable{r,6};
        if iscell(optionRaw)
            optionRaw = optionRaw{1};
        end
        optionChar = char(optionRaw);
        tokens = regexp(optionChar, '\\((\\d+)\\)\\s*([^()]+)', 'tokens');
        fprintf('Row %d tokens %d\n', r, numel(tokens));
        values = zeros(numel(tokens),1);
        optionPairs = strings(numel(tokens),1);
        for t=1:numel(tokens)
            values(t) = str2double(tokens{t}{1});
            optionPairs(t) = sprintf('%d:%s', str2double(tokens{t}{1}), strtrim(tokens{t}{2}));
        end
        optionValuesStr = strjoin(arrayfun(@num2str, values', 'UniformOutput', false), ', ');
        optionMapStr = strjoin(optionPairs, ' | ');
        newRow = table("23년 상반기", string(qid), optionPairs(1), optionValuesStr, optionMapStr);
        disp(newRow);
    catch ME
        fprintf('Row %d error: %s\n', r, ME.message);
    end
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
