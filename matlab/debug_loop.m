officialFile = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
officialTable = readtable(officialFile, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');

for r = 1:5
    try
        qid = sanitizeCellValue(officialTable{r,3});
        optionText = sanitizeCellValue(officialTable{r,6});
        optionChar = char(optionText);
        tokens = regexp(optionChar, '\\((\\d+)\\)\\s*([^()]+)', 'tokens');
        fprintf('Row %d QID %s tokens %d\n', r, qid, numel(tokens));
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
