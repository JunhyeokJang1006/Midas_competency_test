file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
foo = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
optSan = sanitizeCellValue(foo{1,6});
optChar = char(optSan);
tokens = regexp(optChar, '\\((\\d+)\\)\\s*([^()]+)', 'tokens');
if isempty(tokens)
    disp('empty');
else
    disp(tokens);
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
