file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
foo = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
raw = foo{1,6}; if iscell(raw), raw = raw{1}; end
optSan = sanitizeCellValue(foo{1,6});
optChar = char(optSan);
rawChar = char(raw);
fprintf('class optSan: %s\n', class(optSan));
fprintf('size optChar: %dx%d\n', size(optChar,1), size(optChar,2));
fprintf('size rawChar: %dx%d\n', size(rawChar,1), size(rawChar,2));

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
