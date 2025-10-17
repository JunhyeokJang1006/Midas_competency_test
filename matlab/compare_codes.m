file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
officialTable = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
optSan = sanitizeCellValue(officialTable{1,6});
sanCharCodes = double(char(optSan));
fprintf('San codes: ');
fprintf('%d ', sanCharCodes);
fprintf('\n');

optRaw = officialTable{1,6}; if iscell(optRaw), optRaw = optRaw{1}; end
rawCharCodes = double(char(optRaw));
fprintf('Raw codes: ');
fprintf('%d ', rawCharCodes);
fprintf('\n');

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
