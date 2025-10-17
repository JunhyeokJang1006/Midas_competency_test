T = readtable('C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx', 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
optRaw = T{1,6}; if iscell(optRaw), optRaw = optRaw{1}; end
matches = regexp(char(optRaw), '\\((\\d+)\\)\\s*[^()]+', 'match');
for i=1:numel(matches)
    fprintf('%s\n', matches{i});
end
