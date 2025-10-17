T = readtable('C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx', 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
optRaw = T{1,6};
if iscell(optRaw)
    optRaw = optRaw{1};
end
optChar = char(optRaw);
% direct
A = regexp(optChar, '\((\d+)\)\s*([^()]+)', 'tokens');
disp(A);
