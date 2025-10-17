file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
tbl = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
opt = tbl{1,6};
if iscell(opt), opt = opt{1}; end
optChar = char(opt);
pos = strfind(optChar, '(');
if isempty(pos)
    fprintf('no ( found\n');
else
    fprintf('found ( at positions: ');
    fprintf('%d ', pos);
    fprintf('\n');
end
