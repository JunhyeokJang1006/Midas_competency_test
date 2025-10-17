file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
tbl = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
opt = tbl{1,6}; if iscell(opt), opt = opt{1}; end; optChar = char(opt);
tokens = regexp(optChar, '\((\d+)\)\s*([^()]+)', 'tokens');
for i=1:numel(tokens)
    fprintf('%s -> %s\n', tokens{i}{1}, tokens{i}{2});
end
