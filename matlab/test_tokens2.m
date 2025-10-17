file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
tbl = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
opt = tbl{1,6};
if iscell(opt), opt = opt{1}; end
pattern = '\\(([-+]?\\d+)\\)\\s*([^\\(\\)]+?)(?=\\s*\\([-+]?\\d+\\)|$)';
tokens = regexp(opt, pattern, 'tokens');
disp(tokens);
