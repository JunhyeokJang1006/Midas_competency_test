file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
tbl = readtable(file, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');
for r = 1:1
    optionText = tbl{r,6};
    if iscell(optionText)
        optionText = optionText{1};
    end
    disp(optionText);
    tokens = regexp(optionText, '\\(([-+]?\\d+)\\)\\s*([^\\(\\)]+)', 'tokens');
    disp(tokens);
end
