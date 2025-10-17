officialFile = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
officialTable = readtable(officialFile, 'Sheet', 'Sheet2', 'VariableNamingRule', 'preserve');

for r = 1:1
    qid = officialTable{r,3}
    fprintf('Row %d
', r);
    optionText = officialTable{r,6};
    if iscell(optionText), optionText = optionText{1}; end
    optionChar = char(optionText);
    tokens = regexp(optionChar, '\\((\\d+)\\)\\s*([^()]+)', 'tokens');
    disp(tokens);
end
