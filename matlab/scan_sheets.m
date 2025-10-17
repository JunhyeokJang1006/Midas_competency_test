fileName = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
[status, sheets] = xlsfinfo(fileName);
for i = 1:numel(sheets)
    sheetName = sheets{i};
    try
        tbl = readtable(fileName, 'Sheet', sheetName, 'VariableNamingRule', 'preserve');
        varNames = tbl.Properties.VariableNames;
        numQ = sum(startsWith(varNames, 'Q') | startsWith(varNames, 'q'));
        fprintf('%d: %s -> vars=%d, Qcols=%d\n', i, sheetName, numel(varNames), numQ);
    catch ME
        fprintf('%d: %s -> 읽기 실패 (%s)\n', i, sheetName, ME.message);
    end
end
