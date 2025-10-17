fileName = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
tbl = readtable(fileName, 'Sheet', '하향진단', 'VariableNamingRule', 'preserve');
disp(tbl.Properties.VariableNames(1:20));
