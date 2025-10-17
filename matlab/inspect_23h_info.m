file = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년 하반기 문항정보.xlsx';
tbl = readtable(file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
disp(tbl.Properties.VariableNames);
disp(tbl(1:min(5,height(tbl)), :));
