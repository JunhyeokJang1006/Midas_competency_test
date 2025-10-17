file = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
questionInfo = readtable(file, 'Sheet', '문항 정보_타인진단', 'VariableNamingRule', 'preserve');
disp(questionInfo.Properties.VariableNames);
disp(questionInfo(1:3,:));
