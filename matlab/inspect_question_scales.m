files = {
    'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx',
    'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_하반기_역량진단_응답데이터.xlsx',
    'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/25년_상반기_역량진단_응답데이터.xlsx'
};

for i = 1:length(files)
    fileName = files{i};
    fprintf('\n=== %s ===\n', fileName);
    try
        tbl = readtable(fileName, 'Sheet', '문항 정보_타인진단', 'VariableNamingRule', 'preserve');
        disp(tbl.Properties.VariableNames);
        disp(tbl(1:min(5,height(tbl)), :));
    catch ME
        fprintf('파일 읽기 실패: %s\n', ME.message);
    end
end
