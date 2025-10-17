file = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
tbl = readtable(file, 'Sheet', '타인진단', 'VariableNamingRule', 'preserve');
varNames = tbl.Properties.VariableNames;
for c = 1:min(10, numel(varNames))
    fprintf('%d: %s\n', c, varNames{c});
end

colIdx = find(strcmp(varNames, 'Q4리커트지난반기동안, OOO님이맡은업무와역할은조직성과에얼마나중요한것이었나요(0) 덜중요한편이다(1) 보통이다(2) 중요한편이다(3) 중요하다(4) 매우중요하다'));
if ~isempty(colIdx)
    colData = tbl{:, colIdx};
    disp(unique(colData(~isnan(colData))));
else
    fprintf('Q4 컬럼을 찾을 수 없습니다.\n');
end
