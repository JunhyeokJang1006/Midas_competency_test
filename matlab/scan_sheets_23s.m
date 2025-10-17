filePath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년_상반기_역량진단_응답데이터.xlsx';
[~, sheets] = xlsfinfo(filePath);
for s = 1:numel(sheets)
    sheetName = sheets{s};
    try
        tbl = readtable(filePath, 'Sheet', sheetName, 'VariableNamingRule', 'preserve');
        varNames = tbl.Properties.VariableNames;
        qIdx = find(startsWith(varNames, 'Q') | startsWith(varNames, 'q'));
        fprintf('Sheet %d: %s -> Q columns: %d\n', s, sheetName, numel(qIdx));
        if ~isempty(qIdx)
            fprintf('  첫 5개 문항: ');
            for k=1:min(5,numel(qIdx))
                fprintf('%s ', varNames{qIdx(k)});
            end
            fprintf('\n');
        end
    catch ME
        fprintf('Sheet %d: %s -> 읽기 실패 (%s)\n', s, sheetName, ME.message);
    end
end
