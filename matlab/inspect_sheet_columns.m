files = {
    'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx',
    'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_하반기_역량진단_응답데이터.xlsx',
    'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/25년_상반기_역량진단_응답데이터.xlsx'
};

for i = 1:numel(files)
    filePath = files{i};
    [~, sheets] = xlsfinfo(filePath);
    fprintf('\n=== %s ===\n', filePath);
    for s = 1:numel(sheets)
        sheetName = sheets{s};
        try
            tbl = readtable(filePath, 'Sheet', sheetName, 'ReadVariableNames', true, 'ReadRowNames', false, 'VariableNamingRule', 'preserve');
            varNames = tbl.Properties.VariableNames;
            qIdx = find(startsWith(varNames, 'Q'));
            fprintf('  Sheet %d: %s -> Q columns: %d\n', s, sheetName, numel(qIdx));
            if ~isempty(qIdx)
                firstNames = varNames(qIdx(1:min(3,end)));
                for k = 1:numel(firstNames)
                    fprintf('    %s\n', firstNames{k});
                end
            end
        catch ME
            fprintf('  Sheet %d: %s -> read failed (%s)\n', s, sheetName, ME.message);
        end
    end
end
