filePath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년_상반기_역량진단_응답데이터.xlsx';
[~, sheets] = xlsfinfo(filePath);
results = {};
for s = 1:numel(sheets)
    sheetName = sheets{s};
    try
        tbl = readtable(filePath, 'Sheet', sheetName, 'VariableNamingRule', 'preserve');
    catch
        continue;
    end
    varNames = tbl.Properties.VariableNames;
    qIdx = find(startsWith(varNames,'Q') | startsWith(varNames,'q'));
    if isempty(qIdx)
        continue;
    end
    fprintf('\n=== Sheet: %s (Q columns: %d) ===\n', sheetName, numel(qIdx));
    for idx = 1:numel(qIdx)
        colName = varNames{qIdx(idx)};
        data = tbl{:, qIdx(idx)};
        if ~isnumeric(data)
            continue;
        end
        valid = data(~isnan(data));
        if isempty(valid)
            continue;
        end
        vals = unique(valid);
        fprintf('%s: min %.0f, max %.0f, unique %s\n', colName, min(vals), max(vals), mat2str(vals'));
    end
end
