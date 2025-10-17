filePath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년_상반기_역량진단_응답데이터.xlsx';
[~, sheets] = xlsfinfo(filePath);
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
    fprintf('\n=== Sheet: %s ===\n', sheetName);
    keys = {};
    counts = [];
    for idx = 1:numel(qIdx)
        colData = tbl{:, qIdx(idx)};
        if ~isnumeric(colData)
            continue;
        end
        data = colData(~isnan(colData));
        if isempty(data)
            continue;
        end
        vals = unique(data)';
        key = mat2str(vals);
        matchIdx = find(strcmp(keys, key), 1);
        if isempty(matchIdx)
            keys{end+1} = key; %#ok<AGROW>
            counts(end+1) = 1; %#ok<AGROW>
        else
            counts(matchIdx) = counts(matchIdx) + 1;
        end
    end
    for k = 1:numel(keys)
        fprintf('%3d 문항 -> %s\n', counts(k), keys{k});
    end
end
