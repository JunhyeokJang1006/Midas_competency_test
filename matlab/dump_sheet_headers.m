filePath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
[status, sheets] = xlsfinfo(filePath);
if isempty(sheets)
    error('시트 없음');
end
for idx = 1:numel(sheets)
    sheetName = sheets{idx};
    fprintf('Sheet %d name: %s\n', idx, sheetName);
    try
        raw = readcell(filePath, 'Sheet', sheetName);
        firstRow = raw(1, :);
        nonEmpty = ~cellfun(@(x) (isempty(x) || (isstring(x)&&strlength(x)==0)), firstRow);
        cols = find(nonEmpty);
        if ~isempty(cols)
            fprintf('  First non-empty column header: %s\n', firstRow{cols(1)});
        else
            fprintf('  First row empty\n');
        end
    catch ME
        fprintf('  readcell 실패: %s\n', ME.message);
    end
end
