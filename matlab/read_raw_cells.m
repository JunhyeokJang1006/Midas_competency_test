filePath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
sheetName = '타인진단';
try
    raw = readcell(filePath, 'Sheet', sheetName);
catch ME
    fprintf('타인진단 시트 읽기 실패: %s\n', ME.message);
    [~, sheets] = xlsfinfo(filePath);
    fprintf('사용 가능한 시트:\n');
    disp(sheets');
    return;
end

firstRow = raw(1, 1:min(10, size(raw,2)));
for i = 1:numel(firstRow)
    val = firstRow{i};
    if isstring(val)
        val = char(val);
    elseif isnumeric(val)
        val = num2str(val);
    elseif ismissing(val)
        val = '';
    elseif ischar(val)
        % already char
    else
        val = ''; 
    end
    fprintf('%d: %s\n', i, val);
end
