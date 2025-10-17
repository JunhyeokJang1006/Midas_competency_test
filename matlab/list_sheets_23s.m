filePath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년_상반기_역량진단_응답데이터.xlsx';
[status,sheets] = xlsfinfo(filePath);
if isempty(sheets)
    error('시트 없음');
end
for idx=1:numel(sheets)
    fprintf('Sheet %d: %s\n', idx, sheets{idx});
end
