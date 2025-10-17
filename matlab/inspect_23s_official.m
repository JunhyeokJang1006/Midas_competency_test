file = 'C:/Users/MIDASIT/Downloads/23년 상반기 문항정보 1.xlsx';
[status,sheets] = xlsfinfo(file);
if isempty(sheets)
    error('시트 없음');
end
for i = 1:numel(sheets)
    fprintf('Sheet %d: %s\n', i, sheets{i});
end
