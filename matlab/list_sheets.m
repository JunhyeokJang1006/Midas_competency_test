fileName = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/24년_상반기_역량진단_응답데이터.xlsx';
[status, sheets] = xlsfinfo(fileName);
if status == "Microsoft Excel Spreadsheet"
    disp(sheets');
else
    fprintf('파일 정보를 가져올 수 없습니다. 상태: %s\n', status);
end
