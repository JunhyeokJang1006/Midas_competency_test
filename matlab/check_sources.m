file = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23_rebuilt.xlsx';
if exist(file, 'file') ~= 2
    disp('파일이 존재하지 않습니다.');
    return;
end

tbl = readtable(file);
sources = unique(tbl.SourceFile);
for i = 1:numel(sources)
    fprintf('%d) %s\n', i, sources{i});
end
fprintf('총 행 수: %d\n', height(tbl));
