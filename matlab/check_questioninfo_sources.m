file = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23_rebuilt.xlsx';
tbl = readtable(file);
if ~isa(tbl.SourceFile,'string')
    sourceStrings = string(tbl.SourceFile);
else
    sourceStrings = tbl.SourceFile;
end
mask = contains(sourceStrings, '문항정보');
uniqueSources = unique(sourceStrings(mask));
disp(uniqueSources);
