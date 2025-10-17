file = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23_rebuilt.xlsx';
tbl = readtable(file);
if ~isa(tbl.SourceFile, 'string')
    sourceStrings = string(tbl.SourceFile);
else
    sourceStrings = tbl.SourceFile;
end
sources = unique(sourceStrings);
for i = 1:numel(sources)
    src = sources(i);
    count = sum(sourceStrings == src);
    fprintf('%d개 - %s\n', count, src);
end
