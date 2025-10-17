tbl = readtable('D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23_rebuilt.xlsx');
mask = tbl.Period == "23년 하반기";
subset = tbl(mask, :);
fprintf('23년 하반기 문항 수: %d\n', height(subset));
if ~isempty(subset)
    disp(subset(1:min(10,height(subset)), :));
end
