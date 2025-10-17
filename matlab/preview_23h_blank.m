tbl = readtable('D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23_rebuilt.xlsx');
mask = (tbl.Period == "23년 하반기") & (tbl.OptionMap == "");
fprintf('빈 OptionMap 행 수: %d\n', sum(mask));
if any(mask)
    disp(tbl(mask, :));
end
