existingPath = 'D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_from_questioninfo.xlsx';
tbl = readtable(existingPath);
disp(tbl.Properties.VariableNames);
fprintf('행 수: %d\n', height(tbl));
