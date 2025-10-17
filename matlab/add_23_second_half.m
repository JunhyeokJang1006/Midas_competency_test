inputPath = "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/question_scale_metadata_with_23.xlsx";
sourcePeriod = "24년 상반기";
targetPeriod = "23년 하반기";
newSourceFile = "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터/23년_하반기_역량진단_응답데이터.xlsx";

tbl = readtable(inputPath);
sourceMask = tbl.Period == sourcePeriod;
if ~any(sourceMask)
    error('원본 기간 %s 데이터를 찾을 수 없습니다.', sourcePeriod);
end

sourceRows = tbl(sourceMask, :);
sourceRows.Period = repmat(string(targetPeriod), height(sourceRows), 1);
sourceRows.SourceFile = repmat(string(newSourceFile), height(sourceRows), 1);

% 이미 추가된 23년 하반기 데이터가 있는 경우 제거
existingMask = tbl.Period == targetPeriod;
tbl(existingMask, :) = [];

mergedTable = [tbl; sourceRows];

% Period, QuestionID 기준으로 정렬 (가독성 용)
if ismember('Period', mergedTable.Properties.VariableNames) && ismember('QuestionID', mergedTable.Properties.VariableNames)
    mergedTable = sortrows(mergedTable, {'Period','QuestionID'});
end

writetable(mergedTable, inputPath, 'Sheet', 'QuestionScales');
fprintf('✓ 23년 하반기 문항 척도 %d건 추가 완료 (%s 기준 복제).\n', height(sourceRows), sourcePeriod);
