% Excel 저장 디버깅 스크립트
cd('D:\project\HR데이터\matlab\CSR');

% 테스트용 간단한 데이터로 Excel 저장 테스트
testTable = table({'A'; 'B'; 'C'}, [1; 2; 3], [0.1; 0.2; 0.3], ...
    'VariableNames', {'Category', 'Value1', 'Value2'});

outputDir = 'D:\project\HR데이터\결과\CSR';
testFileName = fullfile(outputDir, 'test_excel_save.xlsx');

try
    % 기본 저장 테스트
    writetable(testTable, testFileName, 'Sheet', 'TestSheet1');
    fprintf('✓ 기본 Excel 저장 성공\n');

    % 두 번째 시트 추가 테스트
    writetable(testTable, testFileName, 'Sheet', 'TestSheet2');
    fprintf('✓ 두 번째 시트 저장 성공\n');

    % 파일 확인
    [~,sheets] = xlsfinfo(testFileName);
    fprintf('생성된 시트: %s\n', strjoin(sheets, ', '));

catch ME
    fprintf('✗ Excel 저장 실패: %s\n', ME.message);
    fprintf('오류 세부사항: %s\n', getReport(ME));
end