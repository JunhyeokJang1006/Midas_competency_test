% 간단한 CSV 저장 테스트
cd('D:\project\HR데이터\matlab\CSR');

% 테스트 데이터 생성
testData = [1, 0.5, 0.01; 2, -0.3, 0.25; 3, 0.8, 0.001];
categories = {'인지역량'; '감정역량'; '성실성'};

outputDir = 'D:\project\HR데이터\결과\CSR';
csvFileName = fullfile(outputDir, 'CSR_test_results.csv');

try
    % CSV 헤더 작성
    fid = fopen(csvFileName, 'w');
    fprintf(fid, 'Competency_Category,Correlation,P_Value\n');

    % 데이터 작성
    for i = 1:size(testData, 1)
        fprintf(fid, '%s,%.3f,%.3f\n', categories{i}, testData(i,2), testData(i,3));
    end

    fclose(fid);
    fprintf('✓ CSV 저장 성공: %s\n', csvFileName);

catch ME
    fprintf('✗ CSV 저장 실패: %s\n', ME.message);
    if exist('fid', 'var') && fid ~= -1
        fclose(fid);
    end
end