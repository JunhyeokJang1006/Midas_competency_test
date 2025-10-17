% MAT 파일만 저장하는 스크립트
cd('D:\project\HR데이터\matlab\CSR');

try
    % 기존 분석 결과를 다시 실행해서 MAT 파일 생성
    run('corr_CSR_vs_comp_test.m');
    fprintf('✓ MAT 파일 저장 완료\n');
catch ME
    fprintf('✗ MAT 파일 저장 실패: %s\n', ME.message);
end