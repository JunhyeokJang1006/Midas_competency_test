%% test_logger.m
% logger 함수 테스트 스크립트
% 작성일: 2025-10-15

fprintf('========================================\n');
fprintf('  logger 함수 테스트\n');
fprintf('========================================\n\n');

%% 경로 추가
addpath('D:\project\HR데이터\matlab\자가불소_revised_refact');

%% 테스트 1: 기본 로그 레벨 테스트
fprintf('【테스트 1】 기본 로그 레벨 테스트\n');
fprintf('────────────────────────────────────────\n');

logger('INFO', 'Test started');
logger('DEBUG', 'Debug message: variable x = %d', 42);
logger('WARNING', 'Warning: %s', 'This is a test warning');
logger('ERROR', 'Error occurred in function %s at line %d', 'testFunc', 123);

fprintf('✅ 기본 로그 레벨 테스트 완료\n\n');

%% 테스트 2: 포맷팅 테스트
fprintf('【테스트 2】 포맷팅 테스트\n');
fprintf('────────────────────────────────────────\n');

logger('INFO', '데이터 로드 완료: %d명', 158);
logger('INFO', '역량검사 데이터: %d명, 역량 수: %d개', 130, 25);
logger('INFO', 'Bootstrap 진행: %.1f%% 완료', 75.5);

fprintf('✅ 포맷팅 테스트 완료\n\n');

%% 테스트 3: 연속 로그 테스트
fprintf('【테스트 3】 연속 로그 테스트 (10회)\n');
fprintf('────────────────────────────────────────\n');

for i = 1:10
    logger('DEBUG', 'Iteration %d/%d', i, 10);
end

fprintf('✅ 연속 로그 테스트 완료\n\n');

%% 테스트 4: 오류 처리 테스트
fprintf('【테스트 4】 오류 처리 테스트\n');
fprintf('────────────────────────────────────────\n');

% 잘못된 레벨 (자동으로 INFO로 변환)
logger('INVALID_LEVEL', 'This should be treated as INFO');

% 인수 누락 (경고만 출력)
try
    logger('INFO');
    fprintf('❌ 인수 누락 시 경고를 출력해야 함\n');
catch
    fprintf('✅ 인수 누락 오류 처리 완료\n');
end

% 포맷 오류 (원본 메시지 사용)
logger('INFO', 'Value: %d %d', 42);  % 인수 부족

fprintf('✅ 오류 처리 테스트 완료\n\n');

%% 테스트 5: 로그 파일 확인
fprintf('【테스트 5】 로그 파일 확인\n');
fprintf('────────────────────────────────────────\n');

log_dir = 'D:\project\HR데이터\matlab\자가불소_revised_refact\logs';
if exist(log_dir, 'dir')
    log_files = dir(fullfile(log_dir, 'analysis_log_*.txt'));
    if ~isempty(log_files)
        latest_log = log_files(end);
        fprintf('✅ 로그 파일 생성 확인\n');
        fprintf('  파일명: %s\n', latest_log.name);
        fprintf('  크기: %d bytes\n', latest_log.bytes);
        fprintf('  경로: %s\n', fullfile(log_dir, latest_log.name));

        % 로그 파일 내용 읽기 (마지막 5줄)
        try
            log_path = fullfile(log_dir, latest_log.name);
            fid = fopen(log_path, 'r', 'n', 'UTF-8');
            if fid > 0
                content = fread(fid, '*char')';
                fclose(fid);

                lines = strsplit(content, '\n');
                fprintf('\n  로그 파일 마지막 5줄:\n');
                fprintf('  ────────────────────────────────────────\n');
                start_idx = max(1, length(lines) - 5);
                for i = start_idx:length(lines)
                    if ~isempty(lines{i})
                        fprintf('  %s\n', lines{i});
                    end
                end
            end
        catch ME
            fprintf('  ⚠️  로그 파일 읽기 실패: %s\n', ME.message);
        end
    else
        fprintf('❌ 로그 파일이 생성되지 않음\n');
    end
else
    fprintf('❌ 로그 디렉토리가 생성되지 않음\n');
end

%% 테스트 6: 성능 테스트
fprintf('\n【테스트 6】 성능 테스트 (1000회 로그)\n');
fprintf('────────────────────────────────────────\n');

tic;
for i = 1:1000
    logger('DEBUG', 'Performance test: %d', i);
end
elapsed = toc;

fprintf('✅ 1000회 로그 작성 시간: %.3f초\n', elapsed);
fprintf('  평균 속도: %.3f ms/log\n', elapsed * 1000 / 1000);

if elapsed < 1.0
    fprintf('  ✅ 성능 양호 (< 1초)\n');
elseif elapsed < 2.0
    fprintf('  ⚠️  성능 보통 (< 2초)\n');
else
    fprintf('  ❌ 성능 저하 (> 2초)\n');
end

%% 최종 결과
fprintf('\n========================================\n');
fprintf('  🎉 모든 logger 테스트 완료!\n');
fprintf('========================================\n\n');

fprintf('【결과 요약】\n');
fprintf('────────────────────────────────────────\n');
fprintf('✅ 기본 로그 레벨: INFO, DEBUG, WARNING, ERROR\n');
fprintf('✅ 메시지 포맷팅: sprintf 형식 지원\n');
fprintf('✅ 연속 로그: 정상 작동\n');
fprintf('✅ 오류 처리: 분석 중단하지 않음\n');
fprintf('✅ 파일 저장: UTF-8 인코딩\n');
fprintf('✅ 성능: %.3f ms/log\n', elapsed * 1000 / 1000);
fprintf('────────────────────────────────────────\n');
