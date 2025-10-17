function logger(level, message, varargin)
    % LOGGER 분석 로그 기록 함수
    %
    % 입력:
    %   level - 'INFO', 'WARNING', 'ERROR', 'DEBUG'
    %   message - 로그 메시지 (sprintf 형식 가능)
    %   varargin - sprintf 인수들
    %
    % 사용 예:
    %   logger('INFO', '데이터 로드 완료: %d명', size(data, 1));
    %   logger('WARNING', '결측값 발견: %d개', sum(isnan(data(:))));
    %   logger('ERROR', '파일을 찾을 수 없음: %s', filename);
    %   logger('DEBUG', 'Bootstrap 반복: %d/%d', i, n);
    %
    % 특징:
    %   - 콘솔과 파일에 동시 출력
    %   - 타임스탬프 자동 추가
    %   - 로그 레벨별 구분
    %   - UTF-8 인코딩 지원
    %   - 오류 발생 시에도 분석 중단하지 않음
    %
    % 작성일: 2025-10-15
    % 리팩토링: Phase 2 - 로깅 시스템 추가

    persistent log_file;
    persistent log_dir;

    % 첫 호출 시 로그 파일 생성
    if isempty(log_file)
        % 로그 디렉토리 설정
        log_dir = 'D:\project\HR데이터\matlab\자가불소_revised_refact\logs';
        if ~exist(log_dir, 'dir')
            try
                mkdir(log_dir);
            catch
                % 디렉토리 생성 실패 시 현재 디렉토리 사용
                log_dir = pwd;
            end
        end

        % 로그 파일명 생성 (타임스탬프 포함)
        timestamp = datestr(now, 'yyyymmdd_HHMMSS');
        log_file = fullfile(log_dir, sprintf('analysis_log_%s.txt', timestamp));

        % 로그 파일 초기화 (헤더 작성)
        try
            fid = fopen(log_file, 'w', 'n', 'UTF-8');
            if fid > 0
                fprintf(fid, '========================================\n');
                fprintf(fid, '  Analysis Log\n');
                fprintf(fid, '  Start Time: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
                fprintf(fid, '========================================\n\n');
                fclose(fid);
            end
        catch
            % 로그 파일 생성 실패 시에도 계속 진행
        end
    end

    % 입력 검증
    if nargin < 2
        warning('logger:InvalidInput', 'logger requires at least 2 arguments (level and message)');
        return;
    end

    % 레벨 검증
    valid_levels = {'INFO', 'WARNING', 'ERROR', 'DEBUG'};
    if ~ismember(upper(level), valid_levels)
        level = 'INFO';  % 기본값
    else
        level = upper(level);
    end

    % 메시지 포맷팅
    try
        if nargin > 2
            formatted_msg = sprintf(message, varargin{:});
        else
            formatted_msg = message;
        end
    catch ME
        % 포맷팅 실패 시 원본 메시지 사용
        formatted_msg = message;
        warning('logger:FormattingError', 'Message formatting failed: %s', ME.message);
    end

    % 타임스탬프 생성
    timestamp_str = datestr(now, 'yyyy-mm-dd HH:MM:SS');

    % 로그 엔트리 생성
    log_entry = sprintf('[%s] [%-7s] %s\n', timestamp_str, level, formatted_msg);

    % 콘솔 출력 (레벨별 색상 구분)
    try
        switch level
            case 'ERROR'
                fprintf(2, log_entry);  % stderr (빨간색)
            case 'WARNING'
                fprintf(2, log_entry);  % stderr (빨간색)
            otherwise
                fprintf(log_entry);      % stdout
        end
    catch
        % 콘솔 출력 실패 시에도 계속 진행
    end

    % 파일에 저장
    try
        fid = fopen(log_file, 'a', 'n', 'UTF-8');
        if fid > 0
            fprintf(fid, log_entry);
            fclose(fid);
        end
    catch ME
        % 파일 쓰기 실패 시 경고만 출력 (첫 실패 시만)
        persistent write_error_shown;
        if isempty(write_error_shown)
            warning('logger:FileWriteError', '로그 파일 쓰기 실패: %s', ME.message);
            write_error_shown = true;
        end
    end
end
