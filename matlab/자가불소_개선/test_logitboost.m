%% LogitBoost 분석 시스템 테스트 스크립트
% 이 스크립트는 competency_statistical_analysis_logitboost_only.m을 테스트합니다.

clear; clc; close all;

fprintf('═══════════════════════════════════════════════════════════\n');
fprintf('    LogitBoost HR 분석 시스템 테스트\n');
fprintf('═══════════════════════════════════════════════════════════\n\n');

% 1. 파일 존재 여부 확인
fprintf('[1] 필수 파일 확인\n');
fprintf('────────────────────────────────────────────\n');

hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
main_script = 'competency_statistical_analysis_logitboost_only.m';

files_to_check = {hr_file, comp_file, main_script};
file_names = {'HR 데이터', '역량검사 데이터', '메인 스크립트'};

all_files_exist = true;
for i = 1:length(files_to_check)
    if exist(files_to_check{i}, 'file')
        fprintf('  ✓ %s: 존재\n', file_names{i});
    else
        fprintf('  ✗ %s: 없음\n', file_names{i});
        all_files_exist = false;
    end
end

if ~all_files_exist
    error('필수 파일이 없습니다. 파일 경로를 확인해주세요.');
end

% 2. 저장된 모델 확인
fprintf('\n[2] 저장된 모델 확인\n');
fprintf('────────────────────────────────────────────\n');

model_file = fullfile(pwd, 'trained_logitboost_model.mat');
if exist(model_file, 'file')
    fprintf('  ✓ 저장된 모델 발견: %s\n', model_file);
    
    % 모델 정보 출력
    try
        model_info = load(model_file);
        fprintf('    - 학습 날짜: %s\n', model_info.training_date);
        fprintf('    - 특징 수: %d\n', size(model_info.X_train_info, 2));
        fprintf('    - CV 정확도: %.3f\n', model_info.best_logit_params.CV_Accuracy);
    catch
        fprintf('    - 모델 정보를 읽을 수 없습니다.\n');
    end
    
    % 모델 사용 여부 선택
    fprintf('\n  저장된 모델을 사용하시겠습니까? (Y/N): ');
    user_input = input('', 's');
    if strcmpi(user_input, 'N')
        fprintf('  → 새로운 모델을 학습합니다.\n');
        delete(model_file);
    else
        fprintf('  → 저장된 모델을 사용합니다.\n');
    end
else
    fprintf('  ✗ 저장된 모델 없음 (새로 학습 필요)\n');
end

% 3. 메인 스크립트 실행
fprintf('\n[3] 메인 분석 실행\n');
fprintf('────────────────────────────────────────────\n');
fprintf('  분석을 시작합니다...\n\n');

try
    % 타이머 시작
    tic;
    
    % 메인 스크립트 실행
    run(main_script);
    
    % 실행 시간 출력
    elapsed_time = toc;
    fprintf('\n\n═══════════════════════════════════════════════════════════\n');
    fprintf('  분석 완료! (소요 시간: %.2f초)\n', elapsed_time);
    fprintf('═══════════════════════════════════════════════════════════\n');
    
catch ME
    fprintf('\n\n═══════════════════════════════════════════════════════════\n');
    fprintf('  ✗ 오류 발생!\n');
    fprintf('═══════════════════════════════════════════════════════════\n');
    fprintf('오류 메시지: %s\n', ME.message);
    fprintf('오류 위치: %s (Line %d)\n', ME.stack(1).file, ME.stack(1).line);
    
    % 상세 오류 정보
    fprintf('\n상세 정보:\n');
    for i = 1:min(3, length(ME.stack))
        fprintf('  [%d] %s (Line %d)\n', i, ME.stack(i).name, ME.stack(i).line);
    end
    
    % 디버깅 팁
    fprintf('\n디버깅 팁:\n');
    if contains(ME.message, 'Undefined')
        fprintf('  → 변수가 정의되지 않았습니다. 데이터 로딩을 확인하세요.\n');
    elseif contains(ME.message, 'Index')
        fprintf('  → 인덱스 오류입니다. 데이터 크기를 확인하세요.\n');
    elseif contains(ME.message, 'Matrix')
        fprintf('  → 행렬 차원 오류입니다. 데이터 형태를 확인하세요.\n');
    elseif contains(ME.message, 'File')
        fprintf('  → 파일 관련 오류입니다. 파일 경로를 확인하세요.\n');
    else
        fprintf('  → 일반 오류입니다. 오류 메시지를 확인하세요.\n');
    end
end

% 4. 결과 파일 확인
fprintf('\n[4] 생성된 파일 확인\n');
fprintf('────────────────────────────────────────────\n');

% 현재 폴더의 관련 파일들 찾기
excel_files = dir('hr_analysis_logitboost_*.xlsx');
mat_files = dir('hr_analysis_logitboost_*.mat');
png_files = dir('radar_chart_*.png');

if ~isempty(excel_files)
    fprintf('  Excel 보고서: %d개\n', length(excel_files));
    for i = 1:min(3, length(excel_files))
        fprintf('    - %s\n', excel_files(i).name);
    end
end

if ~isempty(mat_files)
    fprintf('  MATLAB 파일: %d개\n', length(mat_files));
    for i = 1:min(3, length(mat_files))
        fprintf('    - %s\n', mat_files(i).name);
    end
end

if ~isempty(png_files)
    fprintf('  레이더 차트: %d개\n', length(png_files));
end

fprintf('\n═══════════════════════════════════════════════════════════\n');
fprintf('  테스트 완료\n');
fprintf('═══════════════════════════════════════════════════════════\n');





