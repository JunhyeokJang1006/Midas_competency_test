%% test_config.m
% load_config 함수 테스트 스크립트
% 작성일: 2025-10-15

fprintf('========================================\n');
fprintf('  load_config 함수 테스트\n');
fprintf('========================================\n\n');

%% 테스트 1: 함수 호출 및 기본 구조 확인
fprintf('【테스트 1】 함수 호출 및 기본 구조 확인\n');
fprintf('────────────────────────────────────────\n');

try
    config = load_config();
    fprintf('✅ load_config() 호출 성공\n');
catch ME
    fprintf('❌ load_config() 호출 실패: %s\n', ME.message);
    return;
end

%% 테스트 2: 필수 필드 존재 확인
fprintf('\n【테스트 2】 필수 필드 존재 확인\n');
fprintf('────────────────────────────────────────\n');

required_fields = {'hr_file', 'comp_file', 'output_dir', 'n_bootstrap', ...
    'outlier_removal', 'performance_ranking', 'high_performers', 'low_performers'};

all_fields_exist = true;
for i = 1:length(required_fields)
    if isfield(config, required_fields{i})
        fprintf('✅ %s: 존재\n', required_fields{i});
    else
        fprintf('❌ %s: 없음\n', required_fields{i});
        all_fields_exist = false;
    end
end

if all_fields_exist
    fprintf('\n✅ 모든 필수 필드 존재\n');
else
    fprintf('\n❌ 일부 필수 필드 누락\n');
    return;
end

%% 테스트 3: 경로 설정 확인
fprintf('\n【테스트 3】 경로 설정 확인\n');
fprintf('────────────────────────────────────────\n');

fprintf('HR 파일: %s\n', config.hr_file);
if exist(config.hr_file, 'file')
    fprintf('  ✅ 파일 존재\n');
else
    fprintf('  ⚠️  파일 없음 (경고만 출력)\n');
end

fprintf('역량검사 파일: %s\n', config.comp_file);
if exist(config.comp_file, 'file')
    fprintf('  ✅ 파일 존재\n');
else
    fprintf('  ⚠️  파일 없음 (경고만 출력)\n');
end

fprintf('출력 디렉토리: %s\n', config.output_dir);
if exist(config.output_dir, 'dir')
    fprintf('  ✅ 디렉토리 존재\n');
else
    fprintf('  ⚠️  디렉토리 없음 (자동 생성됨)\n');
end

%% 테스트 4: 분석 설정 확인
fprintf('\n【테스트 4】 분석 설정 확인\n');
fprintf('────────────────────────────────────────\n');

fprintf('▶ Bootstrap 횟수: %d\n', config.n_bootstrap);
assert(config.n_bootstrap == 5000, '❌ n_bootstrap 값 불일치');
fprintf('  ✅ 기대값(5000)과 일치\n');

fprintf('▶ 극단 그룹 방식: %s\n', config.extreme_group_method);
assert(strcmp(config.extreme_group_method, 'all'), '❌ extreme_group_method 값 불일치');
fprintf('  ✅ 기대값(all)과 일치\n');

fprintf('▶ 기준선 타입: %s\n', config.baseline_type);
assert(strcmp(config.baseline_type, 'weighted'), '❌ baseline_type 값 불일치');
fprintf('  ✅ 기대값(weighted)과 일치\n');

fprintf('▶ 난수 시드: %d\n', config.random_seed);
assert(config.random_seed == 42, '❌ random_seed 값 불일치');
fprintf('  ✅ 기대값(42)과 일치\n');

%% 테스트 5: 이상치 제거 설정 확인
fprintf('\n【테스트 5】 이상치 제거 설정 확인\n');
fprintf('────────────────────────────────────────\n');

fprintf('▶ 이상치 제거 활성화: %d\n', config.outlier_removal.enabled);
fprintf('▶ 방법: %s\n', config.outlier_removal.method);
assert(strcmp(config.outlier_removal.method, 'none'), '❌ outlier_removal.method 값 불일치');
fprintf('  ✅ 기대값(none)과 일치\n');

fprintf('▶ IQR 배수: %.1f\n', config.outlier_removal.iqr_multiplier);
assert(config.outlier_removal.iqr_multiplier == 1.5, '❌ iqr_multiplier 값 불일치');
fprintf('  ✅ 기대값(1.5)과 일치\n');

%% 테스트 6: 성과 순위 및 그룹 정의 확인
fprintf('\n【테스트 6】 성과 순위 및 그룹 정의 확인\n');
fprintf('────────────────────────────────────────\n');

fprintf('▶ 성과 순위 맵 크기: %d\n', config.performance_ranking.Count);
assert(config.performance_ranking.Count == 7, '❌ performance_ranking 크기 불일치');
fprintf('  ✅ 기대값(7)과 일치\n');

fprintf('▶ 고성과자 그룹 수: %d\n', length(config.high_performers));
assert(length(config.high_performers) == 3, '❌ high_performers 개수 불일치');
fprintf('  ✅ 기대값(3)과 일치\n');

fprintf('▶ 저성과자 그룹 수: %d\n', length(config.low_performers));
assert(length(config.low_performers) == 3, '❌ low_performers 개수 불일치');
fprintf('  ✅ 기대값(3)과 일치\n');

%% 테스트 7: 파일 관리 설정 확인
fprintf('\n【테스트 7】 파일 관리 설정 확인\n');
fprintf('────────────────────────────────────────\n');

fprintf('▶ 백업 생성: %d\n', config.create_backup);
assert(config.create_backup == true, '❌ create_backup 값 불일치');
fprintf('  ✅ 기대값(true)과 일치\n');

fprintf('▶ 백업 폴더: %s\n', config.backup_folder);
assert(strcmp(config.backup_folder, 'backup'), '❌ backup_folder 값 불일치');
fprintf('  ✅ 기대값(backup)과 일치\n');

fprintf('▶ 타임스탬프 사용: %d\n', config.use_timestamp);
assert(config.use_timestamp == false, '❌ use_timestamp 값 불일치');
fprintf('  ✅ 기대값(false)과 일치\n');

%% 테스트 8: 구조체 필드 접근 테스트
fprintf('\n【테스트 8】 구조체 필드 접근 테스트\n');
fprintf('────────────────────────────────────────\n');

try
    % 경로 접근
    hr_file = config.hr_file;
    fprintf('✅ config.hr_file 접근 성공\n');

    % 중첩 구조체 접근
    outlier_method = config.outlier_removal.method;
    fprintf('✅ config.outlier_removal.method 접근 성공\n');

    % containers.Map 접근
    자연성_순위 = config.performance_ranking('자연성');
    assert(자연성_순위 == 8, '❌ 자연성 순위 불일치');
    fprintf('✅ config.performance_ranking(''자연성'') = %d\n', 자연성_순위);

    % 셀 배열 접근
    first_high = config.high_performers{1};
    fprintf('✅ config.high_performers{1} = %s\n', first_high);

catch ME
    fprintf('❌ 구조체 접근 실패: %s\n', ME.message);
    return;
end

%% 최종 결과
fprintf('\n========================================\n');
fprintf('  🎉 모든 테스트 통과!\n');
fprintf('========================================\n');

fprintf('\n【설정 요약】\n');
fprintf('────────────────────────────────────────\n');
fprintf('Bootstrap 횟수: %d\n', config.n_bootstrap);
fprintf('Permutation 최대: %d\n', config.n_permutation_max);
fprintf('극단 그룹 방식: %s\n', config.extreme_group_method);
fprintf('기준선 타입: %s\n', config.baseline_type);
fprintf('이상치 제거: %s (%s)\n', ...
    iif(config.outlier_removal.enabled, '활성화', '비활성화'), ...
    config.outlier_removal.method);
fprintf('난수 시드: %d\n', config.random_seed);
fprintf('────────────────────────────────────────\n');

%% 헬퍼 함수
function result = iif(condition, true_value, false_value)
    % 간단한 삼항 연산자 함수
    if condition
        result = true_value;
    else
        result = false_value;
    end
end
