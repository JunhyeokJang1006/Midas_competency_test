function config = load_config()
    % LOAD_CONFIG 분석 설정값 로드
    %
    % 출력:
    %   config - 구조체, 모든 설정값 포함
    %
    % 사용 예:
    %   cfg = load_config();
    %   data = readtable(cfg.hr_file);
    %
    % 작성일: 2025-10-15
    % 리팩토링: Phase 1 - 설정 파일 분리

    %% ======================================================================
    %                           경로 설정
    % =======================================================================

    % 기본 디렉토리
    config.base_dir = 'D:\project\HR데이터';

    % 입력 파일 경로
    config.hr_file = fullfile(config.base_dir, '데이터', '역량검사 요청 정보', ...
        '최근 3년 입사자_인적정보_revised.xlsx');
    config.comp_file = fullfile(config.base_dir, '데이터', '역량검사 요청 정보', ...
        '23-25년 역량검사_개발자추가_filtered.xlsx');

    % 출력 디렉토리
    config.output_dir = fullfile(config.base_dir, '결과', '자가불소_revised_talent');

    % Bootstrap 차트 파일 (기존 하드코딩된 경로를 설정으로 이동)
    config.bootstrap_chart_file = fullfile(config.base_dir, '결과', '자가불소', 'bootstrap.xlsx');

    % 타임스탬프 및 모델 파일
    config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
    config.model_file = fullfile(config.output_dir, 'trained_logitboost_model.mat');

    % 출력 디렉토리 생성
    if ~exist(config.output_dir, 'dir')
        mkdir(config.output_dir);
    end

    %% ======================================================================
    %                           분석 설정
    % =======================================================================

    % 기준선 계산 방식
    config.baseline_type = 'weighted';  % 'simple' 또는 'weighted' 선택

    % 극단 그룹 비교 방식
    config.extreme_group_method = 'all';  % 'extreme' 또는 'all' 선택
    % 'extreme': 가장 확실한 케이스만 (자연성, 성실한 가연성 vs 무능한 불연성, 소화성)
    % 'all': 모든 고성과자 vs 저성과자 비교

    % Bootstrap 및 Permutation 설정
    config.n_bootstrap = 5000;  % Bootstrap 재샘플링 횟수
    config.n_permutation_max = 5000;  % 최대 퍼뮤테이션 횟수 (작은 데이터셋)
    config.n_permutation_mid = 2000;  % 중간 퍼뮤테이션 횟수 (중간 데이터셋)
    config.n_permutation_min = 1000;  % 최소 퍼뮤테이션 횟수 (큰 데이터셋)

    % 개발자 필터링 옵션
    config.developer_filter = 'all';  % 'all', 'exclude_dev', 'only_dev'

    %% ======================================================================
    %                           모델 설정
    % =======================================================================

    config.use_saved_model = false;  % 저장된 모델 사용 여부
    config.force_recalc_permutation = false;  % 퍼뮤테이션 강제 재계산 여부
    config.random_seed = 42;  % 난수 시드 (재현성 보장)

    %% ======================================================================
    %                           이상치 제거 설정
    % =======================================================================

    config.outlier_removal = struct();
    config.outlier_removal.enabled = true;  % 이상치 제거 활성화 여부
    config.outlier_removal.method = 'none';  % 'iqr', 'zscore', 'percentile', 'none' 중 선택
    config.outlier_removal.iqr_multiplier = 1.5;  % IQR 방법의 배수 (기본값: 1.5)
    config.outlier_removal.zscore_threshold = 3;  % Z-score 방법의 임계값 (기본값: 3)
    config.outlier_removal.percentile_bounds = [5, 95];  % 백분위수 방법의 범위 (기본값: 5%, 95%)
    config.outlier_removal.apply_to_competencies = true;  % 역량 점수에 적용 여부
    config.outlier_removal.report_outliers = true;  % 이상치 제거 결과 보고 여부

    %% ======================================================================
    %                           파일 관리 설정
    % =======================================================================

    config.create_backup = true;  % 백업 폴더 생성 여부
    config.backup_folder = 'backup';  % 백업 폴더 이름
    config.use_timestamp = false;  % 파일명에 타임스탬프 사용 여부

    %% ======================================================================
    %                           성과 순위 및 그룹 정의
    % =======================================================================

    % 성과 순위 정의 (위장형 소화성만 제외)
    config.performance_ranking = containers.Map(...
        {'자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
         '게으른 가연성', '무능한 불연성', '소화성'}, ...
        [8, 7, 6, 5, 4, 3, 1]);

    % 순서형 로지스틱 회귀용 레벨 정의 (1: 최하위, 7: 최상위)
    config.ordinal_levels = containers.Map(...
        {'소화성', '무능한 불연성', '게으른 가연성', '유능한 불연성', ...
         '유익한 불연성', '성실한 가연성', '자연성'}, ...
        [1, 2, 3, 4, 5, 6, 7]);

    % 레벨 이름 (순서대로)
    config.level_names = {'소화성', '무능한 불연성', '게으른 가연성', '유능한 불연성', ...
                         '유익한 불연성', '성실한 가연성', '자연성'};

    % 고성과자/저성과자 정의 (이진 분류용) - 유능한 불연성 제외
    config.high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
    config.low_performers = {'게으른 가연성', '무능한 불연성', '소화성'};
    config.excluded_from_analysis = {'유능한 불연성'};  % 분석에서 제외
    config.excluded_types = {'위장형 소화성'};  % 위장형 소화성도 제외

    %% ======================================================================
    %                           설정 검증
    % =======================================================================

    % 파일 존재 확인 (경고만 출력, 오류는 발생시키지 않음)
    if ~exist(config.hr_file, 'file')
        warning('load_config:FileNotFound', 'HR 데이터 파일을 찾을 수 없습니다: %s', config.hr_file);
    end

    if ~exist(config.comp_file, 'file')
        warning('load_config:FileNotFound', '역량검사 데이터 파일을 찾을 수 없습니다: %s', config.comp_file);
    end

    % 설정값 유효성 검증
    assert(ismember(config.baseline_type, {'simple', 'weighted'}), ...
        'baseline_type은 "simple" 또는 "weighted"여야 합니다.');

    assert(ismember(config.extreme_group_method, {'extreme', 'all'}), ...
        'extreme_group_method는 "extreme" 또는 "all"이어야 합니다.');

    assert(ismember(config.outlier_removal.method, {'iqr', 'zscore', 'percentile', 'none'}), ...
        'outlier_removal.method는 "iqr", "zscore", "percentile", "none" 중 하나여야 합니다.');

    assert(config.n_bootstrap > 0, 'n_bootstrap은 양수여야 합니다.');
    assert(config.random_seed >= 0, 'random_seed는 0 이상이어야 합니다.');

end
