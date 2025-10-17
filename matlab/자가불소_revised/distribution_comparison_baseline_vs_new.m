% =========================================================================
% 기존 학습 데이터 vs 2025년 하반기 신규 입사자 데이터 분포 비교
% =========================================================================
% 목적: 로지스틱 회귀 가중치를 새로운 데이터에 적용하기 전에
%       기존 데이터와 새 데이터의 분포 동질성 검증
%
% 비교 그룹:
%   1. 전체 참가자
%   2. 고성과자 (선발대상)
%   3. 저성과자 (비선발대상)
%
% 분석 내용:
%   - 역량별 박스플롯 비교
%   - 역량별 히스토그램 분포 비교
%   - 레이더 차트 프로필 비교
%   - 통계 검정 (KS-test, t-test, Levene's test)
% =========================================================================

clear; clc; close all;
rng(42, 'twister');

%% ========================================================================
%                          PART 1: 초기 설정
% =========================================================================

% 전역 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 10);
set(0, 'DefaultTextFontSize', 10);
set(0, 'DefaultLineLineWidth', 1.5);

fprintf('【분포 비교 분석】 기존 학습 데이터 vs 2025년 하반기 신규 입사자\n');
fprintf('========================================================================\n\n');

% 파일 경로 설정
config = struct();

% 기존 학습 데이터
config.baseline_hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.baseline_comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';

% 새로운 2025년 하반기 데이터
config.new_onboarding_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_온보딩 점수.xlsx';
config.new_comp_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_역검 점수.xlsx';

% 출력 디렉토리
config.output_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent\distribution_comparison';
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

% 타임스탬프
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 기존 데이터의 성과 그룹 정의
config.high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
config.low_performers = {'게으른 가연성', '무능한 불연성', '소화성'};
config.excluded_types = {'유능한 불연성', '위장형 소화성'};

%% ========================================================================
%                    PART 2: 기존 학습 데이터 로딩
% =========================================================================

fprintf('【STEP 1】 기존 학습 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

% 파일 존재 확인
if ~exist(config.baseline_hr_file, 'file')
    error('기존 HR 데이터를 찾을 수 없습니다: %s', config.baseline_hr_file);
end
if ~exist(config.baseline_comp_file, 'file')
    error('기존 역량검사 데이터를 찾을 수 없습니다: %s', config.baseline_comp_file);
end

% 기존 HR 데이터 로딩
baseline_hr = readtable(config.baseline_hr_file, 'VariableNamingRule', 'preserve');
fprintf('  ✓ 기존 HR 데이터: %d명\n', height(baseline_hr));

% 기존 역량검사 데이터 로딩
baseline_comp_upper = readtable(config.baseline_comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
fprintf('  ✓ 기존 역량검사 데이터: %d명\n', height(baseline_comp_upper));

% 기존 데이터 병합 (ID 기준)
baseline_merged = innerjoin(baseline_hr, baseline_comp_upper, 'Keys', 'ID');
fprintf('  ✓ 병합 후: %d명\n', height(baseline_merged));

% 제외 대상 필터링
if ismember('종합인재유형 (최종)', baseline_merged.Properties.VariableNames)
    talent_col = '종합인재유형 (최종)';
elseif ismember('종합인재유형', baseline_merged.Properties.VariableNames)
    talent_col = '종합인재유형';
else
    error('종합인재유형 컬럼을 찾을 수 없습니다.');
end

% 제외 대상 제거
excluded_mask = false(height(baseline_merged), 1);
for i = 1:length(config.excluded_types)
    excluded_mask = excluded_mask | strcmp(baseline_merged.(talent_col), config.excluded_types{i});
end
baseline_merged = baseline_merged(~excluded_mask, :);
fprintf('  ✓ 제외 대상 제거 후: %d명\n', height(baseline_merged));

% 고성과자/저성과자 분류
baseline_high_mask = false(height(baseline_merged), 1);
baseline_low_mask = false(height(baseline_merged), 1);

for i = 1:length(config.high_performers)
    baseline_high_mask = baseline_high_mask | strcmp(baseline_merged.(talent_col), config.high_performers{i});
end

for i = 1:length(config.low_performers)
    baseline_low_mask = baseline_low_mask | strcmp(baseline_merged.(talent_col), config.low_performers{i});
end

baseline_merged.PerformanceGroup = cell(height(baseline_merged), 1);
baseline_merged.PerformanceGroup(baseline_high_mask) = {'고성과자'};
baseline_merged.PerformanceGroup(baseline_low_mask) = {'저성과자'};

fprintf('  ✓ 고성과자: %d명\n', sum(baseline_high_mask));
fprintf('  ✓ 저성과자: %d명\n\n', sum(baseline_low_mask));

%% ========================================================================
%                    PART 3: 새로운 데이터 로딩
% =========================================================================

fprintf('【STEP 2】 2025년 하반기 신규 입사자 데이터 로딩\n');
fprintf('────────────────────────────────────────────\n');

% 파일 존재 확인
if ~exist(config.new_onboarding_file, 'file')
    error('새로운 온보딩 데이터를 찾을 수 없습니다: %s', config.new_onboarding_file);
end
if ~exist(config.new_comp_file, 'file')
    error('새로운 역량검사 데이터를 찾을 수 없습니다: %s', config.new_comp_file);
end

% 온보딩 데이터 로딩 (합격/불합격 정보)
new_onboarding = readtable(config.new_onboarding_file, 'VariableNamingRule', 'preserve');
fprintf('  ✓ 온보딩 데이터: %d명\n', height(new_onboarding));

% 역량검사 데이터 로딩
new_comp_upper = readtable(config.new_comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
fprintf('  ✓ 역량검사 데이터: %d명\n', height(new_comp_upper));

% 데이터 병합
new_merged = innerjoin(new_onboarding, new_comp_upper, 'Keys', 'ID');
fprintf('  ✓ 병합 후: %d명\n', height(new_merged));

% 합격/불합격 기준으로 그룹 분류
if ~ismember('합격 여부', new_merged.Properties.VariableNames)
    error('합격 여부 컬럼을 찾을 수 없습니다.');
end

new_high_mask = strcmp(new_merged.('합격 여부'), '합격');
new_low_mask = strcmp(new_merged.('합격 여부'), '불합격');

new_merged.PerformanceGroup = cell(height(new_merged), 1);
new_merged.PerformanceGroup(new_high_mask) = {'고성과자'};
new_merged.PerformanceGroup(new_low_mask) = {'저성과자'};

fprintf('  ✓ 합격 (고성과자): %d명\n', sum(new_high_mask));
fprintf('  ✓ 불합격 (저성과자): %d명\n\n', sum(new_low_mask));

%% ========================================================================
%                    PART 4: 역량 컬럼 추출
% =========================================================================

fprintf('【STEP 3】 역량 컬럼 추출\n');
fprintf('────────────────────────────────────────────\n');

% 역량 컬럼 식별 (숫자형이고 0-100 범위)
all_cols = baseline_comp_upper.Properties.VariableNames;
competency_cols = {};

for i = 1:length(all_cols)
    col_name = all_cols{i};

    % ID, 이름 등 제외
    if strcmpi(col_name, 'ID') || contains(col_name, '이름') || contains(col_name, '콘텐트') || ...
       contains(col_name, '학습') || contains(col_name, '직업') || contains(col_name, '역할')
        continue;
    end

    % 숫자형 컬럼만 선택
    col_data = baseline_comp_upper.(col_name);
    if isnumeric(col_data)
        % 0-100 범위 확인
        valid_data = col_data(~isnan(col_data));
        if ~isempty(valid_data) && min(valid_data) >= 0 && max(valid_data) <= 100
            competency_cols{end+1} = col_name;
        end
    end
end

fprintf('  ✓ 발견된 역량: %d개\n', length(competency_cols));
fprintf('  역량 목록: %s\n\n', strjoin(competency_cols, ', '));

% 기존 데이터와 새 데이터에 공통으로 있는 역량만 선택
common_competencies = {};
for i = 1:length(competency_cols)
    comp = competency_cols{i};
    if ismember(comp, baseline_merged.Properties.VariableNames) && ...
       ismember(comp, new_merged.Properties.VariableNames)
        common_competencies{end+1} = comp;
    end
end

if isempty(common_competencies)
    error('기존 데이터와 새 데이터 간에 공통 역량이 없습니다.');
end

fprintf('  ✓ 공통 역량: %d개\n', length(common_competencies));
fprintf('  공통 역량 목록: %s\n\n', strjoin(common_competencies, ', '));

competency_cols = common_competencies;

%% ========================================================================
%                    PART 5: 분포 비교 - 전체 참가자
% =========================================================================

fprintf('【STEP 4】 분포 비교 - 전체 참가자\n');
fprintf('────────────────────────────────────────────\n');

% 통계 검정 결과 저장
stats_results_all = table();

% 각 역량별 통계 검정
for i = 1:length(competency_cols)
    comp = competency_cols{i};

    % 데이터 추출
    baseline_data = baseline_merged.(comp);
    new_data = new_merged.(comp);

    % NaN 제거
    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    % 기술통계
    baseline_mean = mean(baseline_data);
    baseline_std = std(baseline_data);
    baseline_n = length(baseline_data);

    new_mean = mean(new_data);
    new_std = std(new_data);
    new_n = length(new_data);

    % KS-test (분포 동질성)
    [~, ks_p, ks_stat] = kstest2(baseline_data, new_data);

    % t-test (평균 차이)
    [~, t_p, ~, t_stats] = ttest2(baseline_data, new_data);
    t_stat = t_stats.tstat;

    % Levene's test (분산 동질성)
    % F-test 대신 사용 (더 robust)
    group_labels = [ones(baseline_n, 1); 2*ones(new_n, 1)];
    combined_data = [baseline_data; new_data];
    levene_p = vartestn(combined_data, group_labels, 'Display', 'off', 'TestType', 'LeveneAbsolute');

    % 효과 크기 (Cohen's d)
    pooled_std = sqrt(((baseline_n-1)*baseline_std^2 + (new_n-1)*new_std^2) / (baseline_n + new_n - 2));
    cohens_d = (baseline_mean - new_mean) / pooled_std;

    % 결과 저장
    stats_results_all = [stats_results_all; table({comp}, baseline_mean, baseline_std, baseline_n, ...
                                                   new_mean, new_std, new_n, ...
                                                   ks_stat, ks_p, t_stat, t_p, levene_p, cohens_d, ...
                                                   'VariableNames', {'Competency', 'Baseline_Mean', 'Baseline_SD', 'Baseline_N', ...
                                                                     'New_Mean', 'New_SD', 'New_N', ...
                                                                     'KS_Statistic', 'KS_pValue', 't_Statistic', 't_pValue', ...
                                                                     'Levene_pValue', 'Cohens_d'})];

    % 결과 출력
    fprintf('  %s:\n', comp);
    fprintf('    기존: M=%.2f, SD=%.2f (N=%d)\n', baseline_mean, baseline_std, baseline_n);
    fprintf('    신규: M=%.2f, SD=%.2f (N=%d)\n', new_mean, new_std, new_n);
    fprintf('    KS-test: p=%.4f %s\n', ks_p, iif(ks_p < 0.05, '(분포 다름 ***)', '(분포 유사)'));
    fprintf('    t-test:  p=%.4f %s\n', t_p, iif(t_p < 0.05, '(평균 다름 ***)', '(평균 유사)'));
    fprintf('    Cohen''s d: %.3f %s\n\n', cohens_d, interpret_effect_size(cohens_d));
end

%% ========================================================================
%                    PART 6: 분포 비교 - 고성과자
% =========================================================================

fprintf('【STEP 5】 분포 비교 - 고성과자 (선발대상)\n');
fprintf('────────────────────────────────────────────\n');

% 고성과자 데이터 필터링
baseline_high = baseline_merged(baseline_high_mask, :);
new_high = new_merged(new_high_mask, :);

% 통계 검정 결과 저장
stats_results_high = table();

% 각 역량별 통계 검정
for i = 1:length(competency_cols)
    comp = competency_cols{i};

    baseline_data = baseline_high.(comp);
    new_data = new_high.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    baseline_mean = mean(baseline_data);
    baseline_std = std(baseline_data);
    baseline_n = length(baseline_data);

    new_mean = mean(new_data);
    new_std = std(new_data);
    new_n = length(new_data);

    [~, ks_p, ks_stat] = kstest2(baseline_data, new_data);
    [~, t_p, ~, t_stats] = ttest2(baseline_data, new_data);
    t_stat = t_stats.tstat;

    group_labels = [ones(baseline_n, 1); 2*ones(new_n, 1)];
    combined_data = [baseline_data; new_data];
    levene_p = vartestn(combined_data, group_labels, 'Display', 'off', 'TestType', 'LeveneAbsolute');

    pooled_std = sqrt(((baseline_n-1)*baseline_std^2 + (new_n-1)*new_std^2) / (baseline_n + new_n - 2));
    cohens_d = (baseline_mean - new_mean) / pooled_std;

    stats_results_high = [stats_results_high; table({comp}, baseline_mean, baseline_std, baseline_n, ...
                                                     new_mean, new_std, new_n, ...
                                                     ks_stat, ks_p, t_stat, t_p, levene_p, cohens_d, ...
                                                     'VariableNames', {'Competency', 'Baseline_Mean', 'Baseline_SD', 'Baseline_N', ...
                                                                       'New_Mean', 'New_SD', 'New_N', ...
                                                                       'KS_Statistic', 'KS_pValue', 't_Statistic', 't_pValue', ...
                                                                       'Levene_pValue', 'Cohens_d'})];

    fprintf('  %s:\n', comp);
    fprintf('    기존: M=%.2f, SD=%.2f (N=%d)\n', baseline_mean, baseline_std, baseline_n);
    fprintf('    신규: M=%.2f, SD=%.2f (N=%d)\n', new_mean, new_std, new_n);
    fprintf('    KS-test: p=%.4f %s\n', ks_p, iif(ks_p < 0.05, '(분포 다름 ***)', '(분포 유사)'));
    fprintf('    t-test:  p=%.4f %s\n', t_p, iif(t_p < 0.05, '(평균 다름 ***)', '(평균 유사)'));
    fprintf('    Cohen''s d: %.3f %s\n\n', cohens_d, interpret_effect_size(cohens_d));
end

%% ========================================================================
%                    PART 7: 분포 비교 - 저성과자
% =========================================================================

fprintf('【STEP 6】 분포 비교 - 저성과자 (비선발대상)\n');
fprintf('────────────────────────────────────────────\n');

% 저성과자 데이터 필터링
baseline_low = baseline_merged(baseline_low_mask, :);
new_low = new_merged(new_low_mask, :);

% 통계 검정 결과 저장
stats_results_low = table();

% 각 역량별 통계 검정
for i = 1:length(competency_cols)
    comp = competency_cols{i};

    baseline_data = baseline_low.(comp);
    new_data = new_low.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    baseline_mean = mean(baseline_data);
    baseline_std = std(baseline_data);
    baseline_n = length(baseline_data);

    new_mean = mean(new_data);
    new_std = std(new_data);
    new_n = length(new_data);

    [~, ks_p, ks_stat] = kstest2(baseline_data, new_data);
    [~, t_p, ~, t_stats] = ttest2(baseline_data, new_data);
    t_stat = t_stats.tstat;

    group_labels = [ones(baseline_n, 1); 2*ones(new_n, 1)];
    combined_data = [baseline_data; new_data];
    levene_p = vartestn(combined_data, group_labels, 'Display', 'off', 'TestType', 'LeveneAbsolute');

    pooled_std = sqrt(((baseline_n-1)*baseline_std^2 + (new_n-1)*new_std^2) / (baseline_n + new_n - 2));
    cohens_d = (baseline_mean - new_mean) / pooled_std;

    stats_results_low = [stats_results_low; table({comp}, baseline_mean, baseline_std, baseline_n, ...
                                                   new_mean, new_std, new_n, ...
                                                   ks_stat, ks_p, t_stat, t_p, levene_p, cohens_d, ...
                                                   'VariableNames', {'Competency', 'Baseline_Mean', 'Baseline_SD', 'Baseline_N', ...
                                                                     'New_Mean', 'New_SD', 'New_N', ...
                                                                     'KS_Statistic', 'KS_pValue', 't_Statistic', 't_pValue', ...
                                                                     'Levene_pValue', 'Cohens_d'})];

    fprintf('  %s:\n', comp);
    fprintf('    기존: M=%.2f, SD=%.2f (N=%d)\n', baseline_mean, baseline_std, baseline_n);
    fprintf('    신규: M=%.2f, SD=%.2f (N=%d)\n', new_mean, new_std, new_n);
    fprintf('    KS-test: p=%.4f %s\n', ks_p, iif(ks_p < 0.05, '(분포 다름 ***)', '(분포 유사)'));
    fprintf('    t-test:  p=%.4f %s\n', t_p, iif(t_p < 0.05, '(평균 다름 ***)', '(평균 유사)'));
    fprintf('    Cohen''s d: %.3f %s\n\n', cohens_d, interpret_effect_size(cohens_d));
end

%% ========================================================================
%                    PART 8: 통계 결과 저장
% =========================================================================

fprintf('【STEP 7】 통계 검정 결과 저장\n');
fprintf('────────────────────────────────────────────\n');

% Excel 파일로 저장
excel_file = fullfile(config.output_dir, sprintf('distribution_comparison_stats_%s.xlsx', config.timestamp));

writetable(stats_results_all, excel_file, 'Sheet', '전체참가자');
writetable(stats_results_high, excel_file, 'Sheet', '고성과자');
writetable(stats_results_low, excel_file, 'Sheet', '저성과자');

fprintf('  ✓ 통계 결과 저장: %s\n\n', excel_file);

%% ========================================================================
%                    PART 9: 시각화 - 박스플롯
% =========================================================================

fprintf('【STEP 8】 시각화 - 박스플롯\n');
fprintf('────────────────────────────────────────────\n');

% 역량 개수에 따라 subplot 구성
n_comps = length(competency_cols);
n_rows = ceil(n_comps / 4);
n_cols = min(4, n_comps);

% 전체 참가자 박스플롯
fig1 = figure('Position', [100, 100, 1600, 400*n_rows], 'Color', 'white');
for i = 1:n_comps
    comp = competency_cols{i};

    subplot(n_rows, n_cols, i);

    baseline_data = baseline_merged.(comp);
    new_data = new_merged.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    % 박스플롯 데이터 준비
    all_data = [baseline_data; new_data];
    group_labels = [repmat({'기존'}, length(baseline_data), 1); repmat({'신규'}, length(new_data), 1)];

    boxplot(all_data, group_labels, 'Colors', [0.3 0.5 0.8; 0.8 0.4 0.3], 'Symbol', 'o');
    ylabel('점수');
    title(comp, 'FontSize', 11, 'FontWeight', 'bold');
    grid on;
    ylim([0 100]);
end

sgtitle('역량별 분포 비교 - 전체 참가자', 'FontSize', 14, 'FontWeight', 'bold');
boxplot_all_file = fullfile(config.output_dir, sprintf('boxplot_all_%s.png', config.timestamp));
saveas(fig1, boxplot_all_file);
fprintf('  ✓ 전체 참가자 박스플롯 저장: %s\n', boxplot_all_file);

% 고성과자 박스플롯
fig2 = figure('Position', [100, 100, 1600, 400*n_rows], 'Color', 'white');
for i = 1:n_comps
    comp = competency_cols{i};

    subplot(n_rows, n_cols, i);

    baseline_data = baseline_high.(comp);
    new_data = new_high.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    all_data = [baseline_data; new_data];
    group_labels = [repmat({'기존'}, length(baseline_data), 1); repmat({'신규'}, length(new_data), 1)];

    boxplot(all_data, group_labels, 'Colors', [0.3 0.5 0.8; 0.8 0.4 0.3], 'Symbol', 'o');
    ylabel('점수');
    title(comp, 'FontSize', 11, 'FontWeight', 'bold');
    grid on;
    ylim([0 100]);
end

sgtitle('역량별 분포 비교 - 고성과자 (선발대상)', 'FontSize', 14, 'FontWeight', 'bold');
boxplot_high_file = fullfile(config.output_dir, sprintf('boxplot_high_%s.png', config.timestamp));
saveas(fig2, boxplot_high_file);
fprintf('  ✓ 고성과자 박스플롯 저장: %s\n', boxplot_high_file);

% 저성과자 박스플롯
fig3 = figure('Position', [100, 100, 1600, 400*n_rows], 'Color', 'white');
for i = 1:n_comps
    comp = competency_cols{i};

    subplot(n_rows, n_cols, i);

    baseline_data = baseline_low.(comp);
    new_data = new_low.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    all_data = [baseline_data; new_data];
    group_labels = [repmat({'기존'}, length(baseline_data), 1); repmat({'신규'}, length(new_data), 1)];

    boxplot(all_data, group_labels, 'Colors', [0.3 0.5 0.8; 0.8 0.4 0.3], 'Symbol', 'o');
    ylabel('점수');
    title(comp, 'FontSize', 11, 'FontWeight', 'bold');
    grid on;
    ylim([0 100]);
end

sgtitle('역량별 분포 비교 - 저성과자 (비선발대상)', 'FontSize', 14, 'FontWeight', 'bold');
boxplot_low_file = fullfile(config.output_dir, sprintf('boxplot_low_%s.png', config.timestamp));
saveas(fig3, boxplot_low_file);
fprintf('  ✓ 저성과자 박스플롯 저장: %s\n\n', boxplot_low_file);

%% ========================================================================
%                    PART 10: 시각화 - 히스토그램
% =========================================================================

fprintf('【STEP 9】 시각화 - 히스토그램\n');
fprintf('────────────────────────────────────────────\n');

% 전체 참가자 히스토그램
fig4 = figure('Position', [100, 100, 1600, 400*n_rows], 'Color', 'white');
for i = 1:n_comps
    comp = competency_cols{i};

    subplot(n_rows, n_cols, i);

    baseline_data = baseline_merged.(comp);
    new_data = new_merged.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    % 히스토그램
    bin_edges = linspace(0, 100, 20);
    histogram(baseline_data, bin_edges, 'FaceColor', [0.3 0.5 0.8], 'FaceAlpha', 0.5, 'EdgeColor', 'none');
    hold on;
    histogram(new_data, bin_edges, 'FaceColor', [0.8 0.4 0.3], 'FaceAlpha', 0.5, 'EdgeColor', 'none');

    xlabel('점수');
    ylabel('빈도');
    title(comp, 'FontSize', 11, 'FontWeight', 'bold');
    legend({'기존', '신규'}, 'Location', 'best');
    grid on;
    xlim([0 100]);
end

sgtitle('역량별 분포 히스토그램 - 전체 참가자', 'FontSize', 14, 'FontWeight', 'bold');
hist_all_file = fullfile(config.output_dir, sprintf('histogram_all_%s.png', config.timestamp));
saveas(fig4, hist_all_file);
fprintf('  ✓ 전체 참가자 히스토그램 저장: %s\n', hist_all_file);

% 고성과자 히스토그램
fig5 = figure('Position', [100, 100, 1600, 400*n_rows], 'Color', 'white');
for i = 1:n_comps
    comp = competency_cols{i};

    subplot(n_rows, n_cols, i);

    baseline_data = baseline_high.(comp);
    new_data = new_high.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    bin_edges = linspace(0, 100, 20);
    histogram(baseline_data, bin_edges, 'FaceColor', [0.3 0.5 0.8], 'FaceAlpha', 0.5, 'EdgeColor', 'none');
    hold on;
    histogram(new_data, bin_edges, 'FaceColor', [0.8 0.4 0.3], 'FaceAlpha', 0.5, 'EdgeColor', 'none');

    xlabel('점수');
    ylabel('빈도');
    title(comp, 'FontSize', 11, 'FontWeight', 'bold');
    legend({'기존', '신규'}, 'Location', 'best');
    grid on;
    xlim([0 100]);
end

sgtitle('역량별 분포 히스토그램 - 고성과자 (선발대상)', 'FontSize', 14, 'FontWeight', 'bold');
hist_high_file = fullfile(config.output_dir, sprintf('histogram_high_%s.png', config.timestamp));
saveas(fig5, hist_high_file);
fprintf('  ✓ 고성과자 히스토그램 저장: %s\n', hist_high_file);

% 저성과자 히스토그램
fig6 = figure('Position', [100, 100, 1600, 400*n_rows], 'Color', 'white');
for i = 1:n_comps
    comp = competency_cols{i};

    subplot(n_rows, n_cols, i);

    baseline_data = baseline_low.(comp);
    new_data = new_low.(comp);

    baseline_data = baseline_data(~isnan(baseline_data));
    new_data = new_data(~isnan(new_data));

    bin_edges = linspace(0, 100, 20);
    histogram(baseline_data, bin_edges, 'FaceColor', [0.3 0.5 0.8], 'FaceAlpha', 0.5, 'EdgeColor', 'none');
    hold on;
    histogram(new_data, bin_edges, 'FaceColor', [0.8 0.4 0.3], 'FaceAlpha', 0.5, 'EdgeColor', 'none');

    xlabel('점수');
    ylabel('빈도');
    title(comp, 'FontSize', 11, 'FontWeight', 'bold');
    legend({'기존', '신규'}, 'Location', 'best');
    grid on;
    xlim([0 100]);
end

sgtitle('역량별 분포 히스토그램 - 저성과자 (비선발대상)', 'FontSize', 14, 'FontWeight', 'bold');
hist_low_file = fullfile(config.output_dir, sprintf('histogram_low_%s.png', config.timestamp));
saveas(fig6, hist_low_file);
fprintf('  ✓ 저성과자 히스토그램 저장: %s\n\n', hist_low_file);

%% ========================================================================
%                    PART 11: 시각화 - 레이더 차트
% =========================================================================

fprintf('【STEP 10】 시각화 - 레이더 차트\n');
fprintf('────────────────────────────────────────────\n');

% 전체 참가자 레이더 차트
baseline_means_all = zeros(1, n_comps);
new_means_all = zeros(1, n_comps);

for i = 1:n_comps
    comp = competency_cols{i};
    baseline_data = baseline_merged.(comp);
    new_data = new_merged.(comp);

    baseline_means_all(i) = mean(baseline_data(~isnan(baseline_data)));
    new_means_all(i) = mean(new_data(~isnan(new_data)));
end

fig7 = figure('Position', [100, 100, 800, 800], 'Color', 'white');
plot_radar_chart(baseline_means_all, new_means_all, competency_cols, '전체 참가자');
radar_all_file = fullfile(config.output_dir, sprintf('radar_all_%s.png', config.timestamp));
saveas(fig7, radar_all_file);
fprintf('  ✓ 전체 참가자 레이더 차트 저장: %s\n', radar_all_file);

% 고성과자 레이더 차트
baseline_means_high = zeros(1, n_comps);
new_means_high = zeros(1, n_comps);

for i = 1:n_comps
    comp = competency_cols{i};
    baseline_data = baseline_high.(comp);
    new_data = new_high.(comp);

    baseline_means_high(i) = mean(baseline_data(~isnan(baseline_data)));
    new_means_high(i) = mean(new_data(~isnan(new_data)));
end

fig8 = figure('Position', [100, 100, 800, 800], 'Color', 'white');
plot_radar_chart(baseline_means_high, new_means_high, competency_cols, '고성과자 (선발대상)');
radar_high_file = fullfile(config.output_dir, sprintf('radar_high_%s.png', config.timestamp));
saveas(fig8, radar_high_file);
fprintf('  ✓ 고성과자 레이더 차트 저장: %s\n', radar_high_file);

% 저성과자 레이더 차트
baseline_means_low = zeros(1, n_comps);
new_means_low = zeros(1, n_comps);

for i = 1:n_comps
    comp = competency_cols{i};
    baseline_data = baseline_low.(comp);
    new_data = new_low.(comp);

    baseline_means_low(i) = mean(baseline_data(~isnan(baseline_data)));
    new_means_low(i) = mean(new_data(~isnan(new_data)));
end

fig9 = figure('Position', [100, 100, 800, 800], 'Color', 'white');
plot_radar_chart(baseline_means_low, new_means_low, competency_cols, '저성과자 (비선발대상)');
radar_low_file = fullfile(config.output_dir, sprintf('radar_low_%s.png', config.timestamp));
saveas(fig9, radar_low_file);
fprintf('  ✓ 저성과자 레이더 차트 저장: %s\n\n', radar_low_file);

%% ========================================================================
%                    PART 12: 종합 요약 리포트
% =========================================================================

fprintf('【STEP 11】 종합 요약 리포트 생성\n');
fprintf('────────────────────────────────────────────\n');

% 분포 차이 심각도 분류
summary_report = struct();

% 전체 참가자
summary_report.all.total_competencies = n_comps;
summary_report.all.different_distribution = sum(stats_results_all.KS_pValue < 0.05);
summary_report.all.different_mean = sum(stats_results_all.t_pValue < 0.05);
summary_report.all.large_effect_size = sum(abs(stats_results_all.Cohens_d) > 0.5);

% 고성과자
summary_report.high.total_competencies = n_comps;
summary_report.high.different_distribution = sum(stats_results_high.KS_pValue < 0.05);
summary_report.high.different_mean = sum(stats_results_high.t_pValue < 0.05);
summary_report.high.large_effect_size = sum(abs(stats_results_high.Cohens_d) > 0.5);

% 저성과자
summary_report.low.total_competencies = n_comps;
summary_report.low.different_distribution = sum(stats_results_low.KS_pValue < 0.05);
summary_report.low.different_mean = sum(stats_results_low.t_pValue < 0.05);
summary_report.low.large_effect_size = sum(abs(stats_results_low.Cohens_d) > 0.5);

% 리포트 출력
fprintf('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
fprintf('               분포 비교 종합 요약\n');
fprintf('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n');

fprintf('【전체 참가자】\n');
fprintf('  • 총 역량 수: %d개\n', summary_report.all.total_competencies);
fprintf('  • 분포 차이 (KS-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.all.different_distribution, ...
        100 * summary_report.all.different_distribution / n_comps);
fprintf('  • 평균 차이 (t-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.all.different_mean, ...
        100 * summary_report.all.different_mean / n_comps);
fprintf('  • 큰 효과 크기 (|d|>0.5): %d개 (%.1f%%)\n\n', ...
        summary_report.all.large_effect_size, ...
        100 * summary_report.all.large_effect_size / n_comps);

fprintf('【고성과자 (선발대상)】\n');
fprintf('  • 총 역량 수: %d개\n', summary_report.high.total_competencies);
fprintf('  • 분포 차이 (KS-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.high.different_distribution, ...
        100 * summary_report.high.different_distribution / n_comps);
fprintf('  • 평균 차이 (t-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.high.different_mean, ...
        100 * summary_report.high.different_mean / n_comps);
fprintf('  • 큰 효과 크기 (|d|>0.5): %d개 (%.1f%%)\n\n', ...
        summary_report.high.large_effect_size, ...
        100 * summary_report.high.large_effect_size / n_comps);

fprintf('【저성과자 (비선발대상)】\n');
fprintf('  • 총 역량 수: %d개\n', summary_report.low.total_competencies);
fprintf('  • 분포 차이 (KS-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.low.different_distribution, ...
        100 * summary_report.low.different_distribution / n_comps);
fprintf('  • 평균 차이 (t-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.low.different_mean, ...
        100 * summary_report.low.different_mean / n_comps);
fprintf('  • 큰 효과 크기 (|d|>0.5): %d개 (%.1f%%)\n\n', ...
        summary_report.low.large_effect_size, ...
        100 * summary_report.low.large_effect_size / n_comps);

fprintf('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n');

% 권장사항 출력
fprintf('【권장사항】\n');
if summary_report.all.different_distribution / n_comps > 0.3
    fprintf('  ⚠ 주의: 전체 참가자의 30%% 이상 역량에서 분포 차이 발견\n');
    fprintf('     → 모델 재학습 또는 보정 권장\n\n');
elseif summary_report.all.different_mean / n_comps > 0.3
    fprintf('  ⚠ 주의: 전체 참가자의 30%% 이상 역량에서 평균 차이 발견\n');
    fprintf('     → 점수 정규화 또는 캘리브레이션 권장\n\n');
else
    fprintf('  ✓ 양호: 대부분의 역량에서 분포가 유사함\n');
    fprintf('     → 기존 모델 적용 가능\n\n');
end

% 리포트 저장
report_file = fullfile(config.output_dir, sprintf('summary_report_%s.txt', config.timestamp));
fid = fopen(report_file, 'w', 'n', 'UTF-8');
fprintf(fid, '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
fprintf(fid, '               분포 비교 종합 요약\n');
fprintf(fid, '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n');
fprintf(fid, '생성 일시: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf(fid, '【전체 참가자】\n');
fprintf(fid, '  • 총 역량 수: %d개\n', summary_report.all.total_competencies);
fprintf(fid, '  • 분포 차이 (KS-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.all.different_distribution, ...
        100 * summary_report.all.different_distribution / n_comps);
fprintf(fid, '  • 평균 차이 (t-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.all.different_mean, ...
        100 * summary_report.all.different_mean / n_comps);
fprintf(fid, '  • 큰 효과 크기 (|d|>0.5): %d개 (%.1f%%)\n\n', ...
        summary_report.all.large_effect_size, ...
        100 * summary_report.all.large_effect_size / n_comps);
fprintf(fid, '【고성과자 (선발대상)】\n');
fprintf(fid, '  • 총 역량 수: %d개\n', summary_report.high.total_competencies);
fprintf(fid, '  • 분포 차이 (KS-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.high.different_distribution, ...
        100 * summary_report.high.different_distribution / n_comps);
fprintf(fid, '  • 평균 차이 (t-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.high.different_mean, ...
        100 * summary_report.high.different_mean / n_comps);
fprintf(fid, '  • 큰 효과 크기 (|d|>0.5): %d개 (%.1f%%)\n\n', ...
        summary_report.high.large_effect_size, ...
        100 * summary_report.high.large_effect_size / n_comps);
fprintf(fid, '【저성과자 (비선발대상)】\n');
fprintf(fid, '  • 총 역량 수: %d개\n', summary_report.low.total_competencies);
fprintf(fid, '  • 분포 차이 (KS-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.low.different_distribution, ...
        100 * summary_report.low.different_distribution / n_comps);
fprintf(fid, '  • 평균 차이 (t-test p<0.05): %d개 (%.1f%%)\n', ...
        summary_report.low.different_mean, ...
        100 * summary_report.low.different_mean / n_comps);
fprintf(fid, '  • 큰 효과 크기 (|d|>0.5): %d개 (%.1f%%)\n\n', ...
        summary_report.low.large_effect_size, ...
        100 * summary_report.low.large_effect_size / n_comps);
fclose(fid);

fprintf('  ✓ 요약 리포트 저장: %s\n\n', report_file);

fprintf('========================================================================\n');
fprintf('분포 비교 분석 완료!\n');
fprintf('========================================================================\n');

%% ========================================================================
%                          보조 함수
% =========================================================================

function result = iif(condition, true_val, false_val)
    % 삼항 연산자 구현
    if condition
        result = true_val;
    else
        result = false_val;
    end
end

function interpretation = interpret_effect_size(d)
    % Cohen's d 효과 크기 해석
    abs_d = abs(d);
    if abs_d < 0.2
        interpretation = '(매우 작음)';
    elseif abs_d < 0.5
        interpretation = '(작음)';
    elseif abs_d < 0.8
        interpretation = '(중간)';
    else
        interpretation = '(큼 ***)';
    end
end

function plot_radar_chart(baseline_means, new_means, labels, title_text)
    % 레이더 차트 그리기
    n = length(labels);
    theta = linspace(0, 2*pi, n+1);

    % 데이터 순환 (닫힌 도형 만들기)
    baseline_data = [baseline_means, baseline_means(1)];
    new_data = [new_means, new_means(1)];

    % 극좌표 플롯
    polarplot(theta, baseline_data, 'o-', 'LineWidth', 2.5, 'Color', [0.3 0.5 0.8], ...
              'MarkerSize', 8, 'MarkerFaceColor', [0.3 0.5 0.8]);
    hold on;
    polarplot(theta, new_data, 's-', 'LineWidth', 2.5, 'Color', [0.8 0.4 0.3], ...
              'MarkerSize', 8, 'MarkerFaceColor', [0.8 0.4 0.3]);

    % 축 설정
    ax = gca;
    ax.ThetaTick = rad2deg(theta(1:end-1));
    ax.ThetaTickLabel = labels;
    ax.RLim = [0 100];
    ax.RGrid = 'on';
    ax.ThetaGrid = 'on';

    % 범례 및 제목
    legend({'기존 데이터', '신규 데이터'}, 'Location', 'best', 'FontSize', 11);
    title(sprintf('역량 프로필 비교 - %s', title_text), 'FontSize', 14, 'FontWeight', 'bold');

    hold off;
end
