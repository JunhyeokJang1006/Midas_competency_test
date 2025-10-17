%% 가중치 적용 vs 원본 역량검사 점수 비교 분석
%
% 목적:
%   - 가중치 적용 역량검사 점수 기준 상위 그룹 (A그룹)
%   - 원본 역량검사 종합점수 기준 상위 그룹 (B그룹)
%   - 각 그룹의 역량진단 성과점수와의 상관관계 비교
%   - 백분위수별 분석: 10%, 25%, 33%, 50%, 100%
%
% 입력:
%   - 가중치 적용: D:\project\HR데이터\결과\자가불소_revised_talent\역량검사_가중치적용점수_talent_*.xlsx
%   - 원본 종합점수: D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx
%   - 역량진단 성과: D:\project\HR데이터\matlab\문항기반_revised\*_workspace_*.mat
%
% 출력:
%   - 결과 엑셀: D:\project\HR데이터\결과\가중치vs역진\
%   - 시각화 그래프: PNG 파일
%
% 작성일: 2025-10-17
% =======================================================================

clear; clc; close all;
rng(42, 'twister');  % 재현성 보장

%% 전역 폰트 설정
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 11);
set(0, 'DefaultTextFontSize', 11);

fprintf('=====================================================\n');
fprintf('  가중치 적용 vs 원본 역량검사 점수 비교 분석\n');
fprintf('=====================================================\n\n');

%% 1) 설정
fprintf('[STEP 1] 설정\n');
fprintf('-----------------------------------------------------\n');

config = struct();
config.weighted_score_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent';
config.original_score_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';
config.performance_data_dir = 'D:\project\HR데이터\matlab\문항기반_revised';
config.output_dir = 'D:\project\HR데이터\결과\가중치vs역진';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');

% 출력 디렉토리 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
    fprintf('  ✓ 출력 디렉토리 생성: %s\n', config.output_dir);
else
    fprintf('  ✓ 출력 디렉토리 확인: %s\n', config.output_dir);
end

%% 2) 가중치 적용 역량검사 점수 로드
fprintf('\n[STEP 2] 가중치 적용 역량검사 점수 로드\n');
fprintf('-----------------------------------------------------\n');

% 최신 가중치 적용 점수 파일 찾기
weighted_files = dir(fullfile(config.weighted_score_dir, '역량검사_가중치적용점수_talent*.xlsx'));
if isempty(weighted_files)
    error('가중치 적용 점수 파일을 찾을 수 없습니다: %s', config.weighted_score_dir);
end

[~, idx] = max([weighted_files.datenum]);
weighted_file = fullfile(weighted_files(idx).folder, weighted_files(idx).name);
fprintf('  ✓ 가중치 점수 파일: %s\n', weighted_files(idx).name);

% 데이터 로드 (한글 컬럼명 보존)
weighted_data = readtable(weighted_file, 'Sheet', '역량검사_종합점수', ...
                         'VariableNamingRule', 'preserve');
fprintf('  ✓ 로드 완료: %d행 x %d열\n', height(weighted_data), width(weighted_data));

% 필요한 컬럼 추출 (컬럼 위치로 접근)
% 1번 컬럼: ID, 3번 컬럼: 총점
weighted_scores = table();
weighted_scores.ID = weighted_data{:, 1};  % ID
weighted_scores.('가중치적용점수') = weighted_data{:, 3};  % 총점

fprintf('  ✓ 가중치 적용 점수: %d명 (평균 %.2f ± %.2f)\n', ...
    height(weighted_scores), ...
    mean(weighted_scores.('가중치적용점수'), 'omitnan'), ...
    std(weighted_scores.('가중치적용점수'), 'omitnan'));

%% 3) 원본 역량검사 종합점수 로드
fprintf('\n[STEP 3] 원본 역량검사 종합점수 로드\n');
fprintf('-----------------------------------------------------\n');

fprintf('  ✓ 원본 점수 파일: 23-25년 역량검사_개발자추가_filtered.xlsx\n');

% 데이터 로드 (한글 컬럼명 보존)
original_data = readtable(config.original_score_file, 'Sheet', '역량검사_종합점수', ...
                         'VariableNamingRule', 'preserve');
fprintf('  ✓ 로드 완료: %d행 x %d열\n', height(original_data), width(original_data));

% 필요한 컬럼 추출
% ID와 총점 컬럼 찾기
col_names = original_data.Properties.VariableNames;
id_col_idx = find(contains(col_names, 'ID', 'IgnoreCase', true), 1);
total_col_idx = find(contains(col_names, '총점', 'IgnoreCase', true), 1);

if isempty(id_col_idx) || isempty(total_col_idx)
    error('ID 또는 총점 컬럼을 찾을 수 없습니다.');
end

original_scores = table();
original_scores.ID = original_data{:, id_col_idx};
original_scores.('원본종합점수') = original_data{:, total_col_idx};

fprintf('  ✓ 원본 종합점수: %d명 (평균 %.2f ± %.2f)\n', ...
    height(original_scores), ...
    mean(original_scores.('원본종합점수'), 'omitnan'), ...
    std(original_scores.('원본종합점수'), 'omitnan'));

%% 4) 역량진단 성과점수 로드
fprintf('\n[STEP 4] 역량진단 성과점수 로드\n');
fprintf('-----------------------------------------------------\n');

% MAT 파일에서 로드
mat_files = dir(fullfile(config.performance_data_dir, '*_workspace_*.mat'));
if isempty(mat_files)
    error('역량진단 MAT 파일을 찾을 수 없습니다: %s', config.performance_data_dir);
end

[~, idx] = max([mat_files.datenum]);
mat_file = fullfile(mat_files(idx).folder, mat_files(idx).name);
fprintf('  ✓ MAT 파일: %s\n', mat_files(idx).name);

loaded_data = load(mat_file);
if ~isfield(loaded_data, 'integratedPerformanceData')
    error('MAT 파일에 integratedPerformanceData가 없습니다.');
end

performance_data = loaded_data.integratedPerformanceData;
fprintf('  ✓ 역량진단 데이터 로드: %d명\n', height(performance_data));

% 필요한 컬럼 추출
performance_scores = table();
performance_scores.ID = performance_data.ID;
performance_scores.('역량진단점수') = performance_data.PerformanceScore;

fprintf('  ✓ 역량진단 점수: %d명 (평균 %.2f ± %.2f)\n', ...
    height(performance_scores), ...
    mean(performance_scores.('역량진단점수'), 'omitnan'), ...
    std(performance_scores.('역량진단점수'), 'omitnan'));

%% 5) ID 매칭 - 3개 데이터 통합
fprintf('\n[STEP 5] 데이터 매칭 (가중치 + 원본 + 역량진단)\n');
fprintf('-----------------------------------------------------\n');

% ID 타입 통일 (cell → double)
if iscell(weighted_scores.ID)
    weighted_scores.ID = cellfun(@(x) str2double(x), weighted_scores.ID);
end
if iscell(original_scores.ID)
    original_scores.ID = cellfun(@(x) str2double(x), original_scores.ID);
end
if iscell(performance_scores.ID)
    performance_scores.ID = cellfun(@(x) str2double(x), performance_scores.ID);
end

% 3-way join
merged_data = innerjoin(weighted_scores, original_scores, 'Keys', 'ID');
merged_data = innerjoin(merged_data, performance_scores, 'Keys', 'ID');

% 결측치 제거
valid_idx = ~isnan(merged_data.('가중치적용점수')) & ...
            ~isnan(merged_data.('원본종합점수')) & ...
            ~isnan(merged_data.('역량진단점수'));
merged_data = merged_data(valid_idx, :);

fprintf('  ✓ 최종 매칭: %d명\n', height(merged_data));
fprintf('    - 가중치 적용 점수: 평균 %.2f ± %.2f\n', ...
    mean(merged_data.('가중치적용점수')), std(merged_data.('가중치적용점수')));
fprintf('    - 원본 종합점수: 평균 %.2f ± %.2f\n', ...
    mean(merged_data.('원본종합점수')), std(merged_data.('원본종합점수')));
fprintf('    - 역량진단 점수: 평균 %.2f ± %.2f\n', ...
    mean(merged_data.('역량진단점수')), std(merged_data.('역량진단점수')));

%% 6) 백분위수별 그룹 선별 및 분석
fprintf('\n[STEP 6] 백분위수별 그룹 선별 및 분석\n');
fprintf('-----------------------------------------------------\n');

% 백분위수 리스트 정의
percentiles = [10, 25, 33, 50, 100];
n_percentiles = length(percentiles);

% 결과 저장을 위한 구조체 배열
stats_A_all = cell(n_percentiles, 1);
stats_B_all = cell(n_percentiles, 1);
groups_A = cell(n_percentiles, 1);
groups_B = cell(n_percentiles, 1);
models_A = cell(n_percentiles, 1);
models_B = cell(n_percentiles, 1);

n_total = height(merged_data);

% 백분위수별 분석 루프
for i = 1:n_percentiles
    pct = percentiles(i);

    fprintf('\n  [백분위: 상위 %d%%]\n', pct);
    fprintf('  -------------------------------------------------\n');

    % 샘플 수 계산
    if pct == 100
        n_samples = n_total;  % 전체 데이터
    else
        n_samples = ceil(n_total * pct / 100);
    end

    %% 그룹 A: 가중치 적용 점수 기준 상위 X%
    [~, idx_weighted_sorted] = sort(merged_data.('가중치적용점수'), 'descend');
    group_A_idx = idx_weighted_sorted(1:n_samples);
    group_A = merged_data(group_A_idx, :);
    groups_A{i} = group_A;

    fprintf('    [그룹 A: 가중치 적용 점수 기준 상위 %d%%]\n', pct);
    fprintf('      • 샘플 수: %d명\n', height(group_A));
    fprintf('      • 가중치 점수 범위: %.2f ~ %.2f\n', ...
        min(group_A.('가중치적용점수')), max(group_A.('가중치적용점수')));

    % 그룹 A 통계
    stats_A = struct();
    stats_A.percentile = pct;
    stats_A.n = height(group_A);
    stats_A.weighted_mean = mean(group_A.('가중치적용점수'));
    stats_A.weighted_std = std(group_A.('가중치적용점수'));
    stats_A.performance_mean = mean(group_A.('역량진단점수'));
    stats_A.performance_std = std(group_A.('역량진단점수'));

    % 상관분석: 가중치 점수 vs 역량진단 점수
    [r_pearson_A, p_pearson_A] = corr(group_A.('가중치적용점수'), group_A.('역량진단점수'), ...
                                      'Type', 'Pearson');
    stats_A.r_pearson = r_pearson_A;
    stats_A.p_pearson = p_pearson_A;

    [r_spearman_A, p_spearman_A] = corr(group_A.('가중치적용점수'), group_A.('역량진단점수'), ...
                                        'Type', 'Spearman');
    stats_A.r_spearman = r_spearman_A;
    stats_A.p_spearman = p_spearman_A;

    % 회귀분석
    mdl_A = fitlm(group_A.('역량진단점수'), group_A.('가중치적용점수'));
    stats_A.rsquared = mdl_A.Rsquared.Ordinary;
    stats_A.rmse = mdl_A.RMSE;
    stats_A.coef_intercept = mdl_A.Coefficients.Estimate(1);
    stats_A.coef_slope = mdl_A.Coefficients.Estimate(2);

    stats_A_all{i} = stats_A;
    models_A{i} = mdl_A;

    fprintf('      • Pearson r = %.4f (p = %.4e)\n', r_pearson_A, p_pearson_A);
    fprintf('      • R² = %.4f, RMSE = %.4f\n', stats_A.rsquared, stats_A.rmse);

    %% 그룹 B: 원본 종합점수 기준 상위 X%
    [~, idx_original_sorted] = sort(merged_data.('원본종합점수'), 'descend');
    group_B_idx = idx_original_sorted(1:n_samples);
    group_B = merged_data(group_B_idx, :);
    groups_B{i} = group_B;

    fprintf('    [그룹 B: 원본 종합점수 기준 상위 %d%%]\n', pct);
    fprintf('      • 샘플 수: %d명\n', height(group_B));
    fprintf('      • 원본 점수 범위: %.2f ~ %.2f\n', ...
        min(group_B.('원본종합점수')), max(group_B.('원본종합점수')));

    % 그룹 B 통계
    stats_B = struct();
    stats_B.percentile = pct;
    stats_B.n = height(group_B);
    stats_B.original_mean = mean(group_B.('원본종합점수'));
    stats_B.original_std = std(group_B.('원본종합점수'));
    stats_B.performance_mean = mean(group_B.('역량진단점수'));
    stats_B.performance_std = std(group_B.('역량진단점수'));

    % 상관분석: 원본 점수 vs 역량진단 점수
    [r_pearson_B, p_pearson_B] = corr(group_B.('원본종합점수'), group_B.('역량진단점수'), ...
                                      'Type', 'Pearson');
    stats_B.r_pearson = r_pearson_B;
    stats_B.p_pearson = p_pearson_B;

    [r_spearman_B, p_spearman_B] = corr(group_B.('원본종합점수'), group_B.('역량진단점수'), ...
                                        'Type', 'Spearman');
    stats_B.r_spearman = r_spearman_B;
    stats_B.p_spearman = p_spearman_B;

    % 회귀분석
    mdl_B = fitlm(group_B.('역량진단점수'), group_B.('원본종합점수'));
    stats_B.rsquared = mdl_B.Rsquared.Ordinary;
    stats_B.rmse = mdl_B.RMSE;
    stats_B.coef_intercept = mdl_B.Coefficients.Estimate(1);
    stats_B.coef_slope = mdl_B.Coefficients.Estimate(2);

    stats_B_all{i} = stats_B;
    models_B{i} = mdl_B;

    fprintf('      • Pearson r = %.4f (p = %.4e)\n', r_pearson_B, p_pearson_B);
    fprintf('      • R² = %.4f, RMSE = %.4f\n', stats_B.rsquared, stats_B.rmse);
end

fprintf('\n  ✓ 백분위수별 분석 완료 (%d개 그룹)\n', n_percentiles);

%% 7) 시각화
fprintf('\n[STEP 7] 시각화\n');
fprintf('-----------------------------------------------------\n');

% 그림 1: 백분위수별 산점도 (그룹 A - 가중치 적용 점수 기준)
fig1 = figure('Position', [100, 100, 1400, 900]);
colors_A = [0.8, 0.2, 0.2; 0.9, 0.4, 0.2; 0.7, 0.5, 0.2; 0.6, 0.6, 0.3; 0.2, 0.4, 0.8];

for i = 1:n_percentiles
    subplot(2, 3, i);
    pct = percentiles(i);
    group_data = groups_A{i};
    mdl = models_A{i};
    stats = stats_A_all{i};

    scatter(group_data.('역량진단점수'), group_data.('가중치적용점수'), 50, 'filled', ...
        'MarkerFaceColor', colors_A(i, :), 'MarkerFaceAlpha', 0.6);
    hold on;
    plot(mdl, 'LineWidth', 2);
    hold off;
    xlabel('역량진단 성과점수', 'FontSize', 11, 'FontWeight', 'bold');
    ylabel('가중치 적용 점수', 'FontSize', 11, 'FontWeight', 'bold');
    title(sprintf('그룹 A: 가중치 기준 상위 %d%% (n=%d)', pct, stats.n), ...
        'FontSize', 12, 'FontWeight', 'bold');
    legend({'데이터', '회귀선', '95% CI'}, 'Location', 'best', 'FontSize', 9);
    grid on;
    text_str = sprintf('r=%.3f\nR²=%.3f', stats.r_pearson, stats.rsquared);
    text(0.05, 0.95, text_str, 'Units', 'normalized', ...
        'VerticalAlignment', 'top', 'FontSize', 10, 'BackgroundColor', 'w');
end

sgtitle('백분위수별 비교: 가중치 적용 점수 기준 상위 그룹', 'FontSize', 16, 'FontWeight', 'bold');
fig1_path = fullfile(config.output_dir, sprintf('scatter_groupA_weighted_%s.png', config.timestamp));
saveas(fig1, fig1_path);
fprintf('  ✓ 그림 저장: scatter_groupA_weighted_%s.png\n', config.timestamp);
close(fig1);

% 그림 2: 백분위수별 산점도 (그룹 B - 원본 종합점수 기준)
fig2 = figure('Position', [100, 100, 1400, 900]);
colors_B = [0.2, 0.8, 0.2; 0.2, 0.7, 0.4; 0.3, 0.6, 0.5; 0.4, 0.5, 0.6; 0.2, 0.4, 0.8];

for i = 1:n_percentiles
    subplot(2, 3, i);
    pct = percentiles(i);
    group_data = groups_B{i};
    mdl = models_B{i};
    stats = stats_B_all{i};

    scatter(group_data.('역량진단점수'), group_data.('원본종합점수'), 50, 'filled', ...
        'MarkerFaceColor', colors_B(i, :), 'MarkerFaceAlpha', 0.6);
    hold on;
    plot(mdl, 'LineWidth', 2);
    hold off;
    xlabel('역량진단 성과점수', 'FontSize', 11, 'FontWeight', 'bold');
    ylabel('원본 종합점수', 'FontSize', 11, 'FontWeight', 'bold');
    title(sprintf('그룹 B: 원본 기준 상위 %d%% (n=%d)', pct, stats.n), ...
        'FontSize', 12, 'FontWeight', 'bold');
    legend({'데이터', '회귀선', '95% CI'}, 'Location', 'best', 'FontSize', 9);
    grid on;
    text_str = sprintf('r=%.3f\nR²=%.3f', stats.r_pearson, stats.rsquared);
    text(0.05, 0.95, text_str, 'Units', 'normalized', ...
        'VerticalAlignment', 'top', 'FontSize', 10, 'BackgroundColor', 'w');
end

sgtitle('백분위수별 비교: 원본 종합점수 기준 상위 그룹', 'FontSize', 16, 'FontWeight', 'bold');
fig2_path = fullfile(config.output_dir, sprintf('scatter_groupB_original_%s.png', config.timestamp));
saveas(fig2, fig2_path);
fprintf('  ✓ 그림 저장: scatter_groupB_original_%s.png\n', config.timestamp);
close(fig2);

% 그림 3: 백분위수별 상관계수 트렌드
fig3 = figure('Position', [100, 100, 1000, 600]);

% 상관계수 추출
r_pearson_A = cellfun(@(x) x.r_pearson, stats_A_all);
r_pearson_B = cellfun(@(x) x.r_pearson, stats_B_all);
r_spearman_A = cellfun(@(x) x.r_spearman, stats_A_all);
r_spearman_B = cellfun(@(x) x.r_spearman, stats_B_all);

% 그래프
plot(percentiles, r_pearson_A, '-o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.8, 0.2, 0.2], 'DisplayName', '그룹 A (가중치) - Pearson');
hold on;
plot(percentiles, r_pearson_B, '-s', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.2, 0.8, 0.2], 'DisplayName', '그룹 B (원본) - Pearson');
plot(percentiles, r_spearman_A, '--o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.6, 0.1, 0.1], 'DisplayName', '그룹 A (가중치) - Spearman');
plot(percentiles, r_spearman_B, '--s', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.1, 0.6, 0.1], 'DisplayName', '그룹 B (원본) - Spearman');
hold off;

xlabel('백분위수 (%)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('상관계수', 'FontSize', 13, 'FontWeight', 'bold');
title('백분위수별 상관계수 변화: 가중치 vs 원본', 'FontSize', 14, 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 11);
grid on;
xticks(percentiles);
xticklabels({'10%', '25%', '33%', '50%', '100%'});

fig3_path = fullfile(config.output_dir, sprintf('trend_correlation_comparison_%s.png', config.timestamp));
saveas(fig3, fig3_path);
fprintf('  ✓ 그림 저장: trend_correlation_comparison_%s.png\n', config.timestamp);
close(fig3);

% 그림 4: 백분위수별 R² 트렌드
fig4 = figure('Position', [100, 100, 1000, 600]);

% R² 추출
rsquared_A = cellfun(@(x) x.rsquared, stats_A_all);
rsquared_B = cellfun(@(x) x.rsquared, stats_B_all);

% 그래프
plot(percentiles, rsquared_A, '-o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.8, 0.2, 0.2], 'DisplayName', '그룹 A (가중치 적용)');
hold on;
plot(percentiles, rsquared_B, '-s', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.2, 0.8, 0.2], 'DisplayName', '그룹 B (원본 종합점수)');
hold off;

xlabel('백분위수 (%)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('R² (결정계수)', 'FontSize', 13, 'FontWeight', 'bold');
title('백분위수별 회귀 설명력 (R²) 변화', 'FontSize', 14, 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 11);
grid on;
xticks(percentiles);
xticklabels({'10%', '25%', '33%', '50%', '100%'});
ylim([0, 1]);

fig4_path = fullfile(config.output_dir, sprintf('trend_rsquared_comparison_%s.png', config.timestamp));
saveas(fig4, fig4_path);
fprintf('  ✓ 그림 저장: trend_rsquared_comparison_%s.png\n', config.timestamp);
close(fig4);

%% 8) 엑셀 결과 저장
fprintf('\n[STEP 8] 엑셀 결과 저장\n');
fprintf('-----------------------------------------------------\n');

excel_file = fullfile(config.output_dir, ...
    sprintf('가중치vs원본_역량검사_비교분석_%s.xlsx', config.timestamp));

% 시트 1: 전체 데이터
writetable(merged_data, excel_file, 'Sheet', '전체데이터', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 전체데이터 (%d행)\n', height(merged_data));

% 시트 2~6: 백분위수별 그룹 A 분석 결과
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_A_all{i};

    result_A = table();
    result_A.('항목') = {
        '백분위수'; '샘플수';
        '가중치점수_평균'; '가중치점수_표준편차';
        '역량진단점수_평균'; '역량진단점수_표준편차';
        'Pearson_r'; 'Pearson_p';
        'Spearman_rho'; 'Spearman_p';
        '회귀_R²'; '회귀_RMSE'; '회귀_절편'; '회귀_기울기'
        };
    result_A.('값') = {
        pct; stats.n;
        stats.weighted_mean; stats.weighted_std;
        stats.performance_mean; stats.performance_std;
        stats.r_pearson; stats.p_pearson;
        stats.r_spearman; stats.p_spearman;
        stats.rsquared; stats.rmse; stats.coef_intercept; stats.coef_slope
        };

    sheet_name = sprintf('그룹A_가중치상위%d%%', pct);
    writetable(result_A, excel_file, 'Sheet', sheet_name, 'WriteMode', 'append');
    fprintf('  ✓ 시트 저장: %s\n', sheet_name);
end

% 시트 7~11: 백분위수별 그룹 B 분석 결과
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_B_all{i};

    result_B = table();
    result_B.('항목') = {
        '백분위수'; '샘플수';
        '원본점수_평균'; '원본점수_표준편차';
        '역량진단점수_평균'; '역량진단점수_표준편차';
        'Pearson_r'; 'Pearson_p';
        'Spearman_rho'; 'Spearman_p';
        '회귀_R²'; '회귀_RMSE'; '회귀_절편'; '회귀_기울기'
        };
    result_B.('값') = {
        pct; stats.n;
        stats.original_mean; stats.original_std;
        stats.performance_mean; stats.performance_std;
        stats.r_pearson; stats.p_pearson;
        stats.r_spearman; stats.p_spearman;
        stats.rsquared; stats.rmse; stats.coef_intercept; stats.coef_slope
        };

    sheet_name = sprintf('그룹B_원본상위%d%%', pct);
    writetable(result_B, excel_file, 'Sheet', sheet_name, 'WriteMode', 'append');
    fprintf('  ✓ 시트 저장: %s\n', sheet_name);
end

% 시트 12: 그룹 A 백분위수별 요약 비교표
summary_A = table();
summary_A.('백분위수') = percentiles';
summary_A.('샘플수') = cellfun(@(x) x.n, stats_A_all);
summary_A.('가중치점수_평균') = cellfun(@(x) x.weighted_mean, stats_A_all);
summary_A.('역량진단점수_평균') = cellfun(@(x) x.performance_mean, stats_A_all);
summary_A.('Pearson_r') = cellfun(@(x) x.r_pearson, stats_A_all);
summary_A.('Pearson_p') = cellfun(@(x) x.p_pearson, stats_A_all);
summary_A.('Spearman_rho') = cellfun(@(x) x.r_spearman, stats_A_all);
summary_A.('R²') = cellfun(@(x) x.rsquared, stats_A_all);
summary_A.('RMSE') = cellfun(@(x) x.rmse, stats_A_all);

writetable(summary_A, excel_file, 'Sheet', '그룹A_가중치_백분위비교', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 그룹A_가중치_백분위비교\n');

% 시트 13: 그룹 B 백분위수별 요약 비교표
summary_B = table();
summary_B.('백분위수') = percentiles';
summary_B.('샘플수') = cellfun(@(x) x.n, stats_B_all);
summary_B.('원본점수_평균') = cellfun(@(x) x.original_mean, stats_B_all);
summary_B.('역량진단점수_평균') = cellfun(@(x) x.performance_mean, stats_B_all);
summary_B.('Pearson_r') = cellfun(@(x) x.r_pearson, stats_B_all);
summary_B.('Pearson_p') = cellfun(@(x) x.p_pearson, stats_B_all);
summary_B.('Spearman_rho') = cellfun(@(x) x.r_spearman, stats_B_all);
summary_B.('R²') = cellfun(@(x) x.rsquared, stats_B_all);
summary_B.('RMSE') = cellfun(@(x) x.rmse, stats_B_all);

writetable(summary_B, excel_file, 'Sheet', '그룹B_원본_백분위비교', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 그룹B_원본_백분위비교\n');

%% 9) 최종 요약
fprintf('\n[STEP 9] 최종 요약\n');
fprintf('=====================================================\n');
fprintf('📊 분석 완료!\n\n');
fprintf('📁 출력 디렉토리: %s\n', config.output_dir);
fprintf('📈 엑셀 파일: %s\n', sprintf('가중치vs원본_역량검사_비교분석_%s.xlsx', config.timestamp));
fprintf('\n');

% 백분위수별 요약
fprintf('【백분위수별 분석 결과 요약】\n');
fprintf('-----------------------------------------------------\n');
fprintf('  그룹 A (가중치 적용 점수 기준 상위 그룹):\n');
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_A_all{i};
    fprintf('    • 상위 %3d%%: n=%4d, r=%.3f, R²=%.3f\n', ...
        pct, stats.n, stats.r_pearson, stats.rsquared);
end
fprintf('\n');

fprintf('  그룹 B (원본 종합점수 기준 상위 그룹):\n');
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_B_all{i};
    fprintf('    • 상위 %3d%%: n=%4d, r=%.3f, R²=%.3f\n', ...
        pct, stats.n, stats.r_pearson, stats.rsquared);
end
fprintf('\n');

fprintf('【생성된 그림】\n');
fprintf('  • scatter_groupA_weighted_%s.png (가중치 기준 백분위수별)\n', config.timestamp);
fprintf('  • scatter_groupB_original_%s.png (원본 기준 백분위수별)\n', config.timestamp);
fprintf('  • trend_correlation_comparison_%s.png (상관계수 비교)\n', config.timestamp);
fprintf('  • trend_rsquared_comparison_%s.png (R² 비교)\n', config.timestamp);
fprintf('\n');

fprintf('【엑셀 시트 구성】\n');
fprintf('  1. 전체데이터 (원본 데이터)\n');
fprintf('  2~6. 그룹A 백분위수별 (가중치: 10%%, 25%%, 33%%, 50%%, 100%%)\n');
fprintf('  7~11. 그룹B 백분위수별 (원본: 10%%, 25%%, 33%%, 50%%, 100%%)\n');
fprintf('  12. 그룹A_가중치_백분위비교 (요약)\n');
fprintf('  13. 그룹B_원본_백분위비교 (요약)\n');
fprintf('\n');
fprintf('=====================================================\n');
fprintf('✅ 모든 작업 완료!\n');
