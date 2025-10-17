%% 가중치 적용 역량검사 점수 vs 역량진단 성과점수 비교 분석
%
% 목적:
%   - 가중치 적용 역량검사 점수와 역량진단 성과점수의 상관관계 분석
%   - 각 점수별 상위 50% 그룹의 특성 비교
%
% 입력:
%   - 가중치 적용 점수: D:\project\HR데이터\결과\자가불소_revised_talent\역량검사_가중치적용점수_talent_*.xlsx
%   - 역량진단 성과점수: D:\project\HR데이터\matlab\문항기반_revised\*_workspace_*.mat (integratedPerformanceData)
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
fprintf('  가중치 적용 역량검사 점수 vs 역량진단 성과점수 비교 분석\n');
fprintf('=====================================================\n\n');

%% 1) 설정
fprintf('[STEP 1] 설정\n');
fprintf('-----------------------------------------------------\n');

config = struct();
config.weighted_score_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent';
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

% 필요한 컬럼 확인 및 추출 (컬럼 위치로 접근)
% 1번 컬럼: ID, 3번 컬럼: 총점
weighted_scores = table();
weighted_scores.ID = weighted_data{:, 1};  % ID (1번 컬럼)
weighted_scores.('역량검사점수') = weighted_data{:, 3};  % 총점 (3번 컬럼)

fprintf('  ✓ 역량검사 점수: %d명 (평균 %.2f ± %.2f)\n', ...
    height(weighted_scores), ...
    mean(weighted_scores.('역량검사점수'), 'omitnan'), ...
    std(weighted_scores.('역량검사점수'), 'omitnan'));

%% 3) 역량진단 성과점수 로드
fprintf('\n[STEP 3] 역량진단 성과점수 로드\n');
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

%% 4) ID 매칭
fprintf('\n[STEP 4] ID 매칭\n');
fprintf('-----------------------------------------------------\n');

% ID 타입 통일 (cell → double)
if iscell(weighted_scores.ID)
    weighted_scores.ID = cellfun(@(x) str2double(x), weighted_scores.ID);
end
if iscell(performance_scores.ID)
    performance_scores.ID = cellfun(@(x) str2double(x), performance_scores.ID);
end

% Inner join으로 매칭
merged_data = innerjoin(weighted_scores, performance_scores, 'Keys', 'ID');

% 결측치 제거 (양쪽 점수가 모두 있는 경우만)
valid_idx = ~isnan(merged_data.('역량검사점수')) & ~isnan(merged_data.('역량진단점수'));
merged_data = merged_data(valid_idx, :);

fprintf('  ✓ 매칭 완료: %d명 (역검: %d명, 역진: %d명)\n', ...
    height(merged_data), height(weighted_scores), height(performance_scores));

% 점수 차이 계산
merged_data.('점수차이') = merged_data.('역량검사점수') - merged_data.('역량진단점수');

% 순위 계산
[~, rank_competency] = sort(merged_data.('역량검사점수'), 'descend');
[~, rank_performance] = sort(merged_data.('역량진단점수'), 'descend');
merged_data.('역량검사순위') = zeros(height(merged_data), 1);
merged_data.('역량진단순위') = zeros(height(merged_data), 1);
merged_data.('역량검사순위')(rank_competency) = (1:height(merged_data))';
merged_data.('역량진단순위')(rank_performance) = (1:height(merged_data))';

fprintf('  ✓ 점수 차이: 평균 %.2f ± %.2f (범위: %.2f ~ %.2f)\n', ...
    mean(merged_data.('점수차이')), std(merged_data.('점수차이')), ...
    min(merged_data.('점수차이')), max(merged_data.('점수차이')));

%% 5) 전체 데이터 상관분석
fprintf('\n[STEP 5] 전체 데이터 상관분석\n');
fprintf('-----------------------------------------------------\n');

% 기술통계
stats_all = struct();
stats_all.n = height(merged_data);
stats_all.competency_mean = mean(merged_data.('역량검사점수'));
stats_all.competency_std = std(merged_data.('역량검사점수'));
stats_all.competency_median = median(merged_data.('역량검사점수'));
stats_all.competency_min = min(merged_data.('역량검사점수'));
stats_all.competency_max = max(merged_data.('역량검사점수'));
stats_all.performance_mean = mean(merged_data.('역량진단점수'));
stats_all.performance_std = std(merged_data.('역량진단점수'));
stats_all.performance_median = median(merged_data.('역량진단점수'));
stats_all.performance_min = min(merged_data.('역량진단점수'));
stats_all.performance_max = max(merged_data.('역량진단점수'));

fprintf('  [기술통계]\n');
fprintf('    • 역량검사점수: %.2f ± %.2f (중앙값: %.2f, 범위: %.2f ~ %.2f)\n', ...
    stats_all.competency_mean, stats_all.competency_std, stats_all.competency_median, ...
    stats_all.competency_min, stats_all.competency_max);
fprintf('    • 역량진단점수: %.2f ± %.2f (중앙값: %.2f, 범위: %.2f ~ %.2f)\n', ...
    stats_all.performance_mean, stats_all.performance_std, stats_all.performance_median, ...
    stats_all.performance_min, stats_all.performance_max);

% Pearson 상관계수
[r_pearson, p_pearson] = corr(merged_data.('역량검사점수'), merged_data.('역량진단점수'), ...
                              'Type', 'Pearson');
stats_all.r_pearson = r_pearson;
stats_all.p_pearson = p_pearson;

fprintf('  [Pearson 상관]\n');
fprintf('    • r = %.4f, p = %.4e\n', r_pearson, p_pearson);

% Spearman 상관계수
[r_spearman, p_spearman] = corr(merged_data.('역량검사점수'), merged_data.('역량진단점수'), ...
                                'Type', 'Spearman');
stats_all.r_spearman = r_spearman;
stats_all.p_spearman = p_spearman;

fprintf('  [Spearman 상관]\n');
fprintf('    • ρ = %.4f, p = %.4e\n', r_spearman, p_spearman);

% 단순 회귀분석 (역량진단점수 → 역량검사점수)
mdl_all = fitlm(merged_data.('역량진단점수'), merged_data.('역량검사점수'));
stats_all.rsquared = mdl_all.Rsquared.Ordinary;
stats_all.rmse = mdl_all.RMSE;
stats_all.coef_intercept = mdl_all.Coefficients.Estimate(1);
stats_all.coef_slope = mdl_all.Coefficients.Estimate(2);
stats_all.coef_p = mdl_all.Coefficients.pValue(2);

fprintf('  [회귀분석: 역량검사점수 = β₀ + β₁×역량진단점수]\n');
fprintf('    • R² = %.4f\n', stats_all.rsquared);
fprintf('    • RMSE = %.4f\n', stats_all.rmse);
fprintf('    • 절편 = %.4f, 기울기 = %.4f (p = %.4e)\n', ...
    stats_all.coef_intercept, stats_all.coef_slope, stats_all.coef_p);

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

    %% 그룹 A: 역량검사 점수 기준 상위 X%
    [~, idx_competency_sorted] = sort(merged_data.('역량검사점수'), 'descend');
    group_A_idx = idx_competency_sorted(1:n_samples);
    group_A = merged_data(group_A_idx, :);
    groups_A{i} = group_A;

    fprintf('    [그룹 A: 역량검사 기준 상위 %d%%]\n', pct);
    fprintf('      • 샘플 수: %d명\n', height(group_A));
    fprintf('      • 역량검사점수 범위: %.2f ~ %.2f\n', ...
        min(group_A.('역량검사점수')), max(group_A.('역량검사점수')));

    % 그룹 A 통계
    stats_A = struct();
    stats_A.percentile = pct;
    stats_A.n = height(group_A);
    stats_A.competency_mean = mean(group_A.('역량검사점수'));
    stats_A.competency_std = std(group_A.('역량검사점수'));
    stats_A.competency_median = median(group_A.('역량검사점수'));
    stats_A.performance_mean = mean(group_A.('역량진단점수'));
    stats_A.performance_std = std(group_A.('역량진단점수'));
    stats_A.performance_median = median(group_A.('역량진단점수'));

    % 상관분석
    [r_pearson_A, p_pearson_A] = corr(group_A.('역량검사점수'), group_A.('역량진단점수'), ...
                                      'Type', 'Pearson');
    stats_A.r_pearson = r_pearson_A;
    stats_A.p_pearson = p_pearson_A;

    [r_spearman_A, p_spearman_A] = corr(group_A.('역량검사점수'), group_A.('역량진단점수'), ...
                                        'Type', 'Spearman');
    stats_A.r_spearman = r_spearman_A;
    stats_A.p_spearman = p_spearman_A;

    % 회귀분석
    mdl_A = fitlm(group_A.('역량진단점수'), group_A.('역량검사점수'));
    stats_A.rsquared = mdl_A.Rsquared.Ordinary;
    stats_A.rmse = mdl_A.RMSE;
    stats_A.coef_intercept = mdl_A.Coefficients.Estimate(1);
    stats_A.coef_slope = mdl_A.Coefficients.Estimate(2);

    stats_A_all{i} = stats_A;
    models_A{i} = mdl_A;

    fprintf('      • Pearson r = %.4f (p = %.4e)\n', r_pearson_A, p_pearson_A);
    fprintf('      • R² = %.4f, RMSE = %.4f\n', stats_A.rsquared, stats_A.rmse);

    %% 그룹 B: 역량진단 점수 기준 상위 X%
    [~, idx_performance_sorted] = sort(merged_data.('역량진단점수'), 'descend');
    group_B_idx = idx_performance_sorted(1:n_samples);
    group_B = merged_data(group_B_idx, :);
    groups_B{i} = group_B;

    fprintf('    [그룹 B: 역량진단 기준 상위 %d%%]\n', pct);
    fprintf('      • 샘플 수: %d명\n', height(group_B));
    fprintf('      • 역량진단점수 범위: %.2f ~ %.2f\n', ...
        min(group_B.('역량진단점수')), max(group_B.('역량진단점수')));

    % 그룹 B 통계
    stats_B = struct();
    stats_B.percentile = pct;
    stats_B.n = height(group_B);
    stats_B.competency_mean = mean(group_B.('역량검사점수'));
    stats_B.competency_std = std(group_B.('역량검사점수'));
    stats_B.competency_median = median(group_B.('역량검사점수'));
    stats_B.performance_mean = mean(group_B.('역량진단점수'));
    stats_B.performance_std = std(group_B.('역량진단점수'));
    stats_B.performance_median = median(group_B.('역량진단점수'));

    % 상관분석
    [r_pearson_B, p_pearson_B] = corr(group_B.('역량검사점수'), group_B.('역량진단점수'), ...
                                      'Type', 'Pearson');
    stats_B.r_pearson = r_pearson_B;
    stats_B.p_pearson = p_pearson_B;

    [r_spearman_B, p_spearman_B] = corr(group_B.('역량검사점수'), group_B.('역량진단점수'), ...
                                        'Type', 'Spearman');
    stats_B.r_spearman = r_spearman_B;
    stats_B.p_spearman = p_spearman_B;

    % 회귀분석
    mdl_B = fitlm(group_B.('역량진단점수'), group_B.('역량검사점수'));
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

% 그림 1: 전체 데이터 산점도 + 회귀선
fig1 = figure('Position', [100, 100, 800, 600]);
scatter(merged_data.('역량진단점수'), merged_data.('역량검사점수'), 50, 'filled', ...
    'MarkerFaceColor', [0.2, 0.4, 0.8], 'MarkerFaceAlpha', 0.6);
hold on;
plot(mdl_all, 'LineWidth', 2);
hold off;
xlabel('역량진단 성과점수', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('역량검사 점수 (가중치 적용)', 'FontSize', 13, 'FontWeight', 'bold');
title(sprintf('전체 데이터: 역량검사 vs 역량진단 (n=%d)', stats_all.n), ...
    'FontSize', 14, 'FontWeight', 'bold');
legend({'데이터', '회귀선', '95% 신뢰구간'}, 'Location', 'best', 'FontSize', 11);
grid on;
text_str = sprintf('r = %.3f (p < %.3f)\nR² = %.3f\nRMSE = %.2f', ...
    stats_all.r_pearson, stats_all.p_pearson, stats_all.rsquared, stats_all.rmse);
text(0.05, 0.95, text_str, 'Units', 'normalized', ...
    'VerticalAlignment', 'top', 'FontSize', 11, 'BackgroundColor', 'w');

fig1_path = fullfile(config.output_dir, sprintf('scatter_all_%s.png', config.timestamp));
saveas(fig1, fig1_path);
fprintf('  ✓ 그림 저장: scatter_all_%s.png\n', config.timestamp);
close(fig1);

% 그림 2: 백분위수별 산점도 (그룹 A - 역량검사 기준)
fig2 = figure('Position', [100, 100, 1400, 900]);
colors_A = [0.8, 0.2, 0.2; 0.9, 0.4, 0.2; 0.7, 0.5, 0.2; 0.6, 0.6, 0.3; 0.2, 0.4, 0.8];

for i = 1:n_percentiles
    subplot(2, 3, i);
    pct = percentiles(i);
    group_data = groups_A{i};
    mdl = models_A{i};
    stats = stats_A_all{i};

    scatter(group_data.('역량진단점수'), group_data.('역량검사점수'), 50, 'filled', ...
        'MarkerFaceColor', colors_A(i, :), 'MarkerFaceAlpha', 0.6);
    hold on;
    plot(mdl, 'LineWidth', 2);
    hold off;
    xlabel('역량진단 성과점수', 'FontSize', 11, 'FontWeight', 'bold');
    ylabel('역량검사 점수', 'FontSize', 11, 'FontWeight', 'bold');
    title(sprintf('그룹 A: 역량검사 상위 %d%% (n=%d)', pct, stats.n), ...
        'FontSize', 12, 'FontWeight', 'bold');
    legend({'데이터', '회귀선', '95% CI'}, 'Location', 'best', 'FontSize', 9);
    grid on;
    text_str = sprintf('r=%.3f\nR²=%.3f', stats.r_pearson, stats.rsquared);
    text(0.05, 0.95, text_str, 'Units', 'normalized', ...
        'VerticalAlignment', 'top', 'FontSize', 10, 'BackgroundColor', 'w');
end

sgtitle('백분위수별 비교: 역량검사 기준 상위 그룹', 'FontSize', 16, 'FontWeight', 'bold');
fig2_path = fullfile(config.output_dir, sprintf('scatter_groupA_percentiles_%s.png', config.timestamp));
saveas(fig2, fig2_path);
fprintf('  ✓ 그림 저장: scatter_groupA_percentiles_%s.png\n', config.timestamp);
close(fig2);

% 그림 3: 백분위수별 산점도 (그룹 B - 역량진단 기준)
fig3 = figure('Position', [100, 100, 1400, 900]);
colors_B = [0.2, 0.8, 0.2; 0.2, 0.7, 0.4; 0.3, 0.6, 0.5; 0.4, 0.5, 0.6; 0.2, 0.4, 0.8];

for i = 1:n_percentiles
    subplot(2, 3, i);
    pct = percentiles(i);
    group_data = groups_B{i};
    mdl = models_B{i};
    stats = stats_B_all{i};

    scatter(group_data.('역량진단점수'), group_data.('역량검사점수'), 50, 'filled', ...
        'MarkerFaceColor', colors_B(i, :), 'MarkerFaceAlpha', 0.6);
    hold on;
    plot(mdl, 'LineWidth', 2);
    hold off;
    xlabel('역량진단 성과점수', 'FontSize', 11, 'FontWeight', 'bold');
    ylabel('역량검사 점수', 'FontSize', 11, 'FontWeight', 'bold');
    title(sprintf('그룹 B: 역량진단 상위 %d%% (n=%d)', pct, stats.n), ...
        'FontSize', 12, 'FontWeight', 'bold');
    legend({'데이터', '회귀선', '95% CI'}, 'Location', 'best', 'FontSize', 9);
    grid on;
    text_str = sprintf('r=%.3f\nR²=%.3f', stats.r_pearson, stats.rsquared);
    text(0.05, 0.95, text_str, 'Units', 'normalized', ...
        'VerticalAlignment', 'top', 'FontSize', 10, 'BackgroundColor', 'w');
end

sgtitle('백분위수별 비교: 역량진단 기준 상위 그룹', 'FontSize', 16, 'FontWeight', 'bold');
fig3_path = fullfile(config.output_dir, sprintf('scatter_groupB_percentiles_%s.png', config.timestamp));
saveas(fig3, fig3_path);
fprintf('  ✓ 그림 저장: scatter_groupB_percentiles_%s.png\n', config.timestamp);
close(fig3);

% 그림 4: 백분위수별 상관계수 트렌드
fig4 = figure('Position', [100, 100, 1000, 600]);

% 상관계수 추출
r_pearson_A = cellfun(@(x) x.r_pearson, stats_A_all);
r_pearson_B = cellfun(@(x) x.r_pearson, stats_B_all);
r_spearman_A = cellfun(@(x) x.r_spearman, stats_A_all);
r_spearman_B = cellfun(@(x) x.r_spearman, stats_B_all);

% 그래프
plot(percentiles, r_pearson_A, '-o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.8, 0.2, 0.2], 'DisplayName', '그룹 A - Pearson');
hold on;
plot(percentiles, r_pearson_B, '-s', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.2, 0.8, 0.2], 'DisplayName', '그룹 B - Pearson');
plot(percentiles, r_spearman_A, '--o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.6, 0.1, 0.1], 'DisplayName', '그룹 A - Spearman');
plot(percentiles, r_spearman_B, '--s', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.1, 0.6, 0.1], 'DisplayName', '그룹 B - Spearman');
hold off;

xlabel('백분위수 (%)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('상관계수', 'FontSize', 13, 'FontWeight', 'bold');
title('백분위수별 상관계수 변화', 'FontSize', 14, 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 11);
grid on;
xticks(percentiles);
xticklabels({'10%', '25%', '33%', '50%', '100%'});

fig4_path = fullfile(config.output_dir, sprintf('trend_correlation_%s.png', config.timestamp));
saveas(fig4, fig4_path);
fprintf('  ✓ 그림 저장: trend_correlation_%s.png\n', config.timestamp);
close(fig4);

% 그림 5: 백분위수별 R² 트렌드
fig5 = figure('Position', [100, 100, 1000, 600]);

% R² 추출
rsquared_A = cellfun(@(x) x.rsquared, stats_A_all);
rsquared_B = cellfun(@(x) x.rsquared, stats_B_all);

% 그래프
plot(percentiles, rsquared_A, '-o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.8, 0.2, 0.2], 'DisplayName', '그룹 A (역량검사 기준)');
hold on;
plot(percentiles, rsquared_B, '-s', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.2, 0.8, 0.2], 'DisplayName', '그룹 B (역량진단 기준)');
hold off;

xlabel('백분위수 (%)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('R² (결정계수)', 'FontSize', 13, 'FontWeight', 'bold');
title('백분위수별 회귀 설명력 (R²) 변화', 'FontSize', 14, 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 11);
grid on;
xticks(percentiles);
xticklabels({'10%', '25%', '33%', '50%', '100%'});
ylim([0, 1]);

fig5_path = fullfile(config.output_dir, sprintf('trend_rsquared_%s.png', config.timestamp));
saveas(fig5, fig5_path);
fprintf('  ✓ 그림 저장: trend_rsquared_%s.png\n', config.timestamp);
close(fig5);

% 그림 6: 점수 차이 분포 (히스토그램)
fig6 = figure('Position', [100, 100, 800, 600]);
histogram(merged_data.('점수차이'), 30, 'FaceColor', [0.4, 0.4, 0.8], ...
    'EdgeColor', 'k', 'FaceAlpha', 0.7);
xlabel('점수 차이 (역량검사 - 역량진단)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('빈도', 'FontSize', 13, 'FontWeight', 'bold');
title(sprintf('점수 차이 분포 (n=%d)', stats_all.n), ...
    'FontSize', 14, 'FontWeight', 'bold');
grid on;
text_str = sprintf('평균 = %.2f\n표준편차 = %.2f\n범위 = [%.2f, %.2f]', ...
    mean(merged_data.('점수차이')), std(merged_data.('점수차이')), ...
    min(merged_data.('점수차이')), max(merged_data.('점수차이')));
text(0.70, 0.95, text_str, 'Units', 'normalized', ...
    'VerticalAlignment', 'top', 'FontSize', 11, 'BackgroundColor', 'w');

fig6_path = fullfile(config.output_dir, sprintf('histogram_diff_%s.png', config.timestamp));
saveas(fig6, fig6_path);
fprintf('  ✓ 그림 저장: histogram_diff_%s.png\n', config.timestamp);
close(fig6);

%% 8) 엑셀 결과 저장
fprintf('\n[STEP 8] 엑셀 결과 저장\n');
fprintf('-----------------------------------------------------\n');

excel_file = fullfile(config.output_dir, ...
    sprintf('역량검사vs역량진단_비교분석_%s.xlsx', config.timestamp));

% 시트 1: 전체 데이터
writetable(merged_data, excel_file, 'Sheet', '전체데이터', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 전체데이터 (%d행)\n', height(merged_data));

% 시트 2: 전체 데이터 분석 결과
result_all = table();
result_all.('항목') = {
    '샘플수';
    '역량검사점수_평균'; '역량검사점수_표준편차'; '역량검사점수_중앙값'; '역량검사점수_최소'; '역량검사점수_최대';
    '역량진단점수_평균'; '역량진단점수_표준편차'; '역량진단점수_중앙값'; '역량진단점수_최소'; '역량진단점수_최대';
    'Pearson_r'; 'Pearson_p';
    'Spearman_rho'; 'Spearman_p';
    '회귀_R²'; '회귀_RMSE'; '회귀_절편'; '회귀_기울기'; '회귀_p값'
    };
result_all.('값') = {
    stats_all.n;
    stats_all.competency_mean; stats_all.competency_std; stats_all.competency_median;
    stats_all.competency_min; stats_all.competency_max;
    stats_all.performance_mean; stats_all.performance_std; stats_all.performance_median;
    stats_all.performance_min; stats_all.performance_max;
    stats_all.r_pearson; stats_all.p_pearson;
    stats_all.r_spearman; stats_all.p_spearman;
    stats_all.rsquared; stats_all.rmse; stats_all.coef_intercept;
    stats_all.coef_slope; stats_all.coef_p
    };

writetable(result_all, excel_file, 'Sheet', '전체분석결과', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 전체분석결과\n');

% 시트 3~7: 백분위수별 그룹 A 분석 결과
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_A_all{i};

    result_A = table();
    result_A.('항목') = {
        '백분위수'; '샘플수';
        '역량검사점수_평균'; '역량검사점수_표준편차'; '역량검사점수_중앙값';
        '역량진단점수_평균'; '역량진단점수_표준편차'; '역량진단점수_중앙값';
        'Pearson_r'; 'Pearson_p';
        'Spearman_rho'; 'Spearman_p';
        '회귀_R²'; '회귀_RMSE'; '회귀_절편'; '회귀_기울기'
        };
    result_A.('값') = {
        pct; stats.n;
        stats.competency_mean; stats.competency_std; stats.competency_median;
        stats.performance_mean; stats.performance_std; stats.performance_median;
        stats.r_pearson; stats.p_pearson;
        stats.r_spearman; stats.p_spearman;
        stats.rsquared; stats.rmse; stats.coef_intercept; stats.coef_slope
        };

    sheet_name = sprintf('그룹A_역검상위%d%%', pct);
    writetable(result_A, excel_file, 'Sheet', sheet_name, 'WriteMode', 'append');
    fprintf('  ✓ 시트 저장: %s\n', sheet_name);
end

% 시트 8~12: 백분위수별 그룹 B 분석 결과
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_B_all{i};

    result_B = table();
    result_B.('항목') = {
        '백분위수'; '샘플수';
        '역량검사점수_평균'; '역량검사점수_표준편차'; '역량검사점수_중앙값';
        '역량진단점수_평균'; '역량진단점수_표준편차'; '역량진단점수_중앙값';
        'Pearson_r'; 'Pearson_p';
        'Spearman_rho'; 'Spearman_p';
        '회귀_R²'; '회귀_RMSE'; '회귀_절편'; '회귀_기울기'
        };
    result_B.('값') = {
        pct; stats.n;
        stats.competency_mean; stats.competency_std; stats.competency_median;
        stats.performance_mean; stats.performance_std; stats.performance_median;
        stats.r_pearson; stats.p_pearson;
        stats.r_spearman; stats.p_spearman;
        stats.rsquared; stats.rmse; stats.coef_intercept; stats.coef_slope
        };

    sheet_name = sprintf('그룹B_역진상위%d%%', pct);
    writetable(result_B, excel_file, 'Sheet', sheet_name, 'WriteMode', 'append');
    fprintf('  ✓ 시트 저장: %s\n', sheet_name);
end

% 시트 13: 그룹 A 백분위수별 요약 비교표
summary_A = table();
summary_A.('백분위수') = percentiles';
summary_A.('샘플수') = cellfun(@(x) x.n, stats_A_all);
summary_A.('역량검사점수_평균') = cellfun(@(x) x.competency_mean, stats_A_all);
summary_A.('역량진단점수_평균') = cellfun(@(x) x.performance_mean, stats_A_all);
summary_A.('Pearson_r') = cellfun(@(x) x.r_pearson, stats_A_all);
summary_A.('Pearson_p') = cellfun(@(x) x.p_pearson, stats_A_all);
summary_A.('Spearman_rho') = cellfun(@(x) x.r_spearman, stats_A_all);
summary_A.('R²') = cellfun(@(x) x.rsquared, stats_A_all);
summary_A.('RMSE') = cellfun(@(x) x.rmse, stats_A_all);

writetable(summary_A, excel_file, 'Sheet', '그룹A_백분위비교', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 그룹A_백분위비교\n');

% 시트 14: 그룹 B 백분위수별 요약 비교표
summary_B = table();
summary_B.('백분위수') = percentiles';
summary_B.('샘플수') = cellfun(@(x) x.n, stats_B_all);
summary_B.('역량검사점수_평균') = cellfun(@(x) x.competency_mean, stats_B_all);
summary_B.('역량진단점수_평균') = cellfun(@(x) x.performance_mean, stats_B_all);
summary_B.('Pearson_r') = cellfun(@(x) x.r_pearson, stats_B_all);
summary_B.('Pearson_p') = cellfun(@(x) x.p_pearson, stats_B_all);
summary_B.('Spearman_rho') = cellfun(@(x) x.r_spearman, stats_B_all);
summary_B.('R²') = cellfun(@(x) x.rsquared, stats_B_all);
summary_B.('RMSE') = cellfun(@(x) x.rmse, stats_B_all);

writetable(summary_B, excel_file, 'Sheet', '그룹B_백분위비교', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 그룹B_백분위비교\n');

%% 9) 최종 요약
fprintf('\n[STEP 9] 최종 요약\n');
fprintf('=====================================================\n');
fprintf('📊 분석 완료!\n\n');
fprintf('📁 출력 디렉토리: %s\n', config.output_dir);
fprintf('📈 엑셀 파일: %s\n', sprintf('역량검사vs역량진단_비교분석_%s.xlsx', config.timestamp));
fprintf('\n');
fprintf('【전체 데이터】\n');
fprintf('  • 샘플: %d명\n', stats_all.n);
fprintf('  • Pearson r = %.3f (p = %.3e)\n', stats_all.r_pearson, stats_all.p_pearson);
fprintf('  • R² = %.3f, RMSE = %.2f\n', stats_all.rsquared, stats_all.rmse);
fprintf('\n');

% 백분위수별 요약
fprintf('【백분위수별 분석 결과 요약】\n');
fprintf('-----------------------------------------------------\n');
fprintf('  그룹 A (역량검사 기준 상위 그룹):\n');
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_A_all{i};
    fprintf('    • 상위 %3d%%: n=%4d, r=%.3f, R²=%.3f\n', ...
        pct, stats.n, stats.r_pearson, stats.rsquared);
end
fprintf('\n');

fprintf('  그룹 B (역량진단 기준 상위 그룹):\n');
for i = 1:n_percentiles
    pct = percentiles(i);
    stats = stats_B_all{i};
    fprintf('    • 상위 %3d%%: n=%4d, r=%.3f, R²=%.3f\n', ...
        pct, stats.n, stats.r_pearson, stats.rsquared);
end
fprintf('\n');

fprintf('【생성된 그림】\n');
fprintf('  • scatter_all_%s.png (전체 데이터 산점도)\n', config.timestamp);
fprintf('  • scatter_groupA_percentiles_%s.png (그룹 A 백분위수별 산점도)\n', config.timestamp);
fprintf('  • scatter_groupB_percentiles_%s.png (그룹 B 백분위수별 산점도)\n', config.timestamp);
fprintf('  • trend_correlation_%s.png (백분위수별 상관계수 트렌드)\n', config.timestamp);
fprintf('  • trend_rsquared_%s.png (백분위수별 R² 트렌드)\n', config.timestamp);
fprintf('  • histogram_diff_%s.png (점수 차이 분포)\n', config.timestamp);
fprintf('\n');

fprintf('【엑셀 시트 구성】\n');
fprintf('  1. 전체데이터 (원본 데이터)\n');
fprintf('  2. 전체분석결과\n');
fprintf('  3~7. 그룹A 백분위수별 결과 (10%%, 25%%, 33%%, 50%%, 100%%)\n');
fprintf('  8~12. 그룹B 백분위수별 결과 (10%%, 25%%, 33%%, 50%%, 100%%)\n');
fprintf('  13. 그룹A_백분위비교 (요약 비교표)\n');
fprintf('  14. 그룹B_백분위비교 (요약 비교표)\n');
fprintf('\n');
fprintf('=====================================================\n');
fprintf('✅ 모든 작업 완료!\n');
