%% 가중치 적용 역량검사 점수 vs 역량진단 성과점수 비교 분석 (다중 백분위수)
%
% 목적:
%   - 가중치 적용 역량검사 점수와 역량진단 성과점수의 상관관계 분석
%   - 상위 10%, 25%, 50%, 100%(전체) 그룹별 특성 비교
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
fprintf('  가중치 적용 역량검사 점수 vs 역량진단 성과점수\n');
fprintf('  다중 백분위수 비교 분석 (10%%, 25%%, 50%%, 100%%)\n');
fprintf('=====================================================\n\n');

%% 1) 설정
fprintf('[STEP 1] 설정\n');
fprintf('-----------------------------------------------------\n');

config = struct();
config.weighted_score_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent';
config.performance_data_dir = 'D:\project\HR데이터\matlab\문항기반_revised';
config.output_dir = 'D:\project\HR데이터\결과\가중치vs역진';
config.timestamp = datestr(now, 'yyyy-mm-dd_HHMMSS');
config.percentiles = [10, 25, 50, 100];  % 분석할 백분위수

% 출력 디렉토리 생성
if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
    fprintf('  ✓ 출력 디렉토리 생성: %s\n', config.output_dir);
else
    fprintf('  ✓ 출력 디렉토리 확인: %s\n', config.output_dir);
end

fprintf('  ✓ 분석 백분위수: %s\n', strjoin(arrayfun(@(x) sprintf('%d%%', x), config.percentiles, 'UniformOutput', false), ', '));

%% 2) 가중치 적용 역량검사 점수 로드
fprintf('\n[STEP 2] 가중치 적용 역량검사 점수 로드\n');
fprintf('-----------------------------------------------------\n');

weighted_files = dir(fullfile(config.weighted_score_dir, '역량검사_가중치적용점수_talent*.xlsx'));
if isempty(weighted_files)
    error('가중치 적용 점수 파일을 찾을 수 없습니다: %s', config.weighted_score_dir);
end

[~, idx] = max([weighted_files.datenum]);
weighted_file = fullfile(weighted_files(idx).folder, weighted_files(idx).name);
fprintf('  ✓ 가중치 점수 파일: %s\n', weighted_files(idx).name);

weighted_data = readtable(weighted_file, 'Sheet', '역량검사_종합점수', ...
                         'VariableNamingRule', 'preserve');
fprintf('  ✓ 로드 완료: %d행 x %d열\n', height(weighted_data), width(weighted_data));

weighted_scores = table();
weighted_scores.ID = weighted_data{:, 1};
weighted_scores.('역량검사점수') = weighted_data{:, 3};

fprintf('  ✓ 역량검사 점수: %d명 (평균 %.2f ± %.2f)\n', ...
    height(weighted_scores), ...
    mean(weighted_scores.('역량검사점수'), 'omitnan'), ...
    std(weighted_scores.('역량검사점수'), 'omitnan'));

%% 3) 역량진단 성과점수 로드
fprintf('\n[STEP 3] 역량진단 성과점수 로드\n');
fprintf('-----------------------------------------------------\n');

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

if iscell(weighted_scores.ID)
    weighted_scores.ID = cellfun(@(x) str2double(x), weighted_scores.ID);
end
if iscell(performance_scores.ID)
    performance_scores.ID = cellfun(@(x) str2double(x), performance_scores.ID);
end

merged_data = innerjoin(weighted_scores, performance_scores, 'Keys', 'ID');
valid_idx = ~isnan(merged_data.('역량검사점수')) & ~isnan(merged_data.('역량진단점수'));
merged_data = merged_data(valid_idx, :);

fprintf('  ✓ 매칭 완료: %d명\n', height(merged_data));

merged_data.('점수차이') = merged_data.('역량검사점수') - merged_data.('역량진단점수');

%% 5) 다중 백분위수 분석
fprintf('\n[STEP 5] 다중 백분위수 분석\n');
fprintf('=====================================================\n');

n_total = height(merged_data);
results = struct();

% 각 백분위수별로 역량검사와 역량진단 기준 분석
for p = 1:length(config.percentiles)
    pct = config.percentiles(p);

    fprintf('\n▶ 상위 %d%% 분석\n', pct);
    fprintf('-----------------------------------------------------\n');

    if pct == 100
        n_samples = n_total;
    else
        n_samples = ceil(n_total * pct / 100);
    end

    % === 역량검사 기준 상위 그룹 ===
    [~, idx_comp_sorted] = sort(merged_data.('역량검사점수'), 'descend');
    group_comp = merged_data(idx_comp_sorted(1:n_samples), :);

    fprintf('  [역량검사 기준 상위 %d%%: %d명]\n', pct, height(group_comp));

    stats_comp = struct();
    stats_comp.n = height(group_comp);
    stats_comp.comp_mean = mean(group_comp.('역량검사점수'));
    stats_comp.comp_std = std(group_comp.('역량검사점수'));
    stats_comp.perf_mean = mean(group_comp.('역량진단점수'));
    stats_comp.perf_std = std(group_comp.('역량진단점수'));

    [r_comp, p_comp] = corr(group_comp.('역량검사점수'), group_comp.('역량진단점수'), 'Type', 'Pearson');
    [rho_comp, p_rho_comp] = corr(group_comp.('역량검사점수'), group_comp.('역량진단점수'), 'Type', 'Spearman');

    mdl_comp = fitlm(group_comp.('역량진단점수'), group_comp.('역량검사점수'));

    stats_comp.r_pearson = r_comp;
    stats_comp.p_pearson = p_comp;
    stats_comp.r_spearman = rho_comp;
    stats_comp.p_spearman = p_rho_comp;
    stats_comp.rsquared = mdl_comp.Rsquared.Ordinary;
    stats_comp.rmse = mdl_comp.RMSE;

    fprintf('    • 역량검사: %.2f ± %.2f\n', stats_comp.comp_mean, stats_comp.comp_std);
    fprintf('    • 역량진단: %.2f ± %.2f\n', stats_comp.perf_mean, stats_comp.perf_std);
    fprintf('    • Pearson r = %.4f (p = %.4f)\n', r_comp, p_comp);
    fprintf('    • R² = %.4f\n', stats_comp.rsquared);

    results.(sprintf('comp_top%d', pct)) = stats_comp;

    % === 역량진단 기준 상위 그룹 ===
    [~, idx_perf_sorted] = sort(merged_data.('역량진단점수'), 'descend');
    group_perf = merged_data(idx_perf_sorted(1:n_samples), :);

    fprintf('\n  [역량진단 기준 상위 %d%%: %d명]\n', pct, height(group_perf));

    stats_perf = struct();
    stats_perf.n = height(group_perf);
    stats_perf.comp_mean = mean(group_perf.('역량검사점수'));
    stats_perf.comp_std = std(group_perf.('역량검사점수'));
    stats_perf.perf_mean = mean(group_perf.('역량진단점수'));
    stats_perf.perf_std = std(group_perf.('역량진단점수'));

    [r_perf, p_perf] = corr(group_perf.('역량검사점수'), group_perf.('역량진단점수'), 'Type', 'Pearson');
    [rho_perf, p_rho_perf] = corr(group_perf.('역량검사점수'), group_perf.('역량진단점수'), 'Type', 'Spearman');

    mdl_perf = fitlm(group_perf.('역량진단점수'), group_perf.('역량검사점수'));

    stats_perf.r_pearson = r_perf;
    stats_perf.p_pearson = p_perf;
    stats_perf.r_spearman = rho_perf;
    stats_perf.p_spearman = p_rho_perf;
    stats_perf.rsquared = mdl_perf.Rsquared.Ordinary;
    stats_perf.rmse = mdl_perf.RMSE;

    fprintf('    • 역량검사: %.2f ± %.2f\n', stats_perf.comp_mean, stats_perf.comp_std);
    fprintf('    • 역량진단: %.2f ± %.2f\n', stats_perf.perf_mean, stats_perf.perf_std);
    fprintf('    • Pearson r = %.4f (p = %.4f)\n', r_perf, p_perf);
    fprintf('    • R² = %.4f\n', stats_perf.rsquared);

    results.(sprintf('perf_top%d', pct)) = stats_perf;
end

%% 6) 시각화 - 백분위수별 상관계수 변화
fprintf('\n[STEP 6] 시각화\n');
fprintf('-----------------------------------------------------\n');

% 상관계수 추출
pearson_comp = zeros(length(config.percentiles), 1);
pearson_perf = zeros(length(config.percentiles), 1);
rsquared_comp = zeros(length(config.percentiles), 1);
rsquared_perf = zeros(length(config.percentiles), 1);

for p = 1:length(config.percentiles)
    pct = config.percentiles(p);
    pearson_comp(p) = results.(sprintf('comp_top%d', pct)).r_pearson;
    pearson_perf(p) = results.(sprintf('perf_top%d', pct)).r_pearson;
    rsquared_comp(p) = results.(sprintf('comp_top%d', pct)).rsquared;
    rsquared_perf(p) = results.(sprintf('perf_top%d', pct)).rsquared;
end

% 그림 1: Pearson 상관계수 변화
fig1 = figure('Position', [100, 100, 900, 600]);
plot(config.percentiles, pearson_comp, '-o', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'Color', [0.8, 0.2, 0.2], 'MarkerFaceColor', [0.8, 0.2, 0.2]);
hold on;
plot(config.percentiles, pearson_perf, '-s', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'Color', [0.2, 0.4, 0.8], 'MarkerFaceColor', [0.2, 0.4, 0.8]);
yline(0, '--k', 'LineWidth', 1.5);
hold off;
xlabel('상위 백분위수 (%)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('Pearson 상관계수 (r)', 'FontSize', 13, 'FontWeight', 'bold');
title('백분위수별 상관계수 변화', 'FontSize', 14, 'FontWeight', 'bold');
legend({'역량검사 기준', '역량진단 기준'}, 'Location', 'best', 'FontSize', 12);
grid on;
xlim([5, 105]);
ylim([-0.3, 0.5]);

fig1_path = fullfile(config.output_dir, sprintf('correlation_by_percentile_%s.png', config.timestamp));
saveas(fig1, fig1_path);
fprintf('  ✓ 그림 저장: correlation_by_percentile_%s.png\n', config.timestamp);
close(fig1);

% 그림 2: R² 변화
fig2 = figure('Position', [100, 100, 900, 600]);
plot(config.percentiles, rsquared_comp, '-o', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'Color', [0.8, 0.2, 0.2], 'MarkerFaceColor', [0.8, 0.2, 0.2]);
hold on;
plot(config.percentiles, rsquared_perf, '-s', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'Color', [0.2, 0.4, 0.8], 'MarkerFaceColor', [0.2, 0.4, 0.8]);
hold off;
xlabel('상위 백분위수 (%)', 'FontSize', 13, 'FontWeight', 'bold');
ylabel('결정계수 (R²)', 'FontSize', 13, 'FontWeight', 'bold');
title('백분위수별 R² 변화', 'FontSize', 14, 'FontWeight', 'bold');
legend({'역량검사 기준', '역량진단 기준'}, 'Location', 'best', 'FontSize', 12);
grid on;
xlim([5, 105]);
ylim([0, 0.3]);

fig2_path = fullfile(config.output_dir, sprintf('rsquared_by_percentile_%s.png', config.timestamp));
saveas(fig2, fig2_path);
fprintf('  ✓ 그림 저장: rsquared_by_percentile_%s.png\n', config.timestamp);
close(fig2);

%% 7) 엑셀 결과 저장
fprintf('\n[STEP 7] 엑셀 결과 저장\n');
fprintf('-----------------------------------------------------\n');

excel_file = fullfile(config.output_dir, ...
    sprintf('역량검사vs역량진단_다중백분위수_%s.xlsx', config.timestamp));

% 시트 1: 요약 테이블
summary = table();
summary.('백분위수') = config.percentiles';
summary.('역검기준_샘플수') = arrayfun(@(x) results.(sprintf('comp_top%d', x)).n, config.percentiles)';
summary.('역검기준_Pearson_r') = pearson_comp;
summary.('역검기준_R2') = rsquared_comp;
summary.('역진기준_샘플수') = arrayfun(@(x) results.(sprintf('perf_top%d', x)).n, config.percentiles)';
summary.('역진기준_Pearson_r') = pearson_perf;
summary.('역진기준_R2') = rsquared_perf;

writetable(summary, excel_file, 'Sheet', '요약', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 요약\n');

% 시트 2: 상세 결과 (역량검사 기준)
for p = 1:length(config.percentiles)
    pct = config.percentiles(p);
    stats = results.(sprintf('comp_top%d', pct));

    detail = table();
    detail.('항목') = {'샘플수'; '역량검사_평균'; '역량검사_표준편차'; ...
                      '역량진단_평균'; '역량진단_표준편차'; ...
                      'Pearson_r'; 'Pearson_p'; 'Spearman_rho'; 'Spearman_p'; ...
                      'R²'; 'RMSE'};
    detail.('값') = {stats.n; stats.comp_mean; stats.comp_std; ...
                    stats.perf_mean; stats.perf_std; ...
                    stats.r_pearson; stats.p_pearson; stats.r_spearman; stats.p_spearman; ...
                    stats.rsquared; stats.rmse};

    writetable(detail, excel_file, 'Sheet', sprintf('역검기준_상위%d%%', pct), 'WriteMode', 'append');
end
fprintf('  ✓ 시트 저장: 역검기준_상위XX%%\n');

% 시트 3: 상세 결과 (역량진단 기준)
for p = 1:length(config.percentiles)
    pct = config.percentiles(p);
    stats = results.(sprintf('perf_top%d', pct));

    detail = table();
    detail.('항목') = {'샘플수'; '역량검사_평균'; '역량검사_표준편차'; ...
                      '역량진단_평균'; '역량진단_표준편차'; ...
                      'Pearson_r'; 'Pearson_p'; 'Spearman_rho'; 'Spearman_p'; ...
                      'R²'; 'RMSE'};
    detail.('값') = {stats.n; stats.comp_mean; stats.comp_std; ...
                    stats.perf_mean; stats.perf_std; ...
                    stats.r_pearson; stats.p_pearson; stats.r_spearman; stats.p_spearman; ...
                    stats.rsquared; stats.rmse};

    writetable(detail, excel_file, 'Sheet', sprintf('역진기준_상위%d%%', pct), 'WriteMode', 'append');
end
fprintf('  ✓ 시트 저장: 역진기준_상위XX%%\n');

%% 8) 최종 요약
fprintf('\n[STEP 8] 최종 요약\n');
fprintf('=====================================================\n');
fprintf('📊 다중 백분위수 분석 완료!\n\n');
fprintf('📁 출력 파일: %s\n', sprintf('역량검사vs역량진단_다중백분위수_%s.xlsx', config.timestamp));
fprintf('\n');
fprintf('【백분위수별 상관계수 (역량검사 기준)】\n');
for p = 1:length(config.percentiles)
    pct = config.percentiles(p);
    fprintf('  • 상위 %3d%%: r = %6.3f, R² = %.3f (n=%d)\n', ...
        pct, pearson_comp(p), rsquared_comp(p), results.(sprintf('comp_top%d', pct)).n);
end
fprintf('\n');
fprintf('【백분위수별 상관계수 (역량진단 기준)】\n');
for p = 1:length(config.percentiles)
    pct = config.percentiles(p);
    fprintf('  • 상위 %3d%%: r = %6.3f, R² = %.3f (n=%d)\n', ...
        pct, pearson_perf(p), rsquared_perf(p), results.(sprintf('perf_top%d', pct)).n);
end
fprintf('\n');
fprintf('=====================================================\n');
fprintf('✅ 모든 작업 완료!\n');
