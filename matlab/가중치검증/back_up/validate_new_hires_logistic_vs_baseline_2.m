% =======================================================================
%     신규 입사자 가중치 검증: Logistic vs 단순합 vs 기존 종합점수
% =======================================================================
% 목적:
%   - 신규 입사자 18명에 대해 Logistic 가중치 점수가 합격/불합격을
%     얼마나 잘 예측하는지 검증
%   - 단순합 점수 대비 개선도 측정
%   - 기존 종합점수와 비교 (12명 한정)
%
% 데이터:
%   - 신규 입사자: 18명 (합격 11명, 불합격 7명)
%   - 제외: 64006610 (조기 이탈)
%   - 기존 학습 데이터: 126명
%
% 작성일: 2025-10-16
% =======================================================================

clear; clc; close all;

% ---- 전역 폰트 설정 ----------------------------------------------------
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 10);
set(0, 'DefaultTextFontSize', 10);

fprintf('================================================================\n');
fprintf('   신규 입사자 가중치 검증: Logistic vs 단순합 vs 기존 종합점수\n');
fprintf('================================================================\n\n');

%% 1) 설정 ---------------------------------------------------------------
fprintf('【STEP 1】 설정\n');
fprintf('================================================================\n');

config = struct();
config.new_comp_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_역검 점수.xlsx';
config.new_onboarding_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_온보딩 점수.xlsx';
config.weight_file = 'D:\project\HR데이터\결과\자가불소_revised_talent\integrated_analysis_results.mat';
config.existing_score_file = 'D:\project\HR데이터\결과\자가불소_revised_talent\backup\역량검사_가중치적용점수_talent_2025-10-16_104347.xlsx';
config.output_dir = 'D:\project\HR데이터\matlab\가중치검증';
config.timestamp = datestr(now, 'yyyymmdd_HHMMSS');
config.output_filename = sprintf('validate_new_hires_logistic_vs_baseline_%s.xlsx', config.timestamp);
config.plot_filename = sprintf('validation_plot_%s.png', config.timestamp);

if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

fprintf('  ✓ 출력 디렉토리: %s\n', config.output_dir);

%% 2) 가중치 로드 --------------------------------------------------------
fprintf('\n【STEP 2】 가중치 로드\n');
fprintf('================================================================\n');

loaded = load(config.weight_file);
integrated_results = loaded.integrated_results;
weight_comp = integrated_results.weight_comparison;

feature_names = string(weight_comp.Feature);
logistic_weights = weight_comp.Logistic(:);

fprintf('  ✓ 역량 수: %d개\n', length(feature_names));
fprintf('  ✓ Logistic 가중치 범위: %.2f%% ~ %.2f%%\n', ...
    min(logistic_weights), max(logistic_weights));
fprintf('  ✓ 역량명: %s\n', strjoin(feature_names, ', '));

%% 3) 기존 학습 데이터 vs 신규 데이터 분포 비교 ---------------------------
fprintf('\n【STEP 3】 기존 학습 데이터 vs 신규 데이터 분포 비교\n');
fprintf('================================================================\n');

% 3-1. 기존 학습 데이터 로드
fprintf('\n  【3-1. 기존 학습 데이터 로드】\n');
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';

hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');

fprintf('    • HR 데이터: %d명\n', height(hr_data));
fprintf('    • 역량검사 데이터: %d명\n', height(comp_upper));

% 인재유형 필터링 (위장형 소화성 제외)
talent_col_idx = find(contains(hr_data.Properties.VariableNames, '인재유형'), 1);
if ~isempty(talent_col_idx)
    talent_types = hr_data{:, talent_col_idx};
    exclude_types = {'위장형 소화성'};
    valid_talent_idx = ~ismember(talent_types, exclude_types);
    hr_data = hr_data(valid_talent_idx, :);
    fprintf('    • 위장형 소화성 제외 후: %d명\n', height(hr_data));
end

% ID 매칭
hr_ids = hr_data.ID;
comp_ids = comp_upper.ID;
[matched_ids_train, hr_idx, comp_idx] = intersect(hr_ids, comp_ids);

matched_comp_train = comp_upper(comp_idx, :);
fprintf('    • ID 매칭 완료: %d명\n', length(matched_ids_train));

% 역량 데이터 추출
X_train = [];
for i = 1:length(feature_names)
    fn = char(feature_names(i));
    if ismember(fn, matched_comp_train.Properties.VariableNames)
        col = matched_comp_train.(fn);
        if ~isnumeric(col)
            col = double(col);
        end
        X_train = [X_train, col(:)]; %#ok<AGROW>
    end
end

fprintf('    • 역량 데이터: %d명 × %d개 역량\n', size(X_train, 1), size(X_train, 2));

% 3-2. 신규 입사자 데이터 로드
fprintf('\n  【3-2. 신규 입사자 데이터 로드】\n');
new_comp_data = readtable(config.new_comp_file, 'Sheet', '역량검사_상위항목', ...
    'VariableNamingRule', 'preserve');
fprintf('    • 신규 데이터: %d명\n', height(new_comp_data));

% 역량 데이터 추출
X_new_comp = [];
for i = 1:length(feature_names)
    fn = char(feature_names(i));
    if ismember(fn, new_comp_data.Properties.VariableNames)
        col = new_comp_data.(fn);
        if ~isnumeric(col)
            col = double(col);
        end
        X_new_comp = [X_new_comp, col(:)]; %#ok<AGROW>
    end
end

fprintf('    • 역량 데이터: %d명 × %d개 역량\n', size(X_new_comp, 1), size(X_new_comp, 2));

% 3-3. 분포 비교 분석
fprintf('\n  【3-3. 역량별 분포 비교】\n');
fprintf('    %-15s %12s %12s %10s %10s\n', '역량명', '기존평균±SD', '신규평균±SD', 't-test p', 'Cohen''s d');
fprintf('    %s\n', repmat('-', 1, 70));

distribution_comparison = table();
for i = 1:length(feature_names)
    fn = char(feature_names(i));

    % 기존 데이터
    train_vals = X_train(:, i);
    train_vals = train_vals(~isnan(train_vals));
    mean_train = mean(train_vals);
    std_train = std(train_vals);

    % 신규 데이터
    new_vals = X_new_comp(:, i);
    new_vals = new_vals(~isnan(new_vals));
    mean_new = mean(new_vals);
    std_new = std(new_vals);

    % t-test
    [~, p_val] = ttest2(train_vals, new_vals);

    % Cohen's d
    pooled_std = sqrt(((length(train_vals)-1)*var(train_vals) + ...
        (length(new_vals)-1)*var(new_vals)) / (length(train_vals) + length(new_vals) - 2));
    cohens_d = (mean_train - mean_new) / pooled_std;

    fprintf('    %-15s %5.1f±%4.1f    %5.1f±%4.1f    %8.4f   %8.3f\n', ...
        fn, mean_train, std_train, mean_new, std_new, p_val, cohens_d);

    distribution_comparison = [distribution_comparison; table({fn}, mean_train, std_train, ...
        mean_new, std_new, p_val, cohens_d, ...
        'VariableNames', {'역량명', '기존_평균', '기존_표준편차', '신규_평균', ...
        '신규_표준편차', 'p_value', 'Cohens_d'})]; %#ok<AGROW>
end

fprintf('\n    ✓ 분포 비교 완료\n');

% 3-4. 시각화
fprintf('\n  【3-4. 분포 시각화】\n');
fig_dist = figure('Position', [100, 100, 1600, 1000]);

n_comps = length(feature_names);
n_rows = ceil(n_comps / 3);

for i = 1:n_comps
    subplot(n_rows, 3, i);

    train_vals = X_train(:, i);
    train_vals = train_vals(~isnan(train_vals));
    new_vals = X_new_comp(:, i);
    new_vals = new_vals(~isnan(new_vals));

    % 히스토그램 오버레이
    hold on;
    histogram(train_vals, 15, 'FaceColor', [0.3 0.5 0.8], 'FaceAlpha', 0.6, ...
        'EdgeColor', 'none', 'Normalization', 'probability');
    histogram(new_vals, 10, 'FaceColor', [0.9 0.4 0.3], 'FaceAlpha', 0.6, ...
        'EdgeColor', 'none', 'Normalization', 'probability');

    title(sprintf('%s', char(feature_names(i))), 'FontSize', 10);
    xlabel('점수');
    ylabel('비율');
    legend({sprintf('기존 (n=%d)', length(train_vals)), ...
        sprintf('신규 (n=%d)', length(new_vals))}, 'Location', 'best', 'FontSize', 8);
    grid on;
    hold off;
end

dist_plot_path = fullfile(config.output_dir, ...
    sprintf('distribution_comparison_%s.png', config.timestamp));
saveas(fig_dist, dist_plot_path);
fprintf('    ✓ 분포 그래프 저장: distribution_comparison_%s.png\n', config.timestamp);
close(fig_dist);

% 3-5. 요약 통계
fprintf('\n  【3-5. 분포 차이 요약】\n');
sig_diff = sum(distribution_comparison.p_value < 0.05);
large_effect = sum(abs(distribution_comparison.Cohens_d) > 0.8);
medium_effect = sum(abs(distribution_comparison.Cohens_d) > 0.5 & abs(distribution_comparison.Cohens_d) <= 0.8);

fprintf('    • 통계적으로 유의한 차이 (p<0.05): %d/%d개 역량\n', sig_diff, n_comps);
fprintf('    • Large effect (|d|>0.8): %d개\n', large_effect);
fprintf('    • Medium effect (0.5<|d|≤0.8): %d개\n', medium_effect);

if sig_diff == 0
    fprintf('\n    ✅ 기존 학습 데이터와 신규 데이터의 분포가 유사합니다.\n');
    fprintf('       → 학습된 가중치를 신규 데이터에 적용하는 것이 타당합니다.\n');
else
    fprintf('\n    ⚠ 일부 역량에서 분포 차이가 발견되었습니다.\n');
    fprintf('      → 가중치 적용 시 주의가 필요합니다.\n');
end

%% 4) 신규 입사자 데이터 로드 ---------------------------------------------
fprintf('\n【STEP 4】 신규 입사자 데이터 로드\n');
fprintf('================================================================\n');

% 역량검사_상위항목 (역량 점수)
comp_data = readtable(config.new_comp_file, 'Sheet', '역량검사_상위항목', ...
                      'VariableNamingRule', 'preserve');
fprintf('  ✓ 역량검사_상위항목: %d명\n', height(comp_data));

% 역량검사_종합점수 (기존 종합점수)
score_data = readtable(config.new_comp_file, 'Sheet', '역량검사_종합점수', ...
                       'VariableNamingRule', 'preserve');
fprintf('  ✓ 역량검사_종합점수: %d명\n', height(score_data));

% 온보딩 점수 (합불 여부)
onboarding_data = readtable(config.new_onboarding_file, 'VariableNamingRule', 'preserve');
fprintf('  ✓ 온보딩 점수 (합불 여부): %d명\n', height(onboarding_data));

%% 5) 레이블 생성 및 필터링 (합격=1, 불합격=0) -----------------------------
fprintf('\n【STEP 5】 레이블 생성 및 데이터 필터링\n');
fprintf('================================================================\n');

n_new = height(comp_data);
labels = zeros(n_new, 1);
exclude_mask = false(n_new, 1);

for i = 1:n_new
    id = comp_data.ID{i};

    % 온보딩 데이터에서 합불 여부 찾기
    onb_idx = find(strcmp(onboarding_data.ID, id), 1);
    if ~isempty(onb_idx)
        pass_fail = onboarding_data.('합불 여부'){onb_idx};
        if strcmp(pass_fail, '합격')
            labels(i) = 1;
        elseif contains(pass_fail, '불합격') && contains(pass_fail, '조기')
            % 조기 이탈은 분석에서 제외
            exclude_mask(i) = true;
            fprintf('  • 제외: %s - "%s"\n', id, pass_fail);
        else
            labels(i) = 0;
            fprintf('  • 불합격: %s - "%s"\n', id, pass_fail);
        end
    end
end

% 제외 대상 필터링
comp_data = comp_data(~exclude_mask, :);
labels = labels(~exclude_mask);
n_new = height(comp_data);

pass_count = sum(labels == 1);
fail_count = sum(labels == 0);

fprintf('  ✓ 최종 분석 대상: %d명\n', n_new);
fprintf('  ✓ 합격: %d명 (%.1f%%)\n', pass_count, pass_count/n_new*100);
fprintf('  ✓ 불합격: %d명 (%.1f%%)\n', fail_count, fail_count/n_new*100);

%% 6) 신규 입사자 점수 계산 -----------------------------------------------
fprintf('\n【STEP 6】 신규 입사자 점수 계산\n');
fprintf('================================================================\n');

% 역량 컬럼 매칭
X_new = [];
matched_features = {};
matched_weights = [];

for i = 1:length(feature_names)
    fn = char(feature_names(i));
    if ismember(fn, comp_data.Properties.VariableNames)
        col = comp_data.(fn);
        if ~isnumeric(col)
            col = double(col);
        end
        X_new = [X_new, col(:)]; %#ok<AGROW>
        matched_features{end+1} = fn; %#ok<AGROW>
        matched_weights(end+1) = logistic_weights(i); %#ok<AGROW>
    end
end

fprintf('  ✓ 매칭된 역량: %d/%d개\n', length(matched_features), length(feature_names));

% 1) Logistic 가중치 점수
w = matched_weights(:) / 100;
score_logistic = nansum(X_new .* repmat(w', n_new, 1), 2) ./ sum(w);

% 2) 단순합 점수 (동일 가중치)
score_simple = nanmean(X_new, 2);

% 3) 기존 종합점수 (12명만)
score_existing = nan(n_new, 1);
for i = 1:n_new
    id = comp_data.ID{i};
    score_idx = find(strcmp(score_data.ID, id), 1);
    if ~isempty(score_idx)
        score_existing(i) = score_data.('종합점수')(score_idx);
    end
end

fprintf('\n  【Logistic 가중치 점수】\n');
fprintf('    • 합격 평균: %.2f ± %.2f\n', mean(score_logistic(labels==1)), std(score_logistic(labels==1)));
fprintf('    • 불합격 평균: %.2f ± %.2f\n', mean(score_logistic(labels==0)), std(score_logistic(labels==0)));

fprintf('\n  【단순합 점수】\n');
fprintf('    • 합격 평균: %.2f ± %.2f\n', mean(score_simple(labels==1)), std(score_simple(labels==1)));
fprintf('    • 불합격 평균: %.2f ± %.2f\n', mean(score_simple(labels==0)), std(score_simple(labels==0)));

valid_existing_count = sum(~isnan(score_existing));
fprintf('\n  【기존 종합점수】 (%d명)\n', valid_existing_count);
if valid_existing_count > 0
    fprintf('    • 합격 평균: %.2f ± %.2f\n', mean(score_existing(labels==1), 'omitnan'), std(score_existing(labels==1), 'omitnan'));
    fprintf('    • 불합격 평균: %.2f ± %.2f\n', mean(score_existing(labels==0), 'omitnan'), std(score_existing(labels==0), 'omitnan'));

    % 기존 종합점수와 새 점수 간 차이 분석
    fprintf('\n  【점수 차이 분석】 (%d명)\n', valid_existing_count);
    valid_idx = ~isnan(score_existing);
    diff_logistic = score_logistic(valid_idx) - score_existing(valid_idx);
    diff_simple = score_simple(valid_idx) - score_existing(valid_idx);

    fprintf('    • Logistic - 기존 종합점수:\n');
    fprintf('      평균 차이: %.2f (SD: %.2f)\n', mean(diff_logistic), std(diff_logistic));
    fprintf('      범위: %.2f ~ %.2f\n', min(diff_logistic), max(diff_logistic));

    fprintf('    • 단순합 - 기존 종합점수:\n');
    fprintf('      평균 차이: %.2f (SD: %.2f)\n', mean(diff_simple), std(diff_simple));
    fprintf('      범위: %.2f ~ %.2f\n', min(diff_simple), max(diff_simple));
else
    fprintf('    • 기존 종합점수 데이터 없음\n');
end

%% 7) ROC-AUC 및 PR-AUC 분석 ---------------------------------------------
fprintf('\n【STEP 7】 ROC-AUC 및 PR-AUC 분석\n');
fprintf('================================================================\n');

% ROC 곡선
[X_roc_logistic, Y_roc_logistic, ~, AUC_logistic] = perfcurve(labels, score_logistic, 1);
[X_roc_simple, Y_roc_simple, ~, AUC_simple] = perfcurve(labels, score_simple, 1);

fprintf('  【ROC-AUC】\n');
fprintf('    • Logistic 가중치: %.4f\n', AUC_logistic);
fprintf('    • 단순합: %.4f\n', AUC_simple);
fprintf('    • 개선도: %.4f (%.1f%%)\n', AUC_logistic - AUC_simple, ...
    (AUC_logistic/AUC_simple - 1)*100);

% PR 곡선
[X_pr_logistic, Y_pr_logistic, ~, AUC_pr_logistic] = perfcurve(labels, score_logistic, 1, ...
    'XCrit', 'reca', 'YCrit', 'prec');
[X_pr_simple, Y_pr_simple, ~, AUC_pr_simple] = perfcurve(labels, score_simple, 1, ...
    'XCrit', 'reca', 'YCrit', 'prec');

fprintf('\n  【PR-AUC】\n');
fprintf('    • Logistic 가중치: %.4f\n', AUC_pr_logistic);
fprintf('    • 단순합: %.4f\n', AUC_pr_simple);
fprintf('    • 개선도: %.4f (%.1f%%)\n', AUC_pr_logistic - AUC_pr_simple, ...
    (AUC_pr_logistic/AUC_pr_simple - 1)*100);

% 기존 종합점수 (12명만)
valid_existing = ~isnan(score_existing);
if sum(valid_existing) >= 5
    [X_roc_existing, Y_roc_existing, ~, AUC_existing] = perfcurve(labels(valid_existing), ...
        score_existing(valid_existing), 1);
    [X_pr_existing, Y_pr_existing, ~, AUC_pr_existing] = perfcurve(labels(valid_existing), ...
        score_existing(valid_existing), 1, 'XCrit', 'reca', 'YCrit', 'prec');

    fprintf('\n  【기존 종합점수 ROC-AUC】 (%d명)\n', sum(valid_existing));
    fprintf('    • AUC: %.4f\n', AUC_existing);
else
    AUC_existing = NaN;
    AUC_pr_existing = NaN;
    X_roc_existing = [];
    Y_roc_existing = [];
    X_pr_existing = [];
    Y_pr_existing = [];
end

%% 8) Top-K Precision 분석 ----------------------------------------------
fprintf('\n【STEP 8】 Top-K Precision 분석\n');
fprintf('================================================================\n');

k_values = [3, 5, 8, 10];
topk_results = table();

for k = k_values
    if k > n_new
        k = n_new;
    end

    % Logistic
    [~, idx_logistic] = sort(score_logistic, 'descend');
    top_k_logistic = labels(idx_logistic(1:k));
    precision_logistic = sum(top_k_logistic) / k * 100;

    % 단순합
    [~, idx_simple] = sort(score_simple, 'descend');
    top_k_simple = labels(idx_simple(1:k));
    precision_simple = sum(top_k_simple) / k * 100;

    improvement = precision_logistic - precision_simple;

    fprintf('  【상위 %d명 선발】\n', k);
    fprintf('    • Logistic: %.1f%% (%d/%d명 합격)\n', precision_logistic, sum(top_k_logistic), k);
    fprintf('    • 단순합: %.1f%% (%d/%d명 합격)\n', precision_simple, sum(top_k_simple), k);
    fprintf('    • 개선: %.1f%%p\n\n', improvement);

    topk_results = [topk_results; table(k, precision_logistic, precision_simple, improvement, ...
        'VariableNames', {'K', 'Logistic_정밀도', '단순합_정밀도', '개선도_퍼센트포인트'})]; %#ok<AGROW>
end

%% 9) 통계적 검정 -------------------------------------------------------
fprintf('\n【STEP 9】 통계적 검정 (t-test, Cohen''s d)\n');
fprintf('================================================================\n');

% t-test: 합격 vs 불합격
[~, p_logistic, ~, stats_logistic] = ttest2(score_logistic(labels==1), score_logistic(labels==0));
[~, p_simple, ~, stats_simple] = ttest2(score_simple(labels==1), score_simple(labels==0));

% Cohen's d (효과 크기)
mean_pass_logistic = mean(score_logistic(labels==1));
mean_fail_logistic = mean(score_logistic(labels==0));
std_pooled_logistic = sqrt(((sum(labels==1)-1)*var(score_logistic(labels==1)) + ...
    (sum(labels==0)-1)*var(score_logistic(labels==0))) / (n_new - 2));
cohens_d_logistic = (mean_pass_logistic - mean_fail_logistic) / std_pooled_logistic;

mean_pass_simple = mean(score_simple(labels==1));
mean_fail_simple = mean(score_simple(labels==0));
std_pooled_simple = sqrt(((sum(labels==1)-1)*var(score_simple(labels==1)) + ...
    (sum(labels==0)-1)*var(score_simple(labels==0))) / (n_new - 2));
cohens_d_simple = (mean_pass_simple - mean_fail_simple) / std_pooled_simple;

fprintf('  【Logistic 가중치 점수】\n');
fprintf('    • t-test p-value: %.4f\n', p_logistic);
fprintf('    • Cohen''s d: %.3f', cohens_d_logistic);
if abs(cohens_d_logistic) > 0.8
    fprintf(' (Large effect)\n');
elseif abs(cohens_d_logistic) > 0.5
    fprintf(' (Medium effect)\n');
else
    fprintf(' (Small effect)\n');
end

fprintf('\n  【단순합 점수】\n');
fprintf('    • t-test p-value: %.4f\n', p_simple);
fprintf('    • Cohen''s d: %.3f', cohens_d_simple);
if abs(cohens_d_simple) > 0.8
    fprintf(' (Large effect)\n');
elseif abs(cohens_d_simple) > 0.5
    fprintf(' (Medium effect)\n');
else
    fprintf(' (Small effect)\n');
end

% 기존 종합점수 통계 검정 (데이터가 있는 경우)
valid_existing_idx = ~isnan(score_existing);
if sum(valid_existing_idx) >= 5 && sum(labels(valid_existing_idx)==1) >= 2 && sum(labels(valid_existing_idx)==0) >= 2
    [~, p_existing, ~, ~] = ttest2(score_existing(labels==1 & valid_existing_idx), ...
        score_existing(labels==0 & valid_existing_idx));

    mean_pass_existing = mean(score_existing(labels==1 & valid_existing_idx), 'omitnan');
    mean_fail_existing = mean(score_existing(labels==0 & valid_existing_idx), 'omitnan');

    n_pass_existing = sum(labels==1 & valid_existing_idx);
    n_fail_existing = sum(labels==0 & valid_existing_idx);

    std_pooled_existing = sqrt(((n_pass_existing-1)*var(score_existing(labels==1 & valid_existing_idx), 'omitnan') + ...
        (n_fail_existing-1)*var(score_existing(labels==0 & valid_existing_idx), 'omitnan')) / ...
        (n_pass_existing + n_fail_existing - 2));
    cohens_d_existing = (mean_pass_existing - mean_fail_existing) / std_pooled_existing;

    fprintf('\n  【기존 종합점수】 (%d명)\n', sum(valid_existing_idx));
    fprintf('    • t-test p-value: %.4f\n', p_existing);
    fprintf('    • Cohen''s d: %.3f', cohens_d_existing);
    if abs(cohens_d_existing) > 0.8
        fprintf(' (Large effect)\n');
    elseif abs(cohens_d_existing) > 0.5
        fprintf(' (Medium effect)\n');
    else
        fprintf(' (Small effect)\n');
    end

    fprintf('\n  【효과 크기 비교】\n');
    fprintf('    • Logistic가 기존 종합점수보다 나음: %s\n', ...
        string(abs(cohens_d_logistic) > abs(cohens_d_existing)));
    fprintf('    • 단순합이 기존 종합점수보다 나음: %s\n', ...
        string(abs(cohens_d_simple) > abs(cohens_d_existing)));
else
    p_existing = NaN;
    cohens_d_existing = NaN;
end

%% 10) 상관관계 분석 ----------------------------------------------------
fprintf('\n【STEP 10】 상관관계 분석 (종합점수 보유자)\n');
fprintf('================================================================\n');

valid_idx = ~isnan(score_existing);
if sum(valid_idx) >= 3
    [corr_logistic_existing, p_corr_log_exist] = corr(score_logistic(valid_idx), ...
        score_existing(valid_idx), 'Type', 'Pearson');
    [corr_simple_existing, p_corr_simple_exist] = corr(score_simple(valid_idx), ...
        score_existing(valid_idx), 'Type', 'Pearson');
    [corr_logistic_simple, p_corr_log_simple] = corr(score_logistic(valid_idx), ...
        score_simple(valid_idx), 'Type', 'Pearson');

    % Spearman 순위 상관
    [spearman_log_exist, p_spear_log_exist] = corr(score_logistic(valid_idx), ...
        score_existing(valid_idx), 'Type', 'Spearman');
    [spearman_simple_exist, p_spear_simple_exist] = corr(score_simple(valid_idx), ...
        score_existing(valid_idx), 'Type', 'Spearman');

    fprintf('  【Pearson 상관계수】 (%d명)\n', sum(valid_idx));
    fprintf('    • Logistic vs 기존 종합점수: r = %.3f (p = %.4f)\n', ...
        corr_logistic_existing, p_corr_log_exist);
    fprintf('    • 단순합 vs 기존 종합점수: r = %.3f (p = %.4f)\n', ...
        corr_simple_existing, p_corr_simple_exist);
    fprintf('    • Logistic vs 단순합: r = %.3f (p = %.4f)\n', ...
        corr_logistic_simple, p_corr_log_simple);

    fprintf('\n  【Spearman 순위 상관】\n');
    fprintf('    • Logistic vs 기존 종합점수: ρ = %.3f (p = %.4f)\n', ...
        spearman_log_exist, p_spear_log_exist);
    fprintf('    • 단순합 vs 기존 종합점수: ρ = %.3f (p = %.4f)\n', ...
        spearman_simple_exist, p_spear_simple_exist);
else
    corr_logistic_existing = NaN;
    corr_simple_existing = NaN;
    corr_logistic_simple = NaN;
    spearman_log_exist = NaN;
    spearman_simple_exist = NaN;
end

%% 11) 기존 학습 데이터 검증 (참고용) ------------------------------------
fprintf('\n【STEP 11】 기존 학습 데이터 검증\n');
fprintf('================================================================\n');

try
    existing_data = readtable(config.existing_score_file, 'Sheet', '역량검사_종합점수', ...
        'VariableNamingRule', 'preserve');

    fprintf('  ✓ 기존 데이터: %d명\n', height(existing_data));

    % 인재유형 레이블 생성
    desired_types = {'성실한 가연성', '자연성', '유익한 불연성'};
    undesired_types = {'게으른 가연성', '무능한 불연성', '소화성'};
    excluded_types = {'유능한 불연성', '위장형 소화성'};

    valid_idx = true(height(existing_data), 1);
    for i = 1:length(excluded_types)
        valid_idx = valid_idx & ~strcmp(existing_data.('인재유형'), excluded_types{i});
    end
    existing_filtered = existing_data(valid_idx, :);

    labels_existing = zeros(height(existing_filtered), 1);
    for i = 1:height(existing_filtered)
        talent_type = existing_filtered.('인재유형'){i};
        if ismember(talent_type, desired_types)
            labels_existing(i) = 1;
        end
    end

    % 기존 데이터에서 원점수 계산 (단순합)
    score_existing_simple = existing_filtered.('총점');

    % ROC-AUC (참고용)
    [~, ~, ~, AUC_existing_logistic] = perfcurve(labels_existing, existing_filtered.('총점'), 1);

    fprintf('  ✓ 분석 대상: %d명 (제외 후)\n', height(existing_filtered));
    fprintf('  ✓ 뽑고 싶은 사람: %d명\n', sum(labels_existing==1));
    fprintf('  ✓ ROC-AUC: %.4f\n', AUC_existing_logistic);
catch ME
    fprintf('  ⚠ 기존 데이터 로드 실패: %s\n', ME.message);
    AUC_existing_logistic = NaN;
end

%% 12) 시각화 -----------------------------------------------------------
fprintf('\n【STEP 12】 시각화 생성\n');
fprintf('================================================================\n');

fig = figure('Position', [100, 100, 1800, 1200]);

% 1. ROC Curve
subplot(3, 3, 1);
plot(X_roc_logistic, Y_roc_logistic, 'r-', 'LineWidth', 2); hold on;
plot(X_roc_simple, Y_roc_simple, 'b-', 'LineWidth', 2);
if ~isempty(X_roc_existing)
    plot(X_roc_existing, Y_roc_existing, 'g--', 'LineWidth', 1.5);
end
plot([0, 1], [0, 1], 'k--');
xlabel('False Positive Rate');
ylabel('True Positive Rate');
title('ROC Curve');
if ~isempty(X_roc_existing)
    legend({sprintf('Logistic (AUC=%.3f)', AUC_logistic), ...
        sprintf('단순합 (AUC=%.3f)', AUC_simple), ...
        sprintf('기존 종합 (AUC=%.3f)', AUC_existing)}, 'Location', 'southeast');
else
    legend({sprintf('Logistic (AUC=%.3f)', AUC_logistic), ...
        sprintf('단순합 (AUC=%.3f)', AUC_simple)}, 'Location', 'southeast');
end
grid on;

% 2. PR Curve
subplot(3, 3, 2);
plot(X_pr_logistic, Y_pr_logistic, 'r-', 'LineWidth', 2); hold on;
plot(X_pr_simple, Y_pr_simple, 'b-', 'LineWidth', 2);
if ~isempty(X_pr_existing)
    plot(X_pr_existing, Y_pr_existing, 'g--', 'LineWidth', 1.5);
end
xlabel('Recall');
ylabel('Precision');
title('Precision-Recall Curve');
legend({sprintf('Logistic (AUC=%.3f)', AUC_pr_logistic), ...
    sprintf('단순합 (AUC=%.3f)', AUC_pr_simple)}, 'Location', 'southwest');
grid on;

% 3. 합격/불합격 분포 (Logistic)
subplot(3, 3, 3);
histogram(score_logistic(labels==1), 10, 'FaceColor', 'g', 'FaceAlpha', 0.6, 'EdgeColor', 'none'); hold on;
histogram(score_logistic(labels==0), 10, 'FaceColor', 'r', 'FaceAlpha', 0.6, 'EdgeColor', 'none');
xlabel('Logistic 가중치 점수');
ylabel('빈도');
title('점수 분포: Logistic');
legend({'합격', '불합격'}, 'Location', 'best');
grid on;

% 4. 합격/불합격 분포 (단순합)
subplot(3, 3, 4);
histogram(score_simple(labels==1), 10, 'FaceColor', 'g', 'FaceAlpha', 0.6, 'EdgeColor', 'none'); hold on;
histogram(score_simple(labels==0), 10, 'FaceColor', 'r', 'FaceAlpha', 0.6, 'EdgeColor', 'none');
xlabel('단순합 점수');
ylabel('빈도');
title('점수 분포: 단순합');
legend({'합격', '불합격'}, 'Location', 'best');
grid on;

% 5. Top-K Precision
subplot(3, 3, 5);
bar(topk_results.K, [topk_results.('Logistic_정밀도'), topk_results.('단순합_정밀도')]);
xlabel('선발 인원 (K)');
ylabel('합격 정밀도 (%)');
title('Top-K Precision');
legend({'Logistic', '단순합'}, 'Location', 'best');
grid on;

% 6. Box plot 비교
subplot(3, 3, 6);
data_boxplot = [score_logistic(labels==1); score_logistic(labels==0); ...
    score_simple(labels==1); score_simple(labels==0)];
group_boxplot = [repmat({'Logistic-합격'}, sum(labels==1), 1); ...
    repmat({'Logistic-불합격'}, sum(labels==0), 1); ...
    repmat({'단순합-합격'}, sum(labels==1), 1); ...
    repmat({'단순합-불합격'}, sum(labels==0), 1)];
boxplot(data_boxplot, group_boxplot);
ylabel('점수');
title('Box Plot 비교');
grid on;

% 7. 상관관계 산점도 (Logistic vs 기존 종합점수)
subplot(3, 3, 7);
if sum(~isnan(score_existing)) >= 3
    valid_idx = ~isnan(score_existing);
    scatter(score_logistic(valid_idx), score_existing(valid_idx), 50, labels(valid_idx), 'filled');
    xlabel('Logistic 가중치 점수');
    ylabel('기존 종합점수');
    title(sprintf('Logistic vs 기존 종합 (r=%.3f)', corr_logistic_existing));
    colormap([1 0 0; 0 1 0]);
    colorbar('Ticks', [0.25, 0.75], 'TickLabels', {'불합격', '합격'});
    grid on;
else
    text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
    axis off;
end

% 8. 상관관계 산점도 (단순합 vs 기존 종합점수)
subplot(3, 3, 8);
if sum(~isnan(score_existing)) >= 3
    valid_idx = ~isnan(score_existing);
    scatter(score_simple(valid_idx), score_existing(valid_idx), 50, labels(valid_idx), 'filled');
    xlabel('단순합 점수');
    ylabel('기존 종합점수');
    title(sprintf('단순합 vs 기존 종합 (r=%.3f)', corr_simple_existing));
    colormap([1 0 0; 0 1 0]);
    colorbar('Ticks', [0.25, 0.75], 'TickLabels', {'불합격', '합격'});
    grid on;
else
    text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center');
    axis off;
end

% 9. 개선도 요약
subplot(3, 3, 9);
axis off;
text(0.1, 0.9, '【성능 개선 요약】', 'FontSize', 12, 'FontWeight', 'bold');
text(0.1, 0.75, sprintf('ROC-AUC 개선: %.1f%%', (AUC_logistic/AUC_simple - 1)*100), 'FontSize', 10);
text(0.1, 0.65, sprintf('PR-AUC 개선: %.1f%%', (AUC_pr_logistic/AUC_pr_simple - 1)*100), 'FontSize', 10);
text(0.1, 0.55, sprintf('Cohen''s d: %.2f → %.2f', cohens_d_simple, cohens_d_logistic), 'FontSize', 10);
text(0.1, 0.45, sprintf('Top-5 개선: %.1f%%p', topk_results.('개선도_퍼센트포인트')(topk_results.K==5)), 'FontSize', 10);
text(0.1, 0.3, '【결론】', 'FontSize', 11, 'FontWeight', 'bold');
if AUC_logistic > AUC_simple
    text(0.1, 0.15, '✓ Logistic 가중치가 더 우수', 'FontSize', 10, 'Color', 'g');
else
    text(0.1, 0.15, '✗ 단순합이 더 우수', 'FontSize', 10, 'Color', 'r');
end

% 그래프 저장
plot_path = fullfile(config.output_dir, config.plot_filename);
saveas(fig, plot_path);
fprintf('  ✓ 그래프 저장: %s\n', config.plot_filename);

%% 13) 엑셀 결과 저장 ---------------------------------------------------
fprintf('\n【STEP 13】 엑셀 결과 저장\n');
fprintf('================================================================\n');

output_path = fullfile(config.output_dir, config.output_filename);

% 시트 1: 개인별 점수
result_individual = table();
result_individual.ID = comp_data.ID;
result_individual.('합불여부') = cell(n_new, 1);
for i = 1:n_new
    if labels(i) == 1
        result_individual.('합불여부'){i} = '합격';
    else
        result_individual.('합불여부'){i} = '불합격';
    end
end
result_individual.('Logistic점수') = round(score_logistic, 2);
result_individual.('단순합점수') = round(score_simple, 2);
result_individual.('기존종합점수') = round(score_existing, 2);
[~, rank_logistic] = sort(score_logistic, 'descend');
[~, rank_simple] = sort(score_simple, 'descend');
result_individual.('Logistic순위') = zeros(n_new, 1);
result_individual.('단순합순위') = zeros(n_new, 1);
result_individual.('Logistic순위')(rank_logistic) = (1:n_new)';
result_individual.('단순합순위')(rank_simple) = (1:n_new)';

writetable(result_individual, output_path, 'Sheet', '개인별점수', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 개인별점수\n');

% 시트 2: 성능 지표
performance = table();
performance.('지표') = {'ROC_AUC_Logistic'; 'ROC_AUC_단순합'; 'ROC_AUC_기존종합'; ...
    'PR_AUC_Logistic'; 'PR_AUC_단순합'; 'PR_AUC_기존종합'; ...
    'Top5_Logistic'; 'Top5_단순합'; 'Top5_개선도'; ...
    'Cohen_d_Logistic'; 'Cohen_d_단순합'; ...
    't_test_p_Logistic'; 't_test_p_단순합'};
performance.('값') = [AUC_logistic; AUC_simple; AUC_existing; ...
    AUC_pr_logistic; AUC_pr_simple; AUC_pr_existing; ...
    topk_results.('Logistic_정밀도')(topk_results.K==5); ...
    topk_results.('단순합_정밀도')(topk_results.K==5); ...
    topk_results.('개선도_퍼센트포인트')(topk_results.K==5); ...
    cohens_d_logistic; cohens_d_simple; ...
    p_logistic; p_simple];

writetable(performance, output_path, 'Sheet', '성능지표', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 성능지표\n');

% 시트 3: Top-K 결과
writetable(topk_results, output_path, 'Sheet', 'TopK결과', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: TopK결과\n');

% 시트 4: 상관관계
correlation = table();
correlation.('비교') = {'Logistic_vs_기존종합'; '단순합_vs_기존종합'; 'Logistic_vs_단순합'};
correlation.('Pearson_r') = [corr_logistic_existing; corr_simple_existing; corr_logistic_simple];
correlation.('Spearman_rho') = [spearman_log_exist; spearman_simple_exist; NaN];
correlation.('샘플수') = [sum(~isnan(score_existing)); sum(~isnan(score_existing)); sum(~isnan(score_existing))];

writetable(correlation, output_path, 'Sheet', '상관관계', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 상관관계\n');

fprintf('\n  ✅ 전체 파일 저장 완료: %s\n', output_path);

%% 14) 종합 리포트 ------------------------------------------------------
fprintf('\n');
fprintf('================================================================\n');
fprintf('                    종합 검증 리포트\n');
fprintf('================================================================\n\n');

fprintf('【 데이터 구성 】\n');
fprintf('  • 전체 인원: %d명\n', n_new);
fprintf('  • 합격: %d명 (%.1f%%)\n', pass_count, pass_count/n_new*100);
fprintf('  • 불합격: %d명 (%.1f%%)\n', fail_count, fail_count/n_new*100);
fprintf('  • 제외: 1명 (조기 이탈)\n\n');

fprintf('【 점수 평균 비교 】\n');
fprintf('  Logistic 가중치:\n');
fprintf('    - 합격: %.2f ± %.2f\n', mean(score_logistic(labels==1)), std(score_logistic(labels==1)));
fprintf('    - 불합격: %.2f ± %.2f\n', mean(score_logistic(labels==0)), std(score_logistic(labels==0)));
fprintf('  단순합:\n');
fprintf('    - 합격: %.2f ± %.2f\n', mean(score_simple(labels==1)), std(score_simple(labels==1)));
fprintf('    - 불합격: %.2f ± %.2f\n\n', mean(score_simple(labels==0)), std(score_simple(labels==0)));

fprintf('【 예측 성능 】\n');
fprintf('  1. ROC-AUC\n');
fprintf('     • Logistic: %.4f\n', AUC_logistic);
fprintf('     • 단순합: %.4f\n', AUC_simple);
fprintf('     • 개선도: %.1f%%\n\n', (AUC_logistic/AUC_simple - 1)*100);

fprintf('  2. Top-5 선발 시\n');
fprintf('     • Logistic: %.0f%% (%d/5명)\n', topk_results.('Logistic_정밀도')(topk_results.K==5), ...
    round(topk_results.('Logistic_정밀도')(topk_results.K==5)/100*5));
fprintf('     • 단순합: %.0f%% (%d/5명)\n', topk_results.('단순합_정밀도')(topk_results.K==5), ...
    round(topk_results.('단순합_정밀도')(topk_results.K==5)/100*5));
fprintf('     • 개선: %.1f%%p\n\n', topk_results.('개선도_퍼센트포인트')(topk_results.K==5));

fprintf('【 통계적 유의성 】\n');
fprintf('  • Logistic: p = %.4f, Cohen''s d = %.3f', p_logistic, cohens_d_logistic);
if abs(cohens_d_logistic) > 0.8
    fprintf(' (Large)\n');
elseif abs(cohens_d_logistic) > 0.5
    fprintf(' (Medium)\n');
else
    fprintf(' (Small)\n');
end
fprintf('  • 단순합: p = %.4f, Cohen''s d = %.3f', p_simple, cohens_d_simple);
if abs(cohens_d_simple) > 0.8
    fprintf(' (Large)\n');
elseif abs(cohens_d_simple) > 0.5
    fprintf(' (Medium)\n');
else
    fprintf(' (Small)\n');
end

if ~isnan(cohens_d_existing)
    fprintf('  • 기존 종합점수: p = %.4f, Cohen''s d = %.3f', p_existing, cohens_d_existing);
    if abs(cohens_d_existing) > 0.8
        fprintf(' (Large)\n');
    elseif abs(cohens_d_existing) > 0.5
        fprintf(' (Medium)\n');
    else
        fprintf(' (Small)\n');
    end
end

fprintf('\n【 결론 】\n');

% 세 가지 점수 비교
scores_comparison = [AUC_logistic, AUC_simple];
methods_name = {'Logistic', '단순합'};

if ~isnan(AUC_existing) && sum(~isnan(score_existing)) >= 5
    scores_comparison(3) = AUC_existing;
    methods_name{3} = '기존 종합점수';
end

[best_auc, best_idx] = max(scores_comparison);
best_method = methods_name{best_idx};

fprintf('  📊 ROC-AUC 성능 순위:\n');
[sorted_auc, sorted_idx] = sort(scores_comparison, 'descend');
for i = 1:length(sorted_auc)
    fprintf('    %d. %s: %.4f\n', i, methods_name{sorted_idx(i)}, sorted_auc(i));
end

fprintf('\n  ✅ 최고 성능: %s (AUC = %.4f)\n', best_method, best_auc);

if strcmp(best_method, 'Logistic')
    fprintf('  ✅ Logistic 가중치가 가장 우수한 성능을 보였습니다.\n');
elseif strcmp(best_method, '단순합')
    fprintf('  ⚠ 단순합이 가장 우수한 성능을 보였습니다.\n');
    fprintf('     (샘플 수가 적어 추가 검증 필요)\n');
else
    fprintf('  ⚠ 기존 종합점수가 가장 우수한 성능을 보였습니다.\n');
    fprintf('     (신규 가중치 재검토 필요)\n');
end

% Cohen's d 비교
if ~isnan(cohens_d_existing)
    fprintf('\n  📊 효과 크기 (Cohen''s d) 비교:\n');
    fprintf('    • Logistic: %.3f\n', abs(cohens_d_logistic));
    fprintf('    • 단순합: %.3f\n', abs(cohens_d_simple));
    fprintf('    • 기존 종합점수: %.3f\n', abs(cohens_d_existing));
end

fprintf('\n================================================================\n');
fprintf('  검증 완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('================================================================\n\n');
