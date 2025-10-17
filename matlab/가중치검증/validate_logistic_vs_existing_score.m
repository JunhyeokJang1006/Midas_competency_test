% =======================================================================
%     신규 입사자 가중치 검증: Logistic 가중치 vs 기존 종합점수
% =======================================================================
% 목적:
%   - 신규 입사자에 대해 Logistic 가중치 점수가 합격/불합격을
%     얼마나 잘 예측하는지 검증
%   - 기존 종합점수 대비 개선도 측정
%   - 실무 담당자용 마크다운 리포트 생성
%
% 데이터:
%   - 신규 입사자: 18명 (합격 11명, 불합격 7명)
%   - 기존 학습 데이터: 130명
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
fprintf('   신규 입사자 가중치 검증: Logistic 가중치 vs 기존 종합점수\n');
fprintf('================================================================\n\n');

%% 1) 설정 ---------------------------------------------------------------
fprintf('【STEP 1】 설정\n');
fprintf('================================================================\n');

config = struct();
config.new_comp_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_역검 점수.xlsx';
config.new_onboarding_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_온보딩 점수.xlsx';
config.weight_file = 'D:\project\HR데이터\결과\자가불소_revised_talent\integrated_analysis_results.mat';
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_revised.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가_filtered.xlsx';
config.output_dir = 'D:\project\HR데이터\결과\가중치검증';
config.timestamp = datestr(now, 'yyyymmdd_HHMMSS');
config.output_filename = sprintf('validate_logistic_vs_existing_%s.xlsx', config.timestamp);
config.plot_filename = sprintf('validation_plot_%s.png', config.timestamp);
config.dist_plot_filename = sprintf('distribution_comparison_%s.png', config.timestamp);
config.report_filename = sprintf('가중치검증_실무리포트_%s.md', config.timestamp);

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

dist_plot_path = fullfile(config.output_dir, config.dist_plot_filename);
saveas(fig_dist, dist_plot_path);
fprintf('    ✓ 분포 그래프 저장: %s\n', config.dist_plot_filename);

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
    distribution_validity = '적합';
else
    fprintf('\n    ⚠ 일부 역량에서 분포 차이가 발견되었습니다.\n');
    fprintf('      → 가중치 적용 시 주의가 필요합니다.\n');
    distribution_validity = '주의필요';
end

%% 4) 신규 입사자 데이터 로드 및 레이블 생성 -------------------------------
fprintf('\n【STEP 4】 신규 입사자 데이터 로드 및 레이블 생성\n');
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

% 레이블 생성 및 필터링
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

%% 5) 신규 입사자 점수 계산 -----------------------------------------------
fprintf('\n【STEP 5】 신규 입사자 점수 계산\n');
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

% 2) 기존 종합점수
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

valid_existing_count = sum(~isnan(score_existing));
fprintf('\n  【기존 종합점수】 (%d명)\n', valid_existing_count);
if valid_existing_count > 0
    fprintf('    • 합격 평균: %.2f ± %.2f\n', mean(score_existing(labels==1), 'omitnan'), std(score_existing(labels==1), 'omitnan'));
    fprintf('    • 불합격 평균: %.2f ± %.2f\n', mean(score_existing(labels==0), 'omitnan'), std(score_existing(labels==0), 'omitnan'));

    % 기존 종합점수와 Logistic 점수 간 차이 분석
    fprintf('\n  【점수 차이 분석】 (%d명)\n', valid_existing_count);
    valid_idx = ~isnan(score_existing);
    diff_logistic = score_logistic(valid_idx) - score_existing(valid_idx);

    fprintf('    • Logistic 가중치 - 기존 종합점수:\n');
    fprintf('      평균 차이: %.2f (SD: %.2f)\n', mean(diff_logistic), std(diff_logistic));
    fprintf('      범위: %.2f ~ %.2f\n', min(diff_logistic), max(diff_logistic));

    % 차이의 방향성 분석
    higher_count = sum(diff_logistic > 0);
    lower_count = sum(diff_logistic < 0);
    same_count = sum(diff_logistic == 0);
    fprintf('      Logistic이 더 높음: %d명 (%.1f%%)\n', higher_count, higher_count/valid_existing_count*100);
    fprintf('      기존이 더 높음: %d명 (%.1f%%)\n', lower_count, lower_count/valid_existing_count*100);
else
    fprintf('    • 기존 종합점수 데이터 없음\n');
end

%% 6) ROC-AUC 및 PR-AUC 분석 ---------------------------------------------
fprintf('\n【STEP 6】 ROC-AUC 및 PR-AUC 분석\n');
fprintf('================================================================\n');

% ROC 곡선
[X_roc_logistic, Y_roc_logistic, ~, AUC_logistic] = perfcurve(labels, score_logistic, 1);

fprintf('  【Logistic 가중치 ROC-AUC】\n');
fprintf('    • AUC: %.4f\n', AUC_logistic);

% PR 곡선
[X_pr_logistic, Y_pr_logistic, ~, AUC_pr_logistic] = perfcurve(labels, score_logistic, 1, ...
    'XCrit', 'reca', 'YCrit', 'prec');

fprintf('\n  【Logistic 가중치 PR-AUC】\n');
fprintf('    • AUC: %.4f\n', AUC_pr_logistic);

% 기존 종합점수
valid_existing = ~isnan(score_existing);
if sum(valid_existing) >= 5
    [X_roc_existing, Y_roc_existing, ~, AUC_existing] = perfcurve(labels(valid_existing), ...
        score_existing(valid_existing), 1);
    [X_pr_existing, Y_pr_existing, ~, AUC_pr_existing] = perfcurve(labels(valid_existing), ...
        score_existing(valid_existing), 1, 'XCrit', 'reca', 'YCrit', 'prec');

    fprintf('\n  【기존 종합점수 ROC-AUC】 (%d명)\n', sum(valid_existing));
    fprintf('    • AUC: %.4f\n', AUC_existing);
    fprintf('    • Logistic 개선도: %.4f (%.1f%%)\n', AUC_logistic - AUC_existing, ...
        (AUC_logistic/AUC_existing - 1)*100);
else
    AUC_existing = NaN;
    AUC_pr_existing = NaN;
    X_roc_existing = [];
    Y_roc_existing = [];
    X_pr_existing = [];
    Y_pr_existing = [];
end

%% 7) Top-K Precision 분석 ----------------------------------------------
fprintf('\n【STEP 7】 Top-K Precision 분석\n');
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

    % 기존 종합점수 (데이터가 있는 경우)
    if sum(valid_existing) >= k
        [~, idx_existing] = sort(score_existing(valid_existing), 'descend');
        valid_labels = labels(valid_existing);
        top_k_existing = valid_labels(idx_existing(1:k));
        precision_existing = sum(top_k_existing) / k * 100;
        improvement = precision_logistic - precision_existing;
    else
        precision_existing = NaN;
        improvement = NaN;
    end

    fprintf('  【상위 %d명 선발】\n', k);
    fprintf('    • Logistic: %.1f%% (%d/%d명 합격)\n', precision_logistic, sum(top_k_logistic), k);
    if ~isnan(precision_existing)
        fprintf('    • 기존 종합점수: %.1f%% (%d/%d명 합격)\n', precision_existing, sum(top_k_existing), k);
        fprintf('    • 개선: %.1f%%p\n\n', improvement);
    else
        fprintf('    • 기존 종합점수: 데이터 부족\n\n');
        improvement = NaN;
    end

    topk_results = [topk_results; table(k, precision_logistic, precision_existing, improvement, ...
        'VariableNames', {'K', 'Logistic_정밀도', '기존종합점수_정밀도', '개선도_퍼센트포인트'})]; %#ok<AGROW>
end

%% 8) 통계적 검정 -------------------------------------------------------
fprintf('\n【STEP 8】 통계적 검정 (t-test, Cohen''s d)\n');
fprintf('================================================================\n');

% t-test: 합격 vs 불합격
[~, p_logistic, ~, ~] = ttest2(score_logistic(labels==1), score_logistic(labels==0));

% Cohen's d (효과 크기)
mean_pass_logistic = mean(score_logistic(labels==1));
mean_fail_logistic = mean(score_logistic(labels==0));
std_pooled_logistic = sqrt(((sum(labels==1)-1)*var(score_logistic(labels==1)) + ...
    (sum(labels==0)-1)*var(score_logistic(labels==0))) / (n_new - 2));
cohens_d_logistic = (mean_pass_logistic - mean_fail_logistic) / std_pooled_logistic;

fprintf('  【Logistic 가중치 점수】\n');
fprintf('    • t-test p-value: %.4f\n', p_logistic);
fprintf('    • Cohen''s d: %.3f', cohens_d_logistic);
if abs(cohens_d_logistic) > 0.8
    fprintf(' (Large effect)\n');
    effect_size_logistic = 'Large';
elseif abs(cohens_d_logistic) > 0.5
    fprintf(' (Medium effect)\n');
    effect_size_logistic = 'Medium';
else
    fprintf(' (Small effect)\n');
    effect_size_logistic = 'Small';
end

% 기존 종합점수 통계 검정
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
        effect_size_existing = 'Large';
    elseif abs(cohens_d_existing) > 0.5
        fprintf(' (Medium effect)\n');
        effect_size_existing = 'Medium';
    else
        fprintf(' (Small effect)\n');
        effect_size_existing = 'Small';
    end

    fprintf('\n  【효과 크기 비교】\n');
    if abs(cohens_d_logistic) > abs(cohens_d_existing)
        effect_ratio = abs(cohens_d_logistic) / abs(cohens_d_existing);
        fprintf('    • ✅ Logistic 가중치가 기존 종합점수보다 %.1f배 더 큰 효과 (차이: %.3f)\n', ...
            effect_ratio, abs(cohens_d_logistic) - abs(cohens_d_existing));
    else
        effect_ratio = abs(cohens_d_existing) / abs(cohens_d_logistic);
        fprintf('    • ⚠ 기존 종합점수가 Logistic 가중치보다 %.1f배 더 큰 효과 (차이: %.3f)\n', ...
            effect_ratio, abs(cohens_d_existing) - abs(cohens_d_logistic));
    end
else
    p_existing = NaN;
    cohens_d_existing = NaN;
    effect_size_existing = 'N/A';
end

%% 9) 상관관계 분석 ----------------------------------------------------
fprintf('\n【STEP 9】 상관관계 분석\n');
fprintf('================================================================\n');

valid_idx = ~isnan(score_existing);
if sum(valid_idx) >= 3
    [corr_logistic_existing, p_corr_log_exist] = corr(score_logistic(valid_idx), ...
        score_existing(valid_idx), 'Type', 'Pearson');

    % Spearman 순위 상관
    [spearman_log_exist, p_spear_log_exist] = corr(score_logistic(valid_idx), ...
        score_existing(valid_idx), 'Type', 'Spearman');

    fprintf('  【상관계수】 (%d명)\n', sum(valid_idx));
    fprintf('    • Pearson: r = %.3f (p = %.4f)\n', ...
        corr_logistic_existing, p_corr_log_exist);
    fprintf('    • Spearman: ρ = %.3f (p = %.4f)\n', ...
        spearman_log_exist, p_spear_log_exist);

    if corr_logistic_existing > 0.8
        fprintf('    • ✅ 매우 높은 상관관계 (기존 시스템과 일치)\n');
    elseif corr_logistic_existing > 0.6
        fprintf('    • ✅ 높은 상관관계\n');
    elseif corr_logistic_existing > 0.4
        fprintf('    • ℹ 중간 상관관계\n');
    else
        fprintf('    • ⚠ 낮은 상관관계 (기존 시스템과 상이)\n');
    end
else
    corr_logistic_existing = NaN;
    spearman_log_exist = NaN;
end

%% 10) 시각화 -----------------------------------------------------------
fprintf('\n【STEP 10】 시각화 생성\n');
fprintf('================================================================\n');

fig = figure('Position', [100, 100, 1600, 1000]);

% 1. ROC Curve
subplot(2, 3, 1);
plot(X_roc_logistic, Y_roc_logistic, 'r-', 'LineWidth', 2.5); hold on;
if ~isempty(X_roc_existing)
    plot(X_roc_existing, Y_roc_existing, 'b--', 'LineWidth', 2);
end
plot([0, 1], [0, 1], 'k--', 'LineWidth', 1);
xlabel('False Positive Rate', 'FontWeight', 'bold');
ylabel('True Positive Rate', 'FontWeight', 'bold');
title('ROC Curve', 'FontSize', 12, 'FontWeight', 'bold');
if ~isempty(X_roc_existing)
    legend({sprintf('Logistic (AUC=%.3f)', AUC_logistic), ...
        sprintf('기존 종합점수 (AUC=%.3f)', AUC_existing)}, 'Location', 'southeast');
else
    legend({sprintf('Logistic (AUC=%.3f)', AUC_logistic)}, 'Location', 'southeast');
end
grid on;

% 2. PR Curve
subplot(2, 3, 2);
plot(X_pr_logistic, Y_pr_logistic, 'r-', 'LineWidth', 2.5); hold on;
if ~isempty(X_pr_existing)
    plot(X_pr_existing, Y_pr_existing, 'b--', 'LineWidth', 2);
end
xlabel('Recall', 'FontWeight', 'bold');
ylabel('Precision', 'FontWeight', 'bold');
title('Precision-Recall Curve', 'FontSize', 12, 'FontWeight', 'bold');
legend({sprintf('Logistic (AUC=%.3f)', AUC_pr_logistic)}, 'Location', 'southwest');
grid on;

% 3. 합격/불합격 분포 (Logistic)
subplot(2, 3, 3);
histogram(score_logistic(labels==1), 10, 'FaceColor', [0.2 0.8 0.4], 'FaceAlpha', 0.7, 'EdgeColor', 'k'); hold on;
histogram(score_logistic(labels==0), 10, 'FaceColor', [0.9 0.3 0.3], 'FaceAlpha', 0.7, 'EdgeColor', 'k');
xlabel('Logistic 가중치 점수', 'FontWeight', 'bold');
ylabel('빈도', 'FontWeight', 'bold');
title('점수 분포: Logistic 가중치', 'FontSize', 12, 'FontWeight', 'bold');
legend({'합격', '불합격'}, 'Location', 'best');
grid on;

% 4. Top-K Precision 비교
subplot(2, 3, 4);
valid_topk = topk_results(~isnan(topk_results.('기존종합점수_정밀도')), :);
if ~isempty(valid_topk)
    bar(valid_topk.K, [valid_topk.('Logistic_정밀도'), valid_topk.('기존종합점수_정밀도')]);
    xlabel('선발 인원 (K)', 'FontWeight', 'bold');
    ylabel('합격 정밀도 (%)', 'FontWeight', 'bold');
    title('Top-K Precision 비교', 'FontSize', 12, 'FontWeight', 'bold');
    legend({'Logistic', '기존 종합점수'}, 'Location', 'best');
    grid on;
else
    text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
    title('Top-K Precision 비교');
    axis off;
end

% 5. 상관관계 산점도
subplot(2, 3, 5);
if sum(~isnan(score_existing)) >= 3
    valid_idx = ~isnan(score_existing);
    scatter(score_logistic(valid_idx), score_existing(valid_idx), 100, labels(valid_idx), 'filled', 'MarkerEdgeColor', 'k');
    hold on;
    % 추세선 추가
    p_fit = polyfit(score_logistic(valid_idx), score_existing(valid_idx), 1);
    x_fit = linspace(min(score_logistic(valid_idx)), max(score_logistic(valid_idx)), 100);
    y_fit = polyval(p_fit, x_fit);
    plot(x_fit, y_fit, 'k--', 'LineWidth', 2);
    xlabel('Logistic 가중치 점수', 'FontWeight', 'bold');
    ylabel('기존 종합점수', 'FontWeight', 'bold');
    title(sprintf('상관관계 (r=%.3f)', corr_logistic_existing), 'FontSize', 12, 'FontWeight', 'bold');
    colormap([0.9 0.3 0.3; 0.2 0.8 0.4]);
    cb = colorbar('Ticks', [0.25, 0.75], 'TickLabels', {'불합격', '합격'});
    cb.Label.String = '합불 여부';
    grid on;
else
    text(0.5, 0.5, '데이터 부족', 'HorizontalAlignment', 'center', 'FontSize', 12);
    title('상관관계 분석');
    axis off;
end

% 6. 개선도 요약
subplot(2, 3, 6);
axis off;
text(0.05, 0.95, '【검증 결과 요약】', 'FontSize', 13, 'FontWeight', 'bold');
y_pos = 0.85;
if ~isnan(AUC_existing)
    text(0.05, y_pos, sprintf('ROC-AUC 개선: %.1f%%', (AUC_logistic/AUC_existing - 1)*100), 'FontSize', 10);
    y_pos = y_pos - 0.1;
end
text(0.05, y_pos, sprintf('Cohen''s d (Logistic): %.3f', abs(cohens_d_logistic)), 'FontSize', 10);
y_pos = y_pos - 0.1;
if ~isnan(cohens_d_existing)
    text(0.05, y_pos, sprintf('Cohen''s d (기존): %.3f', abs(cohens_d_existing)), 'FontSize', 10);
    y_pos = y_pos - 0.1;
end
if sum(valid_idx) >= 3
    text(0.05, y_pos, sprintf('상관계수: %.3f', corr_logistic_existing), 'FontSize', 10);
    y_pos = y_pos - 0.1;
end

y_pos = y_pos - 0.05;
text(0.05, y_pos, '【결론】', 'FontSize', 12, 'FontWeight', 'bold');
y_pos = y_pos - 0.1;

if abs(cohens_d_logistic) > abs(cohens_d_existing)
    text(0.05, y_pos, '✓ Logistic 가중치 우수', 'FontSize', 11, 'Color', [0.2 0.8 0.4], 'FontWeight', 'bold');
else
    text(0.05, y_pos, '△ 기존 종합점수 유사', 'FontSize', 11, 'Color', [0.9 0.5 0]);
end

% 그래프 저장
plot_path = fullfile(config.output_dir, config.plot_filename);
saveas(fig, plot_path);
fprintf('  ✓ 그래프 저장: %s\n', config.plot_filename);

%% 11) 엑셀 결과 저장 ---------------------------------------------------
fprintf('\n【STEP 11】 엑셀 결과 저장\n');
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
result_individual.('기존종합점수') = round(score_existing, 2);
result_individual.('점수차이') = round(score_logistic - score_existing, 2);
[~, rank_logistic] = sort(score_logistic, 'descend');
result_individual.('Logistic순위') = zeros(n_new, 1);
result_individual.('Logistic순위')(rank_logistic) = (1:n_new)';

writetable(result_individual, output_path, 'Sheet', '개인별점수', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 개인별점수\n');

% 시트 2: 성능 지표
performance = table();
performance.('지표') = {'ROC_AUC_Logistic'; 'ROC_AUC_기존종합'; 'ROC_AUC_개선도_퍼센트'; ...
    'PR_AUC_Logistic'; 'PR_AUC_기존종합'; ...
    'Top5_Logistic'; 'Top5_기존종합'; 'Top5_개선도'; ...
    'Cohen_d_Logistic'; 'Cohen_d_기존종합'; ...
    't_test_p_Logistic'; 't_test_p_기존종합'; ...
    'Pearson_r'; 'Spearman_rho'};

improvement_pct = NaN;
if ~isnan(AUC_existing)
    improvement_pct = (AUC_logistic/AUC_existing - 1) * 100;
end

top5_log = topk_results.('Logistic_정밀도')(topk_results.K==5);
top5_exist = topk_results.('기존종합점수_정밀도')(topk_results.K==5);
top5_improve = topk_results.('개선도_퍼센트포인트')(topk_results.K==5);

performance.('값') = [AUC_logistic; AUC_existing; improvement_pct; ...
    AUC_pr_logistic; AUC_pr_existing; ...
    top5_log; top5_exist; top5_improve; ...
    cohens_d_logistic; cohens_d_existing; ...
    p_logistic; p_existing; ...
    corr_logistic_existing; spearman_log_exist];

writetable(performance, output_path, 'Sheet', '성능지표', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 성능지표\n');

% 시트 3: Top-K 결과
writetable(topk_results, output_path, 'Sheet', 'TopK결과', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: TopK결과\n');

% 시트 4: 분포 비교
writetable(distribution_comparison, output_path, 'Sheet', '분포비교', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 분포비교\n');

fprintf('\n  ✅ 전체 파일 저장 완료: %s\n', output_path);

%% 12) 실무 담당자용 마크다운 리포트 생성 ---------------------------------
fprintf('\n【STEP 12】 실무 담당자용 리포트 생성\n');
fprintf('================================================================\n');

report_path = fullfile(config.output_dir, config.report_filename);
fid = fopen(report_path, 'w', 'n', 'UTF-8');

fprintf(fid, '# 역량검사 가중치 검증 리포트\n\n');
fprintf(fid, '**작성일**: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM'));
fprintf(fid, '**검증 대상**: 2025년 하반기 신규 입사자 (18명)\n\n');
fprintf(fid, '---\n\n');

% 1. 요약
fprintf(fid, '## 📊 핵심 요약 (Executive Summary)\n\n');

if abs(cohens_d_logistic) > abs(cohens_d_existing)
    fprintf(fid, '### ✅ **새로운 Logistic 가중치가 기존 종합점수보다 우수한 성능을 보임**\n\n');
    fprintf(fid, '- **효과 크기**: Logistic 가중치(Cohen''s d = %.3f)가 기존 종합점수(%.3f)보다 **%.0f%% 더 큰 효과**\n', ...
        abs(cohens_d_logistic), abs(cohens_d_existing), ...
        (abs(cohens_d_logistic)/abs(cohens_d_existing) - 1)*100);
else
    fprintf(fid, '### ℹ **새로운 Logistic 가중치가 기존 종합점수와 유사한 성능을 보임**\n\n');
    fprintf(fid, '- **효과 크기**: Logistic 가중치(Cohen''s d = %.3f)와 기존 종합점수(%.3f)가 유사\n', ...
        abs(cohens_d_logistic), abs(cohens_d_existing));
end

if ~isnan(AUC_existing)
    fprintf(fid, '- **ROC-AUC 개선도**: %.1f%%\n', (AUC_logistic/AUC_existing - 1)*100);
end

if corr_logistic_existing > 0.8
    fprintf(fid, '- **기존 시스템과의 일관성**: 매우 높음 (r = %.3f) ✅\n', corr_logistic_existing);
elseif corr_logistic_existing > 0.6
    fprintf(fid, '- **기존 시스템과의 일관성**: 높음 (r = %.3f)\n', corr_logistic_existing);
else
    fprintf(fid, '- **기존 시스템과의 일관성**: 중간~낮음 (r = %.3f) ⚠️\n', corr_logistic_existing);
end

fprintf(fid, '- **데이터 분포 적합성**: %s\n\n', distribution_validity);

% 2. 검증 개요
fprintf(fid, '---\n\n');
fprintf(fid, '## 1. 검증 개요\n\n');
fprintf(fid, '### 1.1 목적\n\n');
fprintf(fid, '- 기존 역량검사 종합점수 대비 **새로운 Logistic 가중치**의 예측 성능 검증\n');
fprintf(fid, '- 신규 입사자의 합격/불합격 예측 정확도 측정\n');
fprintf(fid, '- 실무 적용 가능성 평가\n\n');

fprintf(fid, '### 1.2 데이터\n\n');
fprintf(fid, '| 구분 | 인원 | 비율 |\n');
fprintf(fid, '|------|------|------|\n');
fprintf(fid, '| **합격** | %d명 | %.1f%% |\n', pass_count, pass_count/n_new*100);
fprintf(fid, '| **불합격** | %d명 | %.1f%% |\n', fail_count, fail_count/n_new*100);
fprintf(fid, '| **합계** | %d명 | 100%% |\n\n', n_new);

fprintf(fid, '> **참고**: 조기 이탈자 1명은 분석에서 제외\n\n');

% 3. 주요 결과
fprintf(fid, '---\n\n');
fprintf(fid, '## 2. 주요 결과\n\n');

fprintf(fid, '### 2.1 예측 성능 비교\n\n');
fprintf(fid, '#### 📈 ROC-AUC (높을수록 우수)\n\n');
fprintf(fid, '| 방법 | AUC | 개선도 |\n');
fprintf(fid, '|------|-----|--------|\n');
fprintf(fid, '| **Logistic 가중치** | **%.4f** | - |\n', AUC_logistic);
if ~isnan(AUC_existing)
    fprintf(fid, '| 기존 종합점수 | %.4f | %.1f%% |\n\n', AUC_existing, (AUC_logistic/AUC_existing - 1)*100);
else
    fprintf(fid, '| 기존 종합점수 | N/A | - |\n\n');
end

fprintf(fid, '> **해석**: ROC-AUC는 합격자와 불합격자를 구분하는 능력을 나타냅니다.\n');
fprintf(fid, '> - 0.5 = 무작위 추측\n');
fprintf(fid, '> - 0.7-0.8 = 양호\n');
fprintf(fid, '> - 0.8-0.9 = 우수\n');
fprintf(fid, '> - 0.9-1.0 = 매우 우수\n\n');

fprintf(fid, '#### 🎯 Top-5 선발 정밀도 (실무 시나리오)\n\n');
fprintf(fid, '상위 5명을 선발했을 때 실제 합격자 비율:\n\n');
fprintf(fid, '| 방법 | 합격률 | 개선도 |\n');
fprintf(fid, '|------|--------|--------|\n');
fprintf(fid, '| **Logistic 가중치** | **%.0f%%** (%d/5명) | - |\n', top5_log, round(top5_log/100*5));
if ~isnan(top5_exist)
    fprintf(fid, '| 기존 종합점수 | %.0f%% (%d/5명) | %+.0f%%p |\n\n', ...
        top5_exist, round(top5_exist/100*5), top5_improve);
else
    fprintf(fid, '| 기존 종합점수 | N/A | - |\n\n');
end

fprintf(fid, '### 2.2 통계적 유의성\n\n');
fprintf(fid, '#### 효과 크기 (Cohen''s d)\n\n');
fprintf(fid, '합격자와 불합격자 점수 차이의 크기:\n\n');
fprintf(fid, '| 방법 | Cohen''s d | 효과 크기 | 평가 |\n');
fprintf(fid, '|------|------------|-----------|------|\n');
fprintf(fid, '| **Logistic 가중치** | **%.3f** | %s | ', abs(cohens_d_logistic), effect_size_logistic);
if abs(cohens_d_logistic) > abs(cohens_d_existing)
    fprintf(fid, '✅ **우수** |\n');
else
    fprintf(fid, '△ 유사 |\n');
end

if ~isnan(cohens_d_existing)
    fprintf(fid, '| 기존 종합점수 | %.3f | %s | - |\n\n', abs(cohens_d_existing), effect_size_existing);
else
    fprintf(fid, '| 기존 종합점수 | N/A | N/A | - |\n\n');
end

fprintf(fid, '> **해석**: Cohen''s d는 두 그룹 간 차이의 실질적 크기를 나타냅니다.\n');
fprintf(fid, '> - 0.2 = Small (작은 효과)\n');
fprintf(fid, '> - 0.5 = Medium (중간 효과)\n');
fprintf(fid, '> - 0.8 = Large (큰 효과)\n\n');

fprintf(fid, '### 2.3 기존 시스템과의 일관성\n\n');
if ~isnan(corr_logistic_existing)
    fprintf(fid, '- **Pearson 상관계수**: r = %.3f (p < %.4f)\n', corr_logistic_existing, p_corr_log_exist);
    fprintf(fid, '- **Spearman 순위 상관**: ρ = %.3f (p < %.4f)\n\n', spearman_log_exist, p_spear_log_exist);

    if corr_logistic_existing > 0.8
        fprintf(fid, '✅ **매우 높은 상관관계**: 새로운 가중치가 기존 시스템과 잘 일치합니다.\n\n');
    elseif corr_logistic_existing > 0.6
        fprintf(fid, '✅ **높은 상관관계**: 새로운 가중치가 기존 시스템과 대체로 일치합니다.\n\n');
    else
        fprintf(fid, '⚠️ **중간~낮은 상관관계**: 새로운 가중치가 기존 시스템과 다른 관점을 제공합니다.\n\n');
    end
else
    fprintf(fid, 'N/A (데이터 부족)\n\n');
end

% 4. 기존 데이터 분포 적합성
fprintf(fid, '---\n\n');
fprintf(fid, '## 3. 기존 학습 데이터와의 분포 비교\n\n');
fprintf(fid, '### 3.1 분포 유사성 검증\n\n');
fprintf(fid, '| 지표 | 결과 |\n');
fprintf(fid, '|------|------|\n');
fprintf(fid, '| 기존 학습 데이터 | %d명 |\n', size(X_train, 1));
fprintf(fid, '| 신규 입사자 데이터 | %d명 |\n', size(X_new_comp, 1));
fprintf(fid, '| 통계적 유의한 차이 (p<0.05) | %d/%d개 역량 |\n', sig_diff, n_comps);
fprintf(fid, '| Large effect (丨d丨>0.8) | %d개 |\n', large_effect);
fprintf(fid, '| Medium effect (0.5<丨d丨≤0.8) | %d개 |\n\n', medium_effect);

if sig_diff == 0
    fprintf(fid, '✅ **결론**: 기존 학습 데이터와 신규 데이터의 분포가 **유사**합니다.\n');
    fprintf(fid, '→ 학습된 가중치를 신규 데이터에 적용하는 것이 **타당**합니다.\n\n');
else
    fprintf(fid, '⚠️ **결론**: 일부 역량에서 분포 차이가 발견되었습니다.\n');
    fprintf(fid, '→ 가중치 적용 시 **주의**가 필요합니다.\n\n');
end

% 5. 실무 적용 제언
fprintf(fid, '---\n\n');
fprintf(fid, '## 4. 실무 적용 제언\n\n');

fprintf(fid, '### 4.1 도입 권장 사항\n\n');

if abs(cohens_d_logistic) > abs(cohens_d_existing) && corr_logistic_existing > 0.7
    fprintf(fid, '#### ✅ **적극 도입 권장**\n\n');
    fprintf(fid, '1. **우수한 예측 성능**: 기존 종합점수 대비 더 높은 효과 크기\n');
    fprintf(fid, '2. **높은 일관성**: 기존 시스템과 높은 상관관계 유지\n');
    fprintf(fid, '3. **분포 적합성**: 신규 데이터에 적용 타당성 확인\n\n');
    fprintf(fid, '**권장 방안**:\n');
    fprintf(fid, '- 기존 종합점수를 Logistic 가중치 점수로 **전면 교체**\n');
    fprintf(fid, '- 또는 두 점수를 **병행** 사용하여 의사결정 보완\n\n');
elseif abs(cohens_d_logistic) > abs(cohens_d_existing)
    fprintf(fid, '#### ℹ️ **신중한 도입 권장**\n\n');
    fprintf(fid, '1. **개선된 성능**: 기존 대비 더 높은 효과 크기\n');
    fprintf(fid, '2. **일부 차이**: 기존 시스템과 중간 수준 상관관계\n\n');
    fprintf(fid, '**권장 방안**:\n');
    fprintf(fid, '- 두 점수를 **병행** 사용하여 의사결정 보완\n');
    fprintf(fid, '- 추가 데이터로 **재검증** 후 전면 도입 고려\n\n');
else
    fprintf(fid, '#### ℹ️ **보완적 활용 권장**\n\n');
    fprintf(fid, '1. **유사한 성능**: 기존 종합점수와 비슷한 수준\n');
    fprintf(fid, '2. **추가 관점 제공**: 다른 각도에서 인재 평가 가능\n\n');
    fprintf(fid, '**권장 방안**:\n');
    fprintf(fid, '- 기존 종합점수를 **주요** 지표로 유지\n');
    fprintf(fid, '- Logistic 가중치를 **보조** 지표로 참고\n\n');
end

fprintf(fid, '### 4.2 주의 사항\n\n');
fprintf(fid, '- 분석 샘플 수가 18명으로 제한적이므로, **추가 데이터**로 재검증 권장\n');
fprintf(fid, '- 특정 인재 유형(조기 이탈 등)은 제외되었으므로, **전체 모집단** 대표성 고려 필요\n');
fprintf(fid, '- 역량검사 외 **다른 요소**(면접, 경력 등)도 채용 결정에 영향을 미칠 수 있음\n\n');

fprintf(fid, '### 4.3 후속 조치\n\n');
fprintf(fid, '1. **추가 검증**: 더 많은 신규 입사자 데이터로 재검증\n');
fprintf(fid, '2. **장기 추적**: 채용 후 실제 성과와의 관계 분석\n');
fprintf(fid, '3. **주기적 업데이트**: 분기별 또는 반기별 가중치 재조정\n\n');

% 6. 기술적 세부사항
fprintf(fid, '---\n\n');
fprintf(fid, '## 5. 기술적 세부사항\n\n');
fprintf(fid, '### 5.1 분석 방법\n\n');
fprintf(fid, '- **ROC-AUC**: Receiver Operating Characteristic - Area Under Curve\n');
fprintf(fid, '- **Cohen''s d**: 효과 크기 측정 (표준화된 평균 차이)\n');
fprintf(fid, '- **t-test**: 두 그룹 간 평균 차이의 통계적 유의성 검정\n');
fprintf(fid, '- **Pearson/Spearman 상관계수**: 두 변수 간 선형/순위 관계 측정\n\n');

fprintf(fid, '### 5.2 사용된 역량 (%d개)\n\n', length(feature_names));
for i = 1:length(feature_names)
    fprintf(fid, '%d. %s (가중치: %.2f%%)\n', i, char(feature_names(i)), logistic_weights(i));
end
fprintf(fid, '\n');

% 7. 부록
fprintf(fid, '---\n\n');
fprintf(fid, '## 6. 부록\n\n');
fprintf(fid, '### 6.1 출력 파일\n\n');
fprintf(fid, '1. **엑셀 결과**: `%s`\n', config.output_filename);
fprintf(fid, '   - 시트1: 개인별 점수 (ID, 합불여부, Logistic점수, 기존종합점수 등)\n');
fprintf(fid, '   - 시트2: 성능 지표 (ROC-AUC, Cohen''s d, 상관계수 등)\n');
fprintf(fid, '   - 시트3: Top-K 결과\n');
fprintf(fid, '   - 시트4: 분포 비교\n\n');

fprintf(fid, '2. **시각화**: `%s`\n', config.plot_filename);
fprintf(fid, '   - ROC Curve, PR Curve, 점수 분포, Top-K Precision 등\n\n');

fprintf(fid, '3. **분포 비교 그래프**: `%s`\n\n', config.dist_plot_filename);

fprintf(fid, '### 6.2 문의\n\n');
fprintf(fid, '분석 결과에 대한 문의사항이 있으시면 담당자에게 연락 바랍니다.\n\n');

fprintf(fid, '---\n\n');
fprintf(fid, '*본 리포트는 Claude Code를 통해 자동 생성되었습니다.*\n');
fprintf(fid, '*생성일시: %s*\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));


fprintf('  ✓ 마크다운 리포트 저장: %s\n', config.report_filename);

%% 13) 종합 리포트 (콘솔) -----------------------------------------------
fprintf('\n');
fprintf('================================================================\n');
fprintf('                    종합 검증 리포트\n');
fprintf('================================================================\n\n');

fprintf('【 데이터 구성 】\n');
fprintf('  • 최종 분석 대상: %d명\n', n_new);
fprintf('  • 합격: %d명 (%.1f%%)\n', pass_count, pass_count/n_new*100);
fprintf('  • 불합격: %d명 (%.1f%%)\n\n', fail_count, fail_count/n_new*100);

fprintf('【 예측 성능 】\n');
fprintf('  1. ROC-AUC\n');
fprintf('     • Logistic: %.4f\n', AUC_logistic);
if ~isnan(AUC_existing)
    fprintf('     • 기존 종합점수: %.4f\n', AUC_existing);
    fprintf('     • 개선도: %.1f%%\n\n', (AUC_logistic/AUC_existing - 1)*100);
else
    fprintf('     • 기존 종합점수: N/A\n\n');
end

fprintf('  2. Top-5 선발 시\n');
fprintf('     • Logistic: %.0f%% (%d/5명)\n', top5_log, round(top5_log/100*5));
if ~isnan(top5_exist)
    fprintf('     • 기존 종합점수: %.0f%% (%d/5명)\n', top5_exist, round(top5_exist/100*5));
    fprintf('     • 개선: %+.0f%%p\n\n', top5_improve);
else
    fprintf('     • 기존 종합점수: N/A\n\n');
end

fprintf('【 통계적 유의성 】\n');
fprintf('  • Logistic: Cohen''s d = %.3f (%s)\n', abs(cohens_d_logistic), effect_size_logistic);
if ~isnan(cohens_d_existing)
    fprintf('  • 기존 종합점수: Cohen''s d = %.3f (%s)\n\n', abs(cohens_d_existing), effect_size_existing);
else
    fprintf('  • 기존 종합점수: N/A\n\n');
end

fprintf('【 결론 】\n');
if abs(cohens_d_logistic) > abs(cohens_d_existing)
    effect_ratio = abs(cohens_d_logistic) / abs(cohens_d_existing);
    fprintf('  ✅ Logistic 가중치가 기존 종합점수보다 우수한 성능을 보였습니다.\n');
    fprintf('     효과 크기가 %.1f배 더 큽니다.\n\n', effect_ratio);
else
    fprintf('  ℹ️ Logistic 가중치가 기존 종합점수와 유사한 성능을 보였습니다.\n\n');
end

fprintf('================================================================\n');
fprintf('  검증 완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('  출력 디렉토리: %s\n', config.output_dir);
fprintf('================================================================\n\n');
