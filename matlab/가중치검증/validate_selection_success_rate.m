% =======================================================================
%     선발 성공률 검증: Logistic 가중치 vs 기존 종합점수
% =======================================================================
% 목적:
%   - 실무자 친화적인 "선발 성공률" 중심 검증
%   - 핵심 질문: "상위 몇 %를 뽑았을 때, 실제 합격자가 몇 명인가?"
%
% 데이터:
%   - 신규 입사자: 18명 (합격 11명, 불합격 7명)
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
fprintf('        선발 성공률 검증: Logistic 가중치 vs 기존 종합점수\n');
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
config.output_filename = sprintf('selection_success_rate_%s.xlsx', config.timestamp);
config.plot_filename = sprintf('success_rate_plot_%s.png', config.timestamp);
config.report_filename = sprintf('선발성공률_검증리포트_%s.md', config.timestamp);

% 선발 비율 설정 (실무 시나리오)
config.selection_ratios = [0.15, 0.20, 0.25, 0.33, 0.40, 0.50];  % 15%, 20%, 28%, ...

if ~exist(config.output_dir, 'dir')
    mkdir(config.output_dir);
end

fprintf('  ✓ 출력 디렉토리: %s\n', config.output_dir);
fprintf('  ✓ 선발 비율: %s\n', strjoin(arrayfun(@(x) sprintf('%.0f%%', x*100), ...
    config.selection_ratios, 'UniformOutput', false), ', '));

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

%% 3) 분포 비교 (선택적) -------------------------------------------------
fprintf('\n【STEP 3】 분포 비교 (선택적)\n');
fprintf('================================================================\n');

% 기존 학습 데이터 로드
hr_data = readtable(config.hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
comp_upper = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');

% 인재유형 필터링
talent_col_idx = find(contains(hr_data.Properties.VariableNames, '인재유형'), 1);
if ~isempty(talent_col_idx)
    talent_types = hr_data{:, talent_col_idx};
    exclude_types = {'위장형 소화성'};
    valid_talent_idx = ~ismember(talent_types, exclude_types);
    hr_data = hr_data(valid_talent_idx, :);
end

% ID 매칭
hr_ids = hr_data.ID;
comp_ids = comp_upper.ID;
[matched_ids_train, ~, comp_idx] = intersect(hr_ids, comp_ids);
matched_comp_train = comp_upper(comp_idx, :);

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

fprintf('  ✓ 기존 학습 데이터: %d명\n', size(X_train, 1));
fprintf('  ℹ️ 분포 비교는 부록에 포함됩니다\n');

%% 4) 신규 입사자 데이터 로드 및 실제 인원 출력 ---------------------------
fprintf('\n【STEP 4】 신규 입사자 데이터 로드\n');
fprintf('================================================================\n');

% 역량검사 데이터
comp_data = readtable(config.new_comp_file, 'Sheet', '역량검사_상위항목', ...
    'VariableNamingRule', 'preserve');
fprintf('  • 역량검사 데이터: %d명\n', height(comp_data));

% 기존 종합점수
score_data = readtable(config.new_comp_file, 'Sheet', '역량검사_종합점수', ...
    'VariableNamingRule', 'preserve');
fprintf('  • 기존 종합점수: %d명\n', height(score_data));

% 온보딩 데이터
onboarding_data = readtable(config.new_onboarding_file, 'VariableNamingRule', 'preserve');
fprintf('  • 온보딩 데이터: %d명\n', height(onboarding_data));

% ID 통일
if isnumeric(comp_data.ID)
    comp_data.ID = arrayfun(@(x) sprintf('%d', x), comp_data.ID, 'UniformOutput', false);
end
if isnumeric(score_data.ID)
    score_data.ID = arrayfun(@(x) sprintf('%d', x), score_data.ID, 'UniformOutput', false);
end
if isnumeric(onboarding_data.ID)
    onboarding_data.ID = arrayfun(@(x) sprintf('%d', x), onboarding_data.ID, 'UniformOutput', false);
end

% 합불 여부 컬럼 찾기
pass_fail_col = '합불 여부';
if ~ismember(pass_fail_col, onboarding_data.Properties.VariableNames)
    possible_names = onboarding_data.Properties.VariableNames(...
        contains(onboarding_data.Properties.VariableNames, '합불'));
    if ~isempty(possible_names)
        pass_fail_col = possible_names{1};
    end
end

% 레이블 생성
n_total = height(comp_data);
labels = nan(n_total, 1);
include_mask = true(n_total, 1);

for i = 1:n_total
    id = comp_data.ID{i};
    onb_idx = find(strcmp(onboarding_data.ID, id), 1);

    if ~isempty(onb_idx)
        pass_fail = onboarding_data.(pass_fail_col){onb_idx};

        if strcmp(pass_fail, '합격')
            labels(i) = 1;
        elseif strcmp(pass_fail, '불합격')
            labels(i) = 0;
        else
            include_mask(i) = false;
        end
    else
        include_mask(i) = false;
    end
end

% 필터링
comp_data = comp_data(include_mask, :);
labels = labels(include_mask);
n_new = sum(include_mask);
pass_count = sum(labels == 1);
fail_count = sum(labels == 0);
baseline_rate = pass_count / n_new * 100;

fprintf('\n【 데이터 개요 】\n');
fprintf('  • 분석 대상: %d명\n', n_new);
fprintf('  • 실제 합격자: %d명 (%.1f%%)\n', pass_count, baseline_rate);
fprintf('  • 실제 불합격자: %d명 (%.1f%%)\n', fail_count, fail_count/n_new*100);
fprintf('  • 랜덤 선발 기대 성공률: %.1f%%\n\n', baseline_rate);

%% 5) 점수 계산 ---------------------------------------------------------
fprintf('【STEP 5】 점수 계산\n');
fprintf('================================================================\n');

% Logistic 가중치 점수
X_new = [];
matched_weights = [];

for i = 1:length(feature_names)
    fn = char(feature_names(i));
    if ismember(fn, comp_data.Properties.VariableNames)
        col = comp_data.(fn);
        if ~isnumeric(col)
            col = double(col);
        end
        X_new = [X_new, col(:)]; %#ok<AGROW>
        matched_weights(end+1) = logistic_weights(i); %#ok<AGROW>
    end
end

w = matched_weights(:) / 100;
score_logistic = nansum(X_new .* repmat(w', n_new, 1), 2) ./ sum(w);

% 기존 종합점수
score_existing = nan(n_new, 1);
for i = 1:n_new
    id = comp_data.ID{i};
    score_idx = find(strcmp(score_data.ID, id), 1);
    if ~isempty(score_idx)
        score_existing(i) = score_data.('종합점수')(score_idx);
    end
end

fprintf('  ✓ Logistic 점수: %.1f ~ %.1f점\n', min(score_logistic), max(score_logistic));
fprintf('  ✓ 기존 종합점수: %.1f ~ %.1f점\n', min(score_existing, [], 'omitnan'), ...
    max(score_existing, [], 'omitnan'));

%% 6) 선발 성공률 분석 (⭐ 핵심!) -----------------------------------------
fprintf('\n【STEP 6】 선발 성공률 분석 (⭐ 핵심!)\n');
fprintf('================================================================\n');
fprintf('\n🎯 핵심 질문: "상위 몇 %%를 뽑았을 때, 실제 합격자가 몇 명인가?"\n\n');

selection_ratios = config.selection_ratios;
results = {};

for ratio = selection_ratios
    k = round(n_new * ratio);
    if k < 1, k = 1; end
    if k > n_new, k = n_new; end

    % Logistic 가중치
    [~, idx_log] = sort(score_logistic, 'descend');
    top_k_log = labels(idx_log(1:k));
    success_rate_log = sum(top_k_log) / k * 100;
    success_count_log = sum(top_k_log);

    % 기존 종합점수
    if sum(~isnan(score_existing)) >= k
        [~, idx_exist] = sort(score_existing, 'descend', 'MissingPlacement', 'last');
        top_k_exist = labels(idx_exist(1:k));
        success_rate_exist = sum(top_k_exist) / k * 100;
        success_count_exist = sum(top_k_exist);
        has_existing = true;
    else
        success_rate_exist = NaN;
        success_count_exist = NaN;
        has_existing = false;
    end

    % 저장
    result = struct();
    result.ratio = ratio * 100;
    result.k = k;
    result.ratio_str = sprintf('상위 %.0f%%', ratio * 100);
    result.k_str = sprintf('%d/%d명', k, n_new);
    result.success_log = success_rate_log;
    result.success_count_log = success_count_log;
    result.success_exist = success_rate_exist;
    result.success_count_exist = success_count_exist;
    result.improvement = success_rate_log - success_rate_exist;
    result.vs_random = success_rate_log - baseline_rate;
    result.has_existing = has_existing;

    results{end+1} = result; %#ok<AGROW>
end

% 표 형태 출력
fprintf('  ┌────────────┬──────────┬─────────────────────────────────────┐\n');
fprintf('  │ 선발 비율  │ 선발인원 │    선발 성공률 (%%)                 │\n');
fprintf('  │            │          │ Logistic │  기존   │  개선  │ 평가│\n');
fprintf('  ├────────────┼──────────┼──────────┼─────────┼────────┼────┤\n');

for i = 1:length(results)
    r = results{i};
    stars = '';
    if r.success_log >= 80
        stars = '⭐⭐⭐';
    elseif r.success_log >= 70
        stars = '⭐⭐';
    elseif r.success_log >= 60
        stars = '⭐';
    end

    if r.has_existing
        fprintf('  │ %9s  │ %8s │  %5.1f   │  %5.1f  │%+6.1f%%p│ %s │\n', ...
            r.ratio_str, r.k_str, r.success_log, r.success_exist, ...
            r.improvement, stars);
    else
        fprintf('  │ %9s  │ %8s │  %5.1f   │   N/A   │   N/A  │ %s │\n', ...
            r.ratio_str, r.k_str, r.success_log, stars);
    end
end
fprintf('  └────────────┴──────────┴──────────┴─────────┴────────┴────┘\n\n');

fprintf('💡 해석:\n');
fprintf('  • 선발 성공률 = (선발 인원 중 실제 합격자) ÷ (선발 인원) × 100%%\n');
fprintf('  • 평가 기준:\n');
fprintf('    - 100%%: 완벽 (선발한 사람 모두 합격자) ⭐⭐⭐\n');
fprintf('    - 80~99%%: 매우 우수 ⭐⭐⭐\n');
fprintf('    - 70~79%%: 우수 ⭐⭐\n');
fprintf('    - 60~69%%: 보통 ⭐\n\n');

% 대표 시나리오 상세 설명
ratios_for_rep = cellfun(@(x) x.ratio, results);
rep_idx = find(ratios_for_rep == 28, 1);
if isempty(rep_idx)
    rep_idx = min(3, length(results));
end
rep_result = results{rep_idx};

fprintf('【예시】%s (%s) 선발 시\n', rep_result.ratio_str, rep_result.k_str);
fprintf('  ✅ Logistic 가중치: %d명 중 %d명 합격 (%.1f%% 성공)\n', ...
    rep_result.k, rep_result.success_count_log, rep_result.success_log);

if rep_result.has_existing
    fprintf('  ⚠️ 기존 종합점수: %d명 중 %d명 합격 (%.1f%% 성공)\n', ...
        rep_result.k, rep_result.success_count_exist, rep_result.success_exist);
    fprintf('  📈 개선 효과: %d명 더 정확하게 선발 (%+.1f%%p 향상)\n\n', ...
        rep_result.success_count_log - rep_result.success_count_exist, rep_result.improvement);
else
    fprintf('  ⚠️ 기존 종합점수: 데이터 없음\n\n');
end

%% 7) ROC-AUC 및 Cohen's d (참고용) ---------------------------------------
fprintf('【STEP 7】 참고 지표 (ROC-AUC, Cohen''s d)\n');
fprintf('================================================================\n');

% ROC-AUC
[~, ~, ~, AUC_logistic] = perfcurve(labels, score_logistic, 1);

valid_existing = ~isnan(score_existing);
if sum(valid_existing) >= 5
    [~, ~, ~, AUC_existing] = perfcurve(labels(valid_existing), ...
        score_existing(valid_existing), 1);
else
    AUC_existing = NaN;
end

fprintf('\n  【ROC-AUC】 (전체 예측 능력 지표)\n');
fprintf('    • Logistic 가중치: %.4f\n', AUC_logistic);
if ~isnan(AUC_existing)
    fprintf('    • 기존 종합점수: %.4f\n', AUC_existing);
    fprintf('    • 개선도: %.1f%%\n', (AUC_logistic/AUC_existing - 1)*100);
else
    fprintf('    • 기존 종합점수: N/A\n');
end

% Cohen's d
mean_pass = mean(score_logistic(labels==1));
mean_fail = mean(score_logistic(labels==0));
std_pooled = sqrt(((sum(labels==1)-1)*var(score_logistic(labels==1)) + ...
    (sum(labels==0)-1)*var(score_logistic(labels==0))) / (n_new - 2));
cohens_d_log = (mean_pass - mean_fail) / std_pooled;

fprintf('\n  【Cohen''s d】 (효과 크기)\n');
fprintf('    • Logistic 가중치: %.3f', abs(cohens_d_log));
if abs(cohens_d_log) > 0.8
    fprintf(' (Large effect)\n');
elseif abs(cohens_d_log) > 0.5
    fprintf(' (Medium effect)\n');
else
    fprintf(' (Small effect)\n');
end

if sum(valid_existing) >= 5 && sum(labels(valid_existing)==1) >= 2 && sum(labels(valid_existing)==0) >= 2
    mean_pass_exist = mean(score_existing(labels==1 & valid_existing), 'omitnan');
    mean_fail_exist = mean(score_existing(labels==0 & valid_existing), 'omitnan');
    n_pass = sum(labels==1 & valid_existing);
    n_fail = sum(labels==0 & valid_existing);
    std_pooled_exist = sqrt(((n_pass-1)*var(score_existing(labels==1 & valid_existing), 'omitnan') + ...
        (n_fail-1)*var(score_existing(labels==0 & valid_existing), 'omitnan')) / (n_pass + n_fail - 2));
    cohens_d_exist = (mean_pass_exist - mean_fail_exist) / std_pooled_exist;

    fprintf('    • 기존 종합점수: %.3f', abs(cohens_d_exist));
    if abs(cohens_d_exist) > 0.8
        fprintf(' (Large effect)\n');
    elseif abs(cohens_d_exist) > 0.5
        fprintf(' (Medium effect)\n');
    else
        fprintf(' (Small effect)\n');
    end
else
    cohens_d_exist = NaN;
end

fprintf('\n  ℹ️ 이 지표들은 참고용이며, 실무에서는 "선발 성공률"에 집중하세요.\n');

%% 8) 시각화 (단순화) ---------------------------------------------------
fprintf('\n【STEP 8】 시각화 생성\n');
fprintf('================================================================\n');

fig = figure('Position', [100, 100, 1400, 900]);

ratios_arr = [results{:}];
success_log_arr = [ratios_arr.success_log];
success_exist_arr = [ratios_arr.success_exist];
improvement_arr = [ratios_arr.improvement];

% 1. 선발 성공률 비교 (막대 그래프)
subplot(2, 2, 1);
if results{1}.has_existing
    b = bar(1:length(results), [success_log_arr', success_exist_arr']);
    b(1).FaceColor = [0.2 0.6 0.9];
    b(2).FaceColor = [0.9 0.5 0.2];
    legend({'Logistic 가중치', '기존 종합점수'}, 'Location', 'best');
else
    bar(1:length(results), success_log_arr');
    legend({'Logistic 가중치'}, 'Location', 'best');
end

xlabel('선발 비율', 'FontWeight', 'bold');
ylabel('선발 성공률 (%)', 'FontWeight', 'bold');
title('선발 비율별 성공률 비교', 'FontSize', 13, 'FontWeight', 'bold');
xticks(1:length(results));
xticklabels({ratios_arr.ratio_str});
xtickangle(45);
grid on;
ylim([0 110]);

% 2. 개선 효과 (막대 그래프 + 텍스트)
subplot(2, 2, 2);
b = bar(1:length(results), improvement_arr');
b.FaceColor = 'flat';
for i = 1:length(improvement_arr)
    if improvement_arr(i) > 5
        b.CData(i,:) = [0.2 0.8 0.4];  % 개선 - 초록
    elseif improvement_arr(i) < -5
        b.CData(i,:) = [0.9 0.3 0.3];  % 악화 - 빨강
    else
        b.CData(i,:) = [0.7 0.7 0.7];  % 차이 없음 - 회색
    end

    % 텍스트 표시
    if abs(improvement_arr(i)) < 0.1
        text(i, improvement_arr(i), '개선없음', ...
            'HorizontalAlignment', 'center', 'VerticalAlignment', 'bottom', ...
            'FontSize', 8, 'FontWeight', 'bold');
    end
end

xlabel('선발 비율', 'FontWeight', 'bold');
ylabel('개선 효과 (%p)', 'FontWeight', 'bold');
title('Logistic 가중치 개선 효과', 'FontSize', 13, 'FontWeight', 'bold');
xticks(1:length(results));
xticklabels({ratios_arr.ratio_str});
xtickangle(45);
yline(0, 'k--', 'LineWidth', 1);
grid on;

% 3. 선발 인원 구성 비교 (기존 vs Logistic)
subplot(2, 2, 3);
if results{1}.has_existing
    success_log_counts = cellfun(@(x) x.success_count_log, results);
    fail_log_counts = arrayfun(@(i) results{i}.k - results{i}.success_count_log, 1:length(results));
    success_exist_counts = cellfun(@(x) x.success_count_exist, results);
    fail_exist_counts = arrayfun(@(i) results{i}.k - results{i}.success_count_exist, 1:length(results));

    x_pos = 1:length(results);
    bar_width = 0.35;

    % Logistic
    b1 = bar(x_pos - bar_width/2, [success_log_counts', fail_log_counts'], 'stacked', 'BarWidth', bar_width);
    hold on;
    % 기존
    b2 = bar(x_pos + bar_width/2, [success_exist_counts', fail_exist_counts'], 'stacked', 'BarWidth', bar_width);

    b1(1).FaceColor = [0.2 0.7 0.9];  % Logistic 합격 - 파랑
    b1(2).FaceColor = [0.9 0.4 0.4];  % Logistic 불합격 - 빨강
    b2(1).FaceColor = [0.9 0.6 0.2];  % 기존 합격 - 주황
    b2(2).FaceColor = [0.7 0.3 0.3];  % 기존 불합격 - 어두운 빨강

    legend({'Logistic 합격', 'Logistic 오판', '기존 합격', '기존 오판'}, ...
        'Location', 'best', 'FontSize', 8);

    xlabel('선발 비율', 'FontWeight', 'bold');
    ylabel('선발 인원 (명)', 'FontWeight', 'bold');
    title('선발 인원 구성 비교 (Logistic vs 기존)', 'FontSize', 13, 'FontWeight', 'bold');
    xticks(1:length(results));
    xticklabels({ratios_arr.ratio_str});
    xtickangle(45);
    grid on;
else
    % 기존 데이터가 없으면 Logistic만 표시
    success_counts = cellfun(@(x) x.success_count_log, results);
    fail_counts = arrayfun(@(i) results{i}.k - results{i}.success_count_log, 1:length(results));

    b = bar(1:length(results), [success_counts', fail_counts'], 'stacked');
    b(1).FaceColor = [0.2 0.8 0.4];
    b(2).FaceColor = [0.9 0.3 0.3];

    legend({'합격자', '불합격자'}, 'Location', 'best');
    xlabel('선발 비율', 'FontWeight', 'bold');
    ylabel('선발 인원 (명)', 'FontWeight', 'bold');
    title('Logistic 선발 인원 구성', 'FontSize', 13, 'FontWeight', 'bold');
    xticks(1:length(results));
    xticklabels({ratios_arr.ratio_str});
    xtickangle(45);
    grid on;
end

% 4. 종합 요약
subplot(2, 2, 4);
axis off;
text(0.05, 0.95, '【선발 성공률 검증 요약】', 'FontSize', 14, 'FontWeight', 'bold');
y_pos = 0.85;

text(0.05, y_pos, sprintf('분석 대상: %d명', n_new), 'FontSize', 10);
y_pos = y_pos - 0.07;
text(0.05, y_pos, sprintf('실제 합격: %d명 (%.1f%%)', pass_count, baseline_rate), 'FontSize', 10);
y_pos = y_pos - 0.07;
text(0.05, y_pos, sprintf('실제 불합격: %d명 (%.1f%%)', fail_count, fail_count/n_new*100), 'FontSize', 10);
y_pos = y_pos - 0.12;

text(0.05, y_pos, sprintf('【대표: %s】', rep_result.ratio_str), 'FontSize', 11, 'FontWeight', 'bold');
y_pos = y_pos - 0.07;
text(0.05, y_pos, sprintf('• Logistic: %.1f%% (%d/%d명)', ...
    rep_result.success_log, rep_result.success_count_log, rep_result.k), 'FontSize', 10);
y_pos = y_pos - 0.07;

if rep_result.has_existing
    text(0.05, y_pos, sprintf('• 기존: %.1f%% (%d/%d명)', ...
        rep_result.success_exist, rep_result.success_count_exist, rep_result.k), 'FontSize', 10);
    y_pos = y_pos - 0.07;
    text(0.05, y_pos, sprintf('• 개선: %+.1f%%p', rep_result.improvement), ...
        'FontSize', 10, 'Color', [0.2 0.7 0.3], 'FontWeight', 'bold');
    y_pos = y_pos - 0.12;
else
    y_pos = y_pos - 0.12;
end

text(0.05, y_pos, '【결론】', 'FontSize', 12, 'FontWeight', 'bold', 'Color', [0 0.4 0.8]);
y_pos = y_pos - 0.08;

if rep_result.has_existing && rep_result.improvement > 5
    text(0.05, y_pos, '✅ Logistic 가중치', 'FontSize', 11, 'Color', [0.2 0.7 0.3], 'FontWeight', 'bold');
    y_pos = y_pos - 0.07;
    text(0.05, y_pos, '   우수한 선발 성공률!', 'FontSize', 10, 'Color', [0.2 0.7 0.3]);
else
    text(0.05, y_pos, 'ℹ️ 두 방법 유사', 'FontSize', 11);
end

sgtitle('선발 성공률 검증 결과', 'FontSize', 16, 'FontWeight', 'bold');

% 그래프는 백업 후 저장 (STEP 9에서 저장됨)

%% 9) 기존 파일 백업 및 Excel 저장 -----------------------------------------
fprintf('\n【STEP 9】 기존 파일 백업 및 Excel 결과 저장\n');
fprintf('================================================================\n');

% 백업 폴더 생성
backup_dir = fullfile(config.output_dir, 'backup');
if ~exist(backup_dir, 'dir')
    mkdir(backup_dir);
end

% 기존 파일들을 백업 폴더로 이동
fprintf('\n  【기존 파일 백업】\n');

% 1. Excel 파일 백업
excel_files = dir(fullfile(config.output_dir, 'selection_success_rate_*.xlsx'));
if ~isempty(excel_files)
    fprintf('    • Excel 파일 %d개 발견\n', length(excel_files));
    for i = 1:length(excel_files)
        old_path = fullfile(excel_files(i).folder, excel_files(i).name);
        new_path = fullfile(backup_dir, excel_files(i).name);
        movefile(old_path, new_path);
    end
    fprintf('    ✓ Excel 파일 백업 완료\n');
else
    fprintf('    • Excel 파일 없음\n');
end

% 2. 그래프 파일 백업
plot_files = dir(fullfile(config.output_dir, 'success_rate_plot_*.png'));
if ~isempty(plot_files)
    fprintf('    • 그래프 파일 %d개 발견\n', length(plot_files));
    for i = 1:length(plot_files)
        old_path = fullfile(plot_files(i).folder, plot_files(i).name);
        new_path = fullfile(backup_dir, plot_files(i).name);
        movefile(old_path, new_path);
    end
    fprintf('    ✓ 그래프 파일 백업 완료\n');
else
    fprintf('    • 그래프 파일 없음\n');
end

% 3. 리포트 파일 백업
report_files = dir(fullfile(config.output_dir, '선발성공률_검증리포트_*.md'));
if ~isempty(report_files)
    fprintf('    • 리포트 파일 %d개 발견\n', length(report_files));
    for i = 1:length(report_files)
        old_path = fullfile(report_files(i).folder, report_files(i).name);
        new_path = fullfile(backup_dir, report_files(i).name);
        movefile(old_path, new_path);
    end
    fprintf('    ✓ 리포트 파일 백업 완료\n');
else
    fprintf('    • 리포트 파일 없음\n');
end

fprintf('\n  【새 파일 생성】\n');

% 그래프 저장 (백업 완료 후 저장)
plot_path = fullfile(config.output_dir, config.plot_filename);
saveas(fig, plot_path);
fprintf('    ✓ 그래프 파일: %s\n', config.plot_filename);

% Excel 파일 경로
output_path = fullfile(config.output_dir, config.output_filename);

% 시트 1: 선발 비율별 성공률 (메인)
summary_table = table();
for i = 1:length(results)
    r = results{i};
    row = table();
    row.('선발비율') = {r.ratio_str};
    row.('선발인원') = {r.k_str};
    row.('Logistic_성공률') = r.success_log;
    row.('Logistic_성공인원') = sprintf('%d/%d', r.success_count_log, r.k);

    if r.has_existing
        row.('기존_성공률') = r.success_exist;
        row.('기존_성공인원') = sprintf('%d/%d', r.success_count_exist, r.k);
        row.('개선효과_퍼센트포인트') = r.improvement;
    else
        row.('기존_성공률') = NaN;
        row.('기존_성공인원') = {'N/A'};
        row.('개선효과_퍼센트포인트') = NaN;
    end

    summary_table = [summary_table; row]; %#ok<AGROW>
end

writetable(summary_table, output_path, 'Sheet', '선발성공률_요약', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 선발성공률_요약\n');

% 시트 2: 개인별 점수 + 순위 + 인재유형 유추
individual_table = table();
individual_table.ID = comp_data.ID;
individual_table.('실제_합불') = arrayfun(@(x) iif(x==1, '합격', '불합격'), labels, 'UniformOutput', false);
individual_table.('Logistic_점수') = round(score_logistic, 2);

[~, rank_log] = sort(score_logistic, 'descend');
individual_table.('Logistic_순위') = zeros(n_new, 1);
individual_table.('Logistic_순위')(rank_log) = (1:n_new)';

individual_table.('기존_점수') = round(score_existing, 2);

if sum(~isnan(score_existing)) >= n_new
    [~, rank_exist] = sort(score_existing, 'descend');
    individual_table.('기존_순위') = zeros(n_new, 1);
    individual_table.('기존_순위')(rank_exist) = (1:n_new)';
end

% 인재유형 (유추) - 온보딩 데이터에서 가져오기
talent_inferred = cell(n_new, 1);

% 온보딩 데이터에서 '인재유형 (유추)' 컬럼 찾기
talent_col_name = '';
onb_var_names = onboarding_data.Properties.VariableNames;
for v = 1:length(onb_var_names)
    if contains(onb_var_names{v}, '인재유형')
        talent_col_name = onb_var_names{v};
        break;
    end
end

if ~isempty(talent_col_name)
    % ID 매칭해서 인재유형 가져오기
    for i = 1:n_new
        id = comp_data.ID{i};
        onb_idx = find(strcmp(onboarding_data.ID, id), 1);

        if ~isempty(onb_idx)
            talent_value = onboarding_data.(talent_col_name){onb_idx};
            if ~isempty(talent_value)
                talent_inferred{i} = talent_value;
            else
                talent_inferred{i} = '정보없음';
            end
        else
            talent_inferred{i} = '정보없음';
        end
    end
else
    % 인재유형 컬럼이 없으면 기본값
    for i = 1:n_new
        talent_inferred{i} = '정보없음';
    end
end

individual_table.('인재유형_유추') = talent_inferred;

writetable(individual_table, output_path, 'Sheet', '개인별_점수_순위', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 개인별_점수_순위 (인재유형 유추 포함)\n');

% 시트 3: 참고 지표
reference_table = table();
reference_table.('지표명') = {'ROC_AUC_Logistic'; 'ROC_AUC_기존'; 'ROC_AUC_개선도_퍼센트'; ...
    'Cohen_d_Logistic'; 'Cohen_d_기존'; '랜덤_선발_기대_성공률'; '실제_합격자수'; '전체_인원'};

if ~isnan(AUC_existing)
    auc_improve = (AUC_logistic/AUC_existing - 1) * 100;
else
    auc_improve = NaN;
end

reference_table.('값') = [AUC_logistic; AUC_existing; auc_improve; ...
    cohens_d_log; cohens_d_exist; baseline_rate; pass_count; n_new];

writetable(reference_table, output_path, 'Sheet', '참고지표', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 참고지표\n');

% 시트 4: 선발 비율별 상세 분석
detail_table = table();
for i = 1:length(results)
    r = results{i};

    row = table();
    row.('선발비율') = {r.ratio_str};
    row.('선발인원_K') = r.k;
    row.('전체인원_N') = n_new;

    % Logistic 상세
    row.('Logistic_성공률') = r.success_log;
    row.('Logistic_합격자수_TP') = r.success_count_log;
    row.('Logistic_오판수_FP') = r.k - r.success_count_log;
    row.('Logistic_정확도') = r.success_log;

    if r.has_existing
        % 기존 종합점수 상세
        row.('기존_성공률') = r.success_exist;
        row.('기존_합격자수_TP') = r.success_count_exist;
        row.('기존_오판수_FP') = r.k - r.success_count_exist;
        row.('기존_정확도') = r.success_exist;

        % 비교
        row.('성공률_차이') = r.improvement;
        row.('합격자수_차이') = r.success_count_log - r.success_count_exist;
        row.('오판_감소') = (r.k - r.success_count_exist) - (r.k - r.success_count_log);
    else
        row.('기존_성공률') = NaN;
        row.('기존_합격자수_TP') = NaN;
        row.('기존_오판수_FP') = NaN;
        row.('기존_정확도') = NaN;
        row.('성공률_차이') = NaN;
        row.('합격자수_차이') = NaN;
        row.('오판_감소') = NaN;
    end

    % 실제 전체 합격/불합격 인원
    row.('실제_전체합격자') = pass_count;
    row.('실제_전체불합격자') = fail_count;

    % 놓친 합격자 (False Negative)
    row.('Logistic_놓친합격자_FN') = pass_count - r.success_count_log;
    if r.has_existing
        row.('기존_놓친합격자_FN') = pass_count - r.success_count_exist;
    else
        row.('기존_놓친합격자_FN') = NaN;
    end

    % 맞춘 불합격자 (True Negative)
    row.('Logistic_맞춘불합격자_TN') = fail_count - (r.k - r.success_count_log);
    if r.has_existing
        row.('기존_맞춘불합격자_TN') = fail_count - (r.k - r.success_count_exist);
    else
        row.('기존_맞춘불합격자_TN') = NaN;
    end

    detail_table = [detail_table; row]; %#ok<AGROW>
end

writetable(detail_table, output_path, 'Sheet', '상세분석', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 상세분석\n');

% 시트 5: 개인별 선발 여부 (각 비율별)
individual_selection = table();
individual_selection.ID = comp_data.ID;
individual_selection.('실제_합불') = arrayfun(@(x) iif(x==1, '합격', '불합격'), labels, 'UniformOutput', false);
individual_selection.('Logistic_점수') = round(score_logistic, 2);
individual_selection.('기존_점수') = round(score_existing, 2);

% 각 선발 비율별 선발 여부 추가
for i = 1:length(results)
    r = results{i};

    % Logistic 선발 여부
    [~, idx_log] = sort(score_logistic, 'descend');
    selected_log = false(n_new, 1);
    selected_log(idx_log(1:r.k)) = true;
    col_name_log = sprintf('Logistic_%s_선발', r.ratio_str);
    individual_selection.(col_name_log) = arrayfun(@(x) iif(x, 'O', 'X'), selected_log, 'UniformOutput', false);

    % 기존 선발 여부
    if r.has_existing
        [~, idx_exist] = sort(score_existing, 'descend', 'MissingPlacement', 'last');
        selected_exist = false(n_new, 1);
        selected_exist(idx_exist(1:r.k)) = true;
        col_name_exist = sprintf('기존_%s_선발', r.ratio_str);
        individual_selection.(col_name_exist) = arrayfun(@(x) iif(x, 'O', 'X'), selected_exist, 'UniformOutput', false);
    end
end

writetable(individual_selection, output_path, 'Sheet', '개인별_선발여부', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 개인별_선발여부\n');

% 시트 6: 혼동행렬 상세 (Confusion Matrix)
cm_detail = table();
for i = 1:length(results)
    r = results{i};

    % Logistic
    row_log = table();
    row_log.('선발비율') = {r.ratio_str};
    row_log.('방법') = {'Logistic'};
    row_log.('TP_선발하고_실제합격') = r.success_count_log;
    row_log.('FP_선발했지만_실제불합격') = r.k - r.success_count_log;
    row_log.('FN_선발안했지만_실제합격') = pass_count - r.success_count_log;
    row_log.('TN_선발안하고_실제불합격') = fail_count - (r.k - r.success_count_log);
    row_log.('선발인원_K') = r.k;
    row_log.('Precision_성공률') = r.success_log;
    row_log.('Recall_재현율') = (r.success_count_log / pass_count) * 100;

    cm_detail = [cm_detail; row_log]; %#ok<AGROW>

    % 기존
    if r.has_existing
        row_exist = table();
        row_exist.('선발비율') = {r.ratio_str};
        row_exist.('방법') = {'기존'};
        row_exist.('TP_선발하고_실제합격') = r.success_count_exist;
        row_exist.('FP_선발했지만_실제불합격') = r.k - r.success_count_exist;
        row_exist.('FN_선발안했지만_실제합격') = pass_count - r.success_count_exist;
        row_exist.('TN_선발안하고_실제불합격') = fail_count - (r.k - r.success_count_exist);
        row_exist.('선발인원_K') = r.k;
        row_exist.('Precision_성공률') = r.success_exist;
        row_exist.('Recall_재현율') = (r.success_count_exist / pass_count) * 100;

        cm_detail = [cm_detail; row_exist]; %#ok<AGROW>
    end
end

writetable(cm_detail, output_path, 'Sheet', '혼동행렬_상세', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 혼동행렬_상세\n');

fprintf('    ✓ Excel 파일: %s\n', config.output_filename);

%% 10) 마크다운 리포트 ---------------------------------------------------
fprintf('\n【STEP 10】 마크다운 리포트 생성\n');
fprintf('================================================================\n');

report_path = fullfile(config.output_dir, config.report_filename);
fid = fopen(report_path, 'w', 'n', 'UTF-8');

fprintf(fid, '# 선발 성공률 검증 리포트\n\n');
fprintf(fid, '**작성일**: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM'));
fprintf(fid, '**검증 대상**: 2025년 하반기 신규 입사자 (%d명)\n\n', n_new);
fprintf(fid, '---\n\n');

% 1. 핵심 요약
fprintf(fid, '## 📊 핵심 요약\n\n');
fprintf(fid, '### 🎯 핵심 질문\n\n');
fprintf(fid, '**"상위 몇 %%를 뽑았을 때, 실제 합격자가 몇 명인가?"**\n\n');

fprintf(fid, '### ✨ 결론부터\n\n');
fprintf(fid, '**대표 시나리오: %s (%d명 선발)**\n\n', rep_result.ratio_str, rep_result.k);

fprintf(fid, '| 방법 | 선발 성공률 | 실제 인원 | 평가 |\n');
fprintf(fid, '|------|------------|-----------|------|\n');
fprintf(fid, '| **Logistic 가중치** | **%.1f%%** | %d/%d명 | ', ...
    rep_result.success_log, rep_result.success_count_log, rep_result.k);

if rep_result.success_log >= 90
    fprintf(fid, '⭐⭐⭐ |\n');
elseif rep_result.success_log >= 80
    fprintf(fid, '⭐⭐ |\n');
else
    fprintf(fid, '⭐ |\n');
end

if rep_result.has_existing
    fprintf(fid, '| 기존 종합점수 | %.1f%% | %d/%d명 | - |\n', ...
        rep_result.success_exist, rep_result.success_count_exist, rep_result.k);
    fprintf(fid, '| **개선 효과** | **%+.1f%%p** | %+d명 | ', ...
        rep_result.improvement, rep_result.success_count_log - rep_result.success_count_exist);

    if rep_result.improvement > 10
        fprintf(fid, '🎉 **우수** |\n\n');
    elseif rep_result.improvement > 0
        fprintf(fid, '✅ **개선** |\n\n');
    else
        fprintf(fid, 'ℹ️ 유사 |\n\n');
    end
end

% 2. 데이터 개요
fprintf(fid, '---\n\n');
fprintf(fid, '## 1. 데이터 개요\n\n');
fprintf(fid, '| 구분 | 인원 | 비율 |\n');
fprintf(fid, '|------|------|------|\n');
fprintf(fid, '| **실제 합격자** | %d명 | %.1f%% |\n', pass_count, baseline_rate);
fprintf(fid, '| **실제 불합격자** | %d명 | %.1f%% |\n', fail_count, fail_count/n_new*100);
fprintf(fid, '| **전체** | %d명 | 100%% |\n\n', n_new);

fprintf(fid, '> HR이 결정한 합격/불합격이 **정답(Ground Truth)**이며, 선발 성공률은 이 정답을 얼마나 맞추는지 평가합니다.\n\n');

% 3. 선발 비율별 결과
fprintf(fid, '---\n\n');
fprintf(fid, '## 2. 선발 비율별 성공률\n\n');

fprintf(fid, '| 선발 비율 | 선발 인원 | Logistic | 기존 | 개선 | 평가 |\n');
fprintf(fid, '|-----------|-----------|----------|------|------|------|\n');

for i = 1:length(results)
    r = results{i};
    stars = '';
    if r.success_log >= 80
        stars = '⭐⭐⭐';
    elseif r.success_log >= 70
        stars = '⭐⭐';
    elseif r.success_log >= 60
        stars = '⭐';
    end

    if r.has_existing
        fprintf(fid, '| %s | %d명 | %.1f%% (%d/%d) | %.1f%% (%d/%d) | %+.1f%%p | %s |\n', ...
            r.ratio_str, r.k, r.success_log, r.success_count_log, r.k, ...
            r.success_exist, r.success_count_exist, r.k, ...
            r.improvement, stars);
    else
        fprintf(fid, '| %s | %d명 | %.1f%% (%d/%d) | N/A | N/A | %s |\n', ...
            r.ratio_str, r.k, r.success_log, r.success_count_log, r.k, ...
            stars);
    end
end
fprintf(fid, '\n');

fprintf(fid, '### 해석 가이드\n\n');
fprintf(fid, '- **선발 성공률**: 선발한 인원 중 실제 합격자 비율 (Precision)\n');
fprintf(fid, '- **개선**: Logistic - 기존 (양수면 Logistic이 우수)\n');
fprintf(fid, '- **⭐⭐⭐**: 80%% 이상 (매우 우수)\n');
fprintf(fid, '- **⭐⭐**: 70~80%% (우수)\n');
fprintf(fid, '- **⭐**: 60~70%% (보통)\n\n');

% 4. 대표 시나리오 상세
fprintf(fid, '---\n\n');
fprintf(fid, '## 3. 대표 시나리오 상세: %s\n\n', rep_result.ratio_str);

fprintf(fid, '### 선발 조건\n\n');
fprintf(fid, '- 선발 비율: %s\n', rep_result.ratio_str);
fprintf(fid, '- 선발 인원: %d명\n\n', rep_result.k);

fprintf(fid, '### Logistic 가중치로 선발 시\n\n');
fprintf(fid, '- **성공률**: %.1f%%\n', rep_result.success_log);
fprintf(fid, '- **실제 합격자**: %d명 (선발 %d명 중)\n', rep_result.success_count_log, rep_result.k);
fprintf(fid, '- **오판**: %d명 (불합격자를 잘못 선발)\n\n', rep_result.k - rep_result.success_count_log);

if rep_result.has_existing
    fprintf(fid, '### 기존 종합점수로 선발 시\n\n');
    fprintf(fid, '- **성공률**: %.1f%%\n', rep_result.success_exist);
    fprintf(fid, '- **실제 합격자**: %d명 (선발 %d명 중)\n', rep_result.success_count_exist, rep_result.k);
    fprintf(fid, '- **오판**: %d명 (불합격자를 잘못 선발)\n\n', rep_result.k - rep_result.success_count_exist);

    fprintf(fid, '### 비교: 개선 효과\n\n');
    fprintf(fid, '- **성공률 차이**: %+.1f%%p (%.1f%% → %.1f%%)\n', ...
        rep_result.improvement, rep_result.success_exist, rep_result.success_log);
    fprintf(fid, '- **인원 차이**: %+d명 더 정확하게 선발\n', ...
        rep_result.success_count_log - rep_result.success_count_exist);

    if rep_result.improvement > 10
        fprintf(fid, '\n✅ **Logistic 가중치가 %.1f%%p 더 우수합니다!**\n\n', rep_result.improvement);
    elseif rep_result.improvement > 0
        fprintf(fid, '\n✅ **Logistic 가중치가 %.1f%%p 개선되었습니다.**\n\n', rep_result.improvement);
    else
        fprintf(fid, '\nℹ️ **두 방법이 유사한 성능을 보입니다.**\n\n');
    end
end

% 5. 결론 및 권장사항
fprintf(fid, '---\n\n');
fprintf(fid, '## 4. 결론 및 권장사항\n\n');

fprintf(fid, '### 4.1 핵심 결론\n\n');

% 최고 성공률 시나리오 찾기
best_idx = 1;
best_success = results{1}.success_log;
for i = 2:length(results)
    if results{i}.success_log > best_success
        best_success = results{i}.success_log;
        best_idx = i;
    end
end
best_result = results{best_idx};

fprintf(fid, '1. **최고 성공률**: %s 선발 시 **%.1f%%** ⭐\n', ...
    best_result.ratio_str, best_result.success_log);
fprintf(fid, '   - %d명 선발 중 %d명이 실제 합격자 (%d명 오판)\n', best_result.k, best_result.success_count_log, ...
    best_result.k - best_result.success_count_log);

if rep_result.has_existing && rep_result.improvement > 5
    fprintf(fid, '2. **Logistic 가중치 우수**\n');
    fprintf(fid, '   - 대표 시나리오(%s)에서 %.1f%%p 개선\n', rep_result.ratio_str, rep_result.improvement);
    fprintf(fid, '   - %d명 더 정확하게 선발\n\n', rep_result.success_count_log - rep_result.success_count_exist);
end

fprintf(fid, '### 4.2 실무 적용 권장\n\n');

% 80% 이상인 시나리오 찾기
excellent_scenarios = {};
for i = 1:length(results)
    if results{i}.success_log >= 80
        excellent_scenarios{end+1} = results{i}; %#ok<AGROW>
    end
end

if ~isempty(excellent_scenarios)
    fprintf(fid, '#### ✅ 매우 우수한 시나리오 (80%% 이상)\n\n');
    for i = 1:length(excellent_scenarios)
        sc = excellent_scenarios{i};
        fprintf(fid, '- **%s**: %.1f%% (%d/%d명 성공)\n', ...
            sc.ratio_str, sc.success_log, sc.success_count_log, sc.k);
    end
    fprintf(fid, '\n');
end

if rep_result.has_existing && rep_result.improvement > 10
    fprintf(fid, '#### 💡 권장사항\n\n');
    fprintf(fid, '1. **Logistic 가중치 적극 활용 권장**\n');
    fprintf(fid, '   - 기존 대비 뚜렷한 개선 효과 확인\n');
    fprintf(fid, '   - 선발 성공률 %.1f%%로 매우 우수\n\n', rep_result.success_log);
elseif rep_result.has_existing && rep_result.improvement > 0
    fprintf(fid, '#### 💡 권장사항\n\n');
    fprintf(fid, '1. **Logistic 가중치 사용 권장**\n');
    fprintf(fid, '   - 기존 대비 개선 효과 있음\n');
    fprintf(fid, '   - 추가 데이터로 지속 검증 필요\n\n');
else
    fprintf(fid, '#### 💡 권장사항\n\n');
    fprintf(fid, '1. **두 방법 병행 검토**\n');
    fprintf(fid, '   - 유사한 성능으로 보완적 활용 가능\n\n');
end

fprintf(fid, '### 4.3 주의사항\n\n');
fprintf(fid, '1. **소규모 데이터**: %d명 분석으로 추가 검증 필요\n', n_new);
fprintf(fid, '2. **지속적 모니터링**: 실제 적용 후 성과 추적\n');
fprintf(fid, '3. **다면 평가**: 정량 지표 외 정성적 요소 고려\n');
fprintf(fid, '4. **목적별 조정**: 선발 목적에 맞는 비율 선택\n\n');

% 6. 참고 지표
fprintf(fid, '---\n\n');
fprintf(fid, '## 5. 참고 지표 (부록)\n\n');

fprintf(fid, '### ROC-AUC (전체 예측 능력)\n\n');
fprintf(fid, '- Logistic 가중치: %.4f\n', AUC_logistic);
if ~isnan(AUC_existing)
    fprintf(fid, '- 기존 종합점수: %.4f\n', AUC_existing);
    fprintf(fid, '- 개선도: %.1f%%\n\n', (AUC_logistic/AUC_existing - 1)*100);
else
    fprintf(fid, '- 기존 종합점수: N/A\n\n');
end

fprintf(fid, '### Cohen''s d (통계적 효과 크기)\n\n');
fprintf(fid, '- Logistic 가중치: %.3f', abs(cohens_d_log));
if abs(cohens_d_log) > 0.8
    fprintf(fid, ' (Large)\n');
elseif abs(cohens_d_log) > 0.5
    fprintf(fid, ' (Medium)\n');
else
    fprintf(fid, ' (Small)\n');
end

if ~isnan(cohens_d_exist)
    fprintf(fid, '- 기존 종합점수: %.3f', abs(cohens_d_exist));
    if abs(cohens_d_exist) > 0.8
        fprintf(fid, ' (Large)\n\n');
    elseif abs(cohens_d_exist) > 0.5
        fprintf(fid, ' (Medium)\n\n');
    else
        fprintf(fid, ' (Small)\n\n');
    end
end

% 7. 출력 파일
fprintf(fid, '---\n\n');
fprintf(fid, '## 6. 출력 파일\n\n');
fprintf(fid, '1. **Excel**: `%s`\n', config.output_filename);
fprintf(fid, '   - 선발성공률_요약: 선발 비율별 성공률\n');
fprintf(fid, '   - 개인별_점수_순위: ID별 점수 및 순위\n');
fprintf(fid, '   - 참고지표: ROC-AUC, Cohen''s d 등\n\n');

fprintf(fid, '2. **시각화**: `%s`\n', config.plot_filename);
fprintf(fid, '   - 선발 성공률 비교\n');
fprintf(fid, '   - 점수 분포\n');
fprintf(fid, '   - 개선 효과\n\n');

fprintf(fid, '---\n\n');
fprintf(fid, '*본 리포트는 MATLAB 자동 분석으로 생성되었습니다.*\n\n');
fprintf(fid, '*생성일시: %s*\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

fclose(fid);
fprintf('    ✓ 리포트 파일: %s\n', config.report_filename);

%% 11) 최종 리포트 (콘솔) ------------------------------------------------
fprintf('\n');
fprintf('================================================================\n');
fprintf('              🎯 선발 성공률 검증 최종 리포트 🎯\n');
fprintf('================================================================\n\n');

fprintf('【 데이터 개요 】\n');
fprintf('  • 분석 대상: %d명\n', n_new);
fprintf('  • 실제 합격자: %d명 (%.1f%%)\n', pass_count, baseline_rate);
fprintf('  • 실제 불합격자: %d명 (%.1f%%)\n\n', fail_count, fail_count/n_new*100);

fprintf('【 핵심 결과 】\n\n');
fprintf('  대표 시나리오: %s (%d명 선발)\n\n', rep_result.ratio_str, rep_result.k);

fprintf('  ┌──────────────┬──────────┬──────────┬─────────┐\n');
fprintf('  │    방법      │ 성공률   │ 실제인원 │  오판   │\n');
fprintf('  ├──────────────┼──────────┼──────────┼─────────┤\n');
fprintf('  │ Logistic     │  %.1f%%  │  %d/%d명  │  %d명   │\n', ...
    rep_result.success_log, rep_result.success_count_log, rep_result.k, ...
    rep_result.k - rep_result.success_count_log);

if rep_result.has_existing
    fprintf('  │ 기존 점수    │  %.1f%%  │  %d/%d명  │  %d명   │\n', ...
        rep_result.success_exist, rep_result.success_count_exist, rep_result.k, ...
        rep_result.k - rep_result.success_count_exist);
    fprintf('  │ 개선         │ %+.1f%%p │  %+d명   │  %d명↓  │\n', ...
        rep_result.improvement, rep_result.success_count_log - rep_result.success_count_exist, ...
        (rep_result.k - rep_result.success_count_exist) - (rep_result.k - rep_result.success_count_log));
end

fprintf('  └──────────────┴──────────┴──────────┴─────────┘\n\n');

fprintf('【 결론 】\n');

if rep_result.has_existing && rep_result.improvement > 10
    fprintf('  ✅ Logistic 가중치가 뚜렷한 개선 효과를 보였습니다!\n');
    fprintf('     %.1f%%p 향상 (%.1f%% → %.1f%%)\n\n', ...
        rep_result.improvement, rep_result.success_exist, rep_result.success_log);
elseif rep_result.has_existing && rep_result.improvement > 0
    fprintf('  ✅ Logistic 가중치가 개선 효과를 보였습니다.\n');
    fprintf('     %.1f%%p 향상 (%.1f%% → %.1f%%)\n\n', ...
        rep_result.improvement, rep_result.success_exist, rep_result.success_log);
else
    fprintf('  ℹ️ 두 방법이 유사한 성능을 보였습니다.\n\n');
end

fprintf('【 최고 성공률 】\n');
fprintf('  • %s 선발 시: %.1f%% (%d/%d명 성공) ⭐\n\n', ...
    best_result.ratio_str, best_result.success_log, best_result.success_count_log, best_result.k);

fprintf('【 출력 파일 】\n');
fprintf('  • Excel: %s\n', config.output_filename);
fprintf('  • 그래프: %s\n', config.plot_filename);
fprintf('  • 리포트: %s\n\n', config.report_filename);

fprintf('================================================================\n');
fprintf('  완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('  위치: %s\n', config.output_dir);
fprintf('================================================================\n\n');

fprintf('💡 핵심 메시지:\n');
fprintf('   • 선발 성공률 (Precision) = 뽑은 사람 중 실제 합격자 비율\n');
fprintf('   • HR이 결정한 합격/불합격 = 정답 (Ground Truth)\n');
fprintf('   • 높을수록 정답을 잘 맞추는 것 (100%% = 완벽)\n');
fprintf('   • 실무에서는 목적에 맞는 선발 비율 선택\n\n');

% Helper function
function out = iif(condition, true_val, false_val)
    if condition
        out = true_val;
    else
        out = false_val;
    end
end
