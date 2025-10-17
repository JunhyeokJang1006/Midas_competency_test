% =======================================================================
%     예측 성능 검증: Logistic 가중치 vs 기존 종합점수
% =======================================================================
% 목적:
%   - 신규 입사자 합/불 예측 성능 비교
%   - 기존 종합점수 대비 Logistic 가중치의 개선 효과 측정
%   - 실무 선발 시나리오별 예측 정확도 분석
%
% 핵심 질문:
%   "새 가중치로 예측하면 기존보다 얼마나 더 정확한가?"
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
fprintf('     예측 성능 검증: Logistic 가중치 vs 기존 종합점수\n');
fprintf('================================================================\n\n');

%% 1) 설정 ---------------------------------------------------------------
fprintf('【STEP 1】 설정\n');
fprintf('================================================================\n');

config = struct();
config.new_comp_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_역검 점수.xlsx';
config.new_onboarding_file = 'D:\project\HR데이터\데이터\25년신규입사자 데이터\25년 하반기 입사자_온보딩 점수.xlsx';
config.weight_file = 'D:\project\HR데이터\결과\자가불소_revised_talent\integrated_analysis_results.mat';
config.output_dir = 'D:\project\HR데이터\결과\가중치검증';
config.timestamp = datestr(now, 'yyyymmdd_HHMMSS');
config.output_filename = sprintf('prediction_performance_%s.xlsx', config.timestamp);
config.plot_filename = sprintf('performance_comparison_%s.png', config.timestamp);
config.report_filename = sprintf('예측성능_검증리포트_%s.md', config.timestamp);

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

%% 3) 신규 입사자 데이터 로드 --------------------------------------------
fprintf('\n【STEP 3】 신규 입사자 데이터 로드\n');
fprintf('================================================================\n');

% 역량검사 데이터
comp_data = readtable(config.new_comp_file, 'Sheet', '역량검사_상위항목', ...
    'VariableNamingRule', 'preserve');
fprintf('  ✓ 역량검사 데이터: %d명\n', height(comp_data));

% 기존 종합점수
score_data = readtable(config.new_comp_file, 'Sheet', '역량검사_종합점수', ...
    'VariableNamingRule', 'preserve');
fprintf('  ✓ 기존 종합점수: %d명\n', height(score_data));

% 온보딩 데이터 (합불 여부)
onboarding_data = readtable(config.new_onboarding_file, 'VariableNamingRule', 'preserve');
fprintf('  ✓ 온보딩 데이터: %d명\n', height(onboarding_data));

% ID를 문자열로 통일
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

%% 4) 레이블 생성 및 필터링 ---------------------------------------------
fprintf('\n【STEP 4】 레이블 생성 및 필터링\n');
fprintf('================================================================\n');

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

fprintf('  ✓ 분석 대상: %d명\n', n_new);
fprintf('  ✓ 합격: %d명 (%.1f%%)\n', pass_count, baseline_rate);
fprintf('  ✓ 불합격: %d명 (%.1f%%)\n\n', fail_count, fail_count/n_new*100);

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

fprintf('  ✓ Logistic 가중치 점수: %.1f ~ %.1f점\n', min(score_logistic), max(score_logistic));
fprintf('  ✓ 기존 종합점수: %.1f ~ %.1f점\n', min(score_existing), max(score_existing));
fprintf('  ✓ 점수 계산 완료\n\n');

%% 6) 예측 성능 비교 분석 (★ 핵심) ----------------------------------------
fprintf('【STEP 6】 예측 성능 비교 분석 (★ 핵심)\n');
fprintf('================================================================\n');
fprintf('\n  🎯 핵심 질문: 어느 방법이 합격/불합격을 더 정확하게 예측하는가?\n\n');

% 실무 시나리오: 상위 N% 선발
selection_ratios = [0.15, 0.20, 0.28, 0.33, 0.44, 0.56];  % 15%, 20%, 28%, 33%, 44%, 56%
selection_scenarios = {};

for ratio = selection_ratios
    k = round(n_new * ratio);
    if k < 1, k = 1; end
    if k > n_new, k = n_new; end
    
    scenario = struct();
    scenario.ratio = ratio * 100;
    scenario.k = k;
    scenario.ratio_str = sprintf('상위 %.0f%%', ratio * 100);
    scenario.k_str = sprintf('%d/%d명', k, n_new);
    
    %% Logistic 가중치로 예측
    [~, idx_log] = sort(score_logistic, 'descend');
    pred_log = zeros(n_new, 1);
    pred_log(idx_log(1:k)) = 1;  % 상위 k명만 합격 예측
    
    % 정확도 지표
    TP_log = sum(pred_log == 1 & labels == 1);  % True Positive
    FP_log = sum(pred_log == 1 & labels == 0);  % False Positive
    FN_log = sum(pred_log == 0 & labels == 1);  % False Negative
    TN_log = sum(pred_log == 0 & labels == 0);  % True Negative
    
    scenario.logistic.TP = TP_log;
    scenario.logistic.FP = FP_log;
    scenario.logistic.FN = FN_log;
    scenario.logistic.TN = TN_log;
    scenario.logistic.accuracy = (TP_log + TN_log) / n_new * 100;
    scenario.logistic.precision = TP_log / max(TP_log + FP_log, 1) * 100;
    scenario.logistic.recall = TP_log / max(TP_log + FN_log, 1) * 100;
    scenario.logistic.f1 = 2 * scenario.logistic.precision * scenario.logistic.recall / ...
        max(scenario.logistic.precision + scenario.logistic.recall, 1);
    
    %% 기존 종합점수로 예측
    if sum(~isnan(score_existing)) >= k
        [~, idx_exist] = sort(score_existing, 'descend', 'MissingPlacement', 'last');
        pred_exist = zeros(n_new, 1);
        pred_exist(idx_exist(1:k)) = 1;
        
        TP_exist = sum(pred_exist == 1 & labels == 1);
        FP_exist = sum(pred_exist == 1 & labels == 0);
        FN_exist = sum(pred_exist == 0 & labels == 1);
        TN_exist = sum(pred_exist == 0 & labels == 0);
        
        scenario.existing.TP = TP_exist;
        scenario.existing.FP = FP_exist;
        scenario.existing.FN = FN_exist;
        scenario.existing.TN = TN_exist;
        scenario.existing.accuracy = (TP_exist + TN_exist) / n_new * 100;
        scenario.existing.precision = TP_exist / max(TP_exist + FP_exist, 1) * 100;
        scenario.existing.recall = TP_exist / max(TP_exist + FN_exist, 1) * 100;
        scenario.existing.f1 = 2 * scenario.existing.precision * scenario.existing.recall / ...
            max(scenario.existing.precision + scenario.existing.recall, 1);
        
        scenario.has_existing = true;
    else
        scenario.has_existing = false;
    end
    
    selection_scenarios{end+1} = scenario; %#ok<AGROW>
end

%% 주요 시나리오 출력
fprintf('  ┌────────────┬──────────┬──────────────────────────────────────┐\n');
fprintf('  │ 선발 비율  │ 선발인원 │          예측 정확도 (Accuracy)      │\n');
fprintf('  │            │          │  Logistic  │   기존    │   개선      │\n');
fprintf('  ├────────────┼──────────┼────────────┼───────────┼─────────────┤\n');

for i = 1:length(selection_scenarios)
    sc = selection_scenarios{i};
    if sc.has_existing
        fprintf('  │ %9s  │ %8s │   %.1f%%    │  %.1f%%   │  %+.1f%%p    │\n', ...
            sc.ratio_str, sc.k_str, sc.logistic.accuracy, ...
            sc.existing.accuracy, sc.logistic.accuracy - sc.existing.accuracy);
    else
        fprintf('  │ %9s  │ %8s │   %.1f%%    │    N/A    │     N/A      │\n', ...
            sc.ratio_str, sc.k_str, sc.logistic.accuracy);
    end
end
fprintf('  └────────────┴──────────┴────────────┴───────────┴─────────────┘\n\n');

fprintf('  💡 해석:\n');
fprintf('     • 정확도 = (맞춘 예측) / (전체 예측) × 100%%\n');
fprintf('     • 높을수록 합격/불합격을 정확하게 구분\n');
fprintf('     • 개선 = Logistic - 기존 (양수면 Logistic이 더 우수)\n\n');

%% 7) 상세 성능 지표 분석 ------------------------------------------------
fprintf('【STEP 7】 상세 성능 지표 분석\n');
fprintf('================================================================\n\n');

% 대표 시나리오 선택 (상위 28% ≈ 5명)
rep_idx = find([selection_scenarios{:}].ratio == 28);
if isempty(rep_idx)
    rep_idx = 3;  % 기본값
end
rep_scenario = selection_scenarios{rep_idx};

fprintf('  📊 대표 시나리오: %s 선발 (%s)\n\n', rep_scenario.ratio_str, rep_scenario.k_str);

fprintf('  【Logistic 가중치 예측 결과】\n');
fprintf('    ┌────────────┬─────────┬─────────┐\n');
fprintf('    │            │ 실제합격│ 실제불합│\n');
fprintf('    ├────────────┼─────────┼─────────┤\n');
fprintf('    │ 예측합격   │   %2d    │   %2d    │\n', rep_scenario.logistic.TP, rep_scenario.logistic.FP);
fprintf('    │ 예측불합격 │   %2d    │   %2d    │\n', rep_scenario.logistic.FN, rep_scenario.logistic.TN);
fprintf('    └────────────┴─────────┴─────────┘\n');
fprintf('    • 정확도: %.1f%% (%d/%d명 맞춤)\n', rep_scenario.logistic.accuracy, ...
        rep_scenario.logistic.TP + rep_scenario.logistic.TN, n_new);
fprintf('    • 정밀도: %.1f%% (합격 예측 중 실제 합격 비율)\n', rep_scenario.logistic.precision);
fprintf('    • 재현율: %.1f%% (실제 합격자 중 찾아낸 비율)\n', rep_scenario.logistic.recall);
fprintf('    • F1 Score: %.2f\n\n', rep_scenario.logistic.f1);

if rep_scenario.has_existing
    fprintf('  【기존 종합점수 예측 결과】\n');
    fprintf('    ┌────────────┬─────────┬─────────┐\n');
    fprintf('    │            │ 실제합격│ 실제불합│\n');
    fprintf('    │ 예측합격   │   %2d    │   %2d    │\n', rep_scenario.existing.TP, rep_scenario.existing.FP);
    fprintf('    │ 예측불합격 │   %2d    │   %2d    │\n', rep_scenario.existing.FN, rep_scenario.existing.TN);
    fprintf('    └────────────┴─────────┴─────────┘\n');
    fprintf('    • 정확도: %.1f%% (%d/%d명 맞춤)\n', rep_scenario.existing.accuracy, ...
        rep_scenario.existing.TP + rep_scenario.existing.TN, n_new);
    fprintf('    • 정밀도: %.1f%%\n', rep_scenario.existing.precision);
    fprintf('    • 재현율: %.1f%%\n', rep_scenario.existing.recall);
    fprintf('    • F1 Score: %.2f\n\n', rep_scenario.existing.f1);
    
    fprintf('  【성능 개선 효과】\n');
    fprintf('    • 정확도: %+.1f%%p 개선 (%.1f%% → %.1f%%)\n', ...
        rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy, ...
        rep_scenario.existing.accuracy, rep_scenario.logistic.accuracy);
    fprintf('    • 오판 감소: %d명 → %d명 (%d명 감소)\n', ...
        rep_scenario.existing.FP + rep_scenario.existing.FN, ...
        rep_scenario.logistic.FP + rep_scenario.logistic.FN, ...
        (rep_scenario.existing.FP + rep_scenario.existing.FN) - ...
        (rep_scenario.logistic.FP + rep_scenario.logistic.FN));
    fprintf('    • F1 Score: %+.2f 개선\n\n', ...
        rep_scenario.logistic.f1 - rep_scenario.existing.f1);
end


%% 9) 시각화 -------------------------------------------------------------
fprintf('【STEP 9】 시각화 생성\n');
fprintf('================================================================\n');

fig = figure('Position', [100, 100, 1800, 1200]);

% 1. 정확도 비교 (막대 그래프)
subplot(3, 3, 1);
ratios_for_plot = [selection_scenarios{:}];
acc_log = [ratios_for_plot.logistic];
acc_log_vals = [acc_log.accuracy];

if selection_scenarios{1}.has_existing
    acc_exist = [ratios_for_plot.existing];
    acc_exist_vals = [acc_exist.accuracy];
    bar(1:length(selection_scenarios), [acc_log_vals', acc_exist_vals']);
    legend({'Logistic', '기존'}, 'Location', 'best');
else
    bar(1:length(selection_scenarios), acc_log_vals');
    legend({'Logistic'}, 'Location', 'best');
end
xlabel('선발 비율', 'FontWeight', 'bold');
ylabel('정확도 (%)', 'FontWeight', 'bold');
title('선발 비율별 예측 정확도', 'FontSize', 13, 'FontWeight', 'bold');
xticks(1:length(selection_scenarios));
xticklabels({ratios_for_plot.ratio_str});
xtickangle(45);
grid on;
ylim([0 100]);

% 2. 정밀도(Precision) 비교
subplot(3, 3, 2);
prec_log_vals = [acc_log.precision];
if selection_scenarios{1}.has_existing
    prec_exist_vals = [acc_exist.precision];
    bar(1:length(selection_scenarios), [prec_log_vals', prec_exist_vals']);
    legend({'Logistic', '기존'}, 'Location', 'best');
else
    bar(1:length(selection_scenarios), prec_log_vals');
end
xlabel('선발 비율', 'FontWeight', 'bold');
ylabel('정밀도 (%)', 'FontWeight', 'bold');
title('정밀도 (합격 예측 정확도)', 'FontSize', 13, 'FontWeight', 'bold');
xticks(1:length(selection_scenarios));
xticklabels({ratios_for_plot.ratio_str});
xtickangle(45);
grid on;
ylim([0 100]);

% 3. 재현율(Recall) 비교
subplot(3, 3, 3);
recall_log_vals = [acc_log.recall];
if selection_scenarios{1}.has_existing
    recall_exist_vals = [acc_exist.recall];
    bar(1:length(selection_scenarios), [recall_log_vals', recall_exist_vals']);
    legend({'Logistic', '기존'}, 'Location', 'best');
else
    bar(1:length(selection_scenarios), recall_log_vals');
end
xlabel('선발 비율', 'FontWeight', 'bold');
ylabel('재현율 (%)', 'FontWeight', 'bold');
title('재현율 (합격자 포착률)', 'FontSize', 13, 'FontWeight', 'bold');
xticks(1:length(selection_scenarios));
xticklabels({ratios_for_plot.ratio_str});
xtickangle(45);
grid on;
ylim([0 100]);

% 4. F1 Score 비교
subplot(3, 3, 4);
f1_log_vals = [acc_log.f1];
if selection_scenarios{1}.has_existing
    f1_exist_vals = [acc_exist.f1];
    bar(1:length(selection_scenarios), [f1_log_vals', f1_exist_vals']);
    legend({'Logistic', '기존'}, 'Location', 'best');
else
    bar(1:length(selection_scenarios), f1_log_vals');
end
xlabel('선발 비율', 'FontWeight', 'bold');
ylabel('F1 Score', 'FontWeight', 'bold');
title('F1 Score (종합 성능)', 'FontSize', 13, 'FontWeight', 'bold');
xticks(1:length(selection_scenarios));
xticklabels({ratios_for_plot.ratio_str});
xtickangle(45);
grid on;
ylim([0 100]);

% 5. 혼동행렬 - Logistic (대표 시나리오)
subplot(3, 3, 5);
cm_log = [rep_scenario.logistic.TP, rep_scenario.logistic.FP; 
          rep_scenario.logistic.FN, rep_scenario.logistic.TN];
imagesc(cm_log);
colormap(flipud(hot));
colorbar;
xticks([1 2]);
xticklabels({'실제 합격', '실제 불합격'});
yticks([1 2]);
yticklabels({'예측 합격', '예측 불합격'});
title(sprintf('Logistic 혼동행렬 (%s)', rep_scenario.ratio_str), 'FontSize', 12, 'FontWeight', 'bold');
text(1, 1, sprintf('%d', rep_scenario.logistic.TP), 'HorizontalAlignment', 'center', 'FontSize', 16, 'FontWeight', 'bold');
text(2, 1, sprintf('%d', rep_scenario.logistic.FP), 'HorizontalAlignment', 'center', 'FontSize', 16, 'FontWeight', 'bold');
text(1, 2, sprintf('%d', rep_scenario.logistic.FN), 'HorizontalAlignment', 'center', 'FontSize', 16, 'FontWeight', 'bold');
text(2, 2, sprintf('%d', rep_scenario.logistic.TN), 'HorizontalAlignment', 'center', 'FontSize', 16, 'FontWeight', 'bold');


% 9. 종합 요약
subplot(3, 3, 9);
axis off;
text(0.05, 0.95, '【예측 성능 검증 요약】', 'FontSize', 14, 'FontWeight', 'bold');
y_pos = 0.82;

text(0.05, y_pos, sprintf('분석 대상: %d명', n_new), 'FontSize', 10);
y_pos = y_pos - 0.08;
text(0.05, y_pos, sprintf('합격: %d명 (%.1f%%)', pass_count, baseline_rate), 'FontSize', 10);
y_pos = y_pos - 0.12;

text(0.05, y_pos, sprintf('【대표 시나리오: %s】', rep_scenario.ratio_str), 'FontSize', 11, 'FontWeight', 'bold');
y_pos = y_pos - 0.08;
text(0.05, y_pos, sprintf('• Logistic 정확도: %.1f%%', rep_scenario.logistic.accuracy), 'FontSize', 10);
y_pos = y_pos - 0.08;

if rep_scenario.has_existing
    text(0.05, y_pos, sprintf('• 기존 정확도: %.1f%%', rep_scenario.existing.accuracy), 'FontSize', 10);
    y_pos = y_pos - 0.08;
    text(0.05, y_pos, sprintf('• 개선: %+.1f%%p', rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy), ...
        'FontSize', 10, 'Color', [0.2 0.7 0.3], 'FontWeight', 'bold');
    y_pos = y_pos - 0.12;
end

text(0.05, y_pos, '【결론】', 'FontSize', 12, 'FontWeight', 'bold', 'Color', [0 0.4 0.8]);
y_pos = y_pos - 0.08;

if rep_scenario.has_existing && rep_scenario.logistic.accuracy > rep_scenario.existing.accuracy
    text(0.05, y_pos, '✅ Logistic 가중치', 'FontSize', 11, 'Color', [0.2 0.7 0.3], 'FontWeight', 'bold');
    y_pos = y_pos - 0.08;
    text(0.05, y_pos, '   예측 정확도 우수!', 'FontSize', 10, 'Color', [0.2 0.7 0.3]);
else
    text(0.05, y_pos, 'ℹ️ 두 방법 유사', 'FontSize', 11);
end

sgtitle('예측 성능 검증 결과', 'FontSize', 16, 'FontWeight', 'bold');

% 저장
plot_path = fullfile(config.output_dir, config.plot_filename);
saveas(fig, plot_path);
fprintf('  ✓ 그래프 저장: %s\n', config.plot_filename);

%% 10) 엑셀 저장 ---------------------------------------------------------
fprintf('\n【STEP 10】 엑셀 결과 저장\n');
fprintf('================================================================\n');

output_path = fullfile(config.output_dir, config.output_filename);

% 시트 1: 성능 비교 요약
summary_table = table();
for i = 1:length(selection_scenarios)
    sc = selection_scenarios{i};
    row = table();
    row.('선발비율') = {sc.ratio_str};
    row.('선발인원') = {sc.k_str};
    row.('Logistic_정확도') = sc.logistic.accuracy;
    row.('Logistic_정밀도') = sc.logistic.precision;
    row.('Logistic_재현율') = sc.logistic.recall;
    row.('Logistic_F1') = sc.logistic.f1;
    
    if sc.has_existing
        row.('기존_정확도') = sc.existing.accuracy;
        row.('기존_정밀도') = sc.existing.precision;
        row.('기존_재현율') = sc.existing.recall;
        row.('기존_F1') = sc.existing.f1;
        row.('정확도_개선') = sc.logistic.accuracy - sc.existing.accuracy;
    else
        row.('기존_정확도') = NaN;
        row.('기존_정밀도') = NaN;
        row.('기존_재현율') = NaN;
        row.('기존_F1') = NaN;
        row.('정확도_개선') = NaN;
    end
    
    summary_table = [summary_table; row]; %#ok<AGROW>
end

writetable(summary_table, output_path, 'Sheet', '성능비교_요약', 'WriteMode', 'overwrite');
fprintf('  ✓ 시트 저장: 성능비교_요약\n');

% 시트 2: 혼동행렬 (대표 시나리오)
cm_table = table();
cm_table.('구분') = {'Logistic_TP'; 'Logistic_FP'; 'Logistic_FN'; 'Logistic_TN'};
cm_table.('값') = [rep_scenario.logistic.TP; rep_scenario.logistic.FP; 
                    rep_scenario.logistic.FN; rep_scenario.logistic.TN];

if rep_scenario.has_existing
    cm_table_exist = table();
    cm_table_exist.('구분') = {'기존_TP'; '기존_FP'; '기존_FN'; '기존_TN'};
    cm_table_exist.('값') = [rep_scenario.existing.TP; rep_scenario.existing.FP; 
                              rep_scenario.existing.FN; rep_scenario.existing.TN];
    cm_table = [cm_table; cm_table_exist];
end

writetable(cm_table, output_path, 'Sheet', '혼동행렬', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 혼동행렬\n');

% 시트 3: 개인별 점수 및 예측 결과
individual_table = table();
individual_table.ID = comp_data.ID;
individual_table.('실제_합불') = labels;
individual_table.('Logistic_점수') = round(score_logistic, 2);
[~, rank_log] = sort(score_logistic, 'descend');
individual_table.('Logistic_순위') = zeros(n_new, 1);
individual_table.('Logistic_순위')(rank_log) = (1:n_new)';

% 대표 시나리오 예측 결과 추가
pred_log_rep = zeros(n_new, 1);
pred_log_rep(rank_log(1:rep_scenario.k)) = 1;
individual_table.('Logistic_예측') = pred_log_rep;

individual_table.('기존_점수') = round(score_existing, 2);

writetable(individual_table, output_path, 'Sheet', '개인별_점수_예측', 'WriteMode', 'append');
fprintf('  ✓ 시트 저장: 개인별_점수_예측\n');

fprintf('\n  ✅ 엑셀 파일 저장 완료: %s\n', output_path);

%% 11) 마크다운 리포트 ---------------------------------------------------
fprintf('\n【STEP 11】 실무 담당자용 리포트 생성\n');
fprintf('================================================================\n');

report_path = fullfile(config.output_dir, config.report_filename);
fid = fopen(report_path, 'w', 'n', 'UTF-8');

fprintf(fid, '# 예측 성능 검증 리포트\n\n');
fprintf(fid, '**작성일**: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM'));
fprintf(fid, '**검증 대상**: 2025년 하반기 신규 입사자 (%d명)\n\n', n_new);
fprintf(fid, '---\n\n');

% 1. Executive Summary
fprintf(fid, '## 📊 핵심 요약 (Executive Summary)\n\n');
fprintf(fid, '### 🎯 핵심 질문\n\n');
fprintf(fid, '**"새 Logistic 가중치로 예측하면 기존보다 얼마나 더 정확한가?"**\n\n');

fprintf(fid, '### ✨ 결론부터\n\n');
fprintf(fid, '| 방법 | 정확도 | 평가 |\n');
fprintf(fid, '|------|--------|------|\n');
fprintf(fid, '| **Logistic 가중치** | **%.1f%%** | ', rep_scenario.logistic.accuracy);

if rep_scenario.has_existing
    if rep_scenario.logistic.accuracy > rep_scenario.existing.accuracy
        fprintf(fid, '✅ **우수** |\n');
    else
        fprintf(fid, 'ℹ️ 유사 |\n');
    end
    fprintf(fid, '| 기존 종합점수 | %.1f%% | - |\n', rep_scenario.existing.accuracy);
    fprintf(fid, '| **개선 효과** | **%+.1f%%p** | ', ...
        rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy);
    
    if rep_scenario.logistic.accuracy > rep_scenario.existing.accuracy
        fprintf(fid, '**🎉 개선!** |\n\n');
    else
        fprintf(fid, '유사 |\n\n');
    end
else
    fprintf(fid, '- |\n\n');
end

fprintf(fid, '> **분석 기준**: %s 선발 시나리오 (%s)\n\n', rep_scenario.ratio_str, rep_scenario.k_str);

% 2. 예측 성능이란?
fprintf(fid, '---\n\n');
fprintf(fid, '## 1. 예측 성능이란?\n\n');

fprintf(fid, '### 1.1 정확도 (Accuracy)\n\n');
fprintf(fid, '**정확도 = (맞춘 예측 수) / (전체 예측 수) × 100%%**\n\n');
fprintf(fid, '- 합격/불합격을 얼마나 정확하게 맞추는가?\n');
fprintf(fid, '- 높을수록 좋음 (100%%가 최고)\n\n');

fprintf(fid, '### 1.2 혼동행렬 (Confusion Matrix)\n\n');
fprintf(fid, '```\n');
fprintf(fid, '         │ 실제 합격 │ 실제 불합격\n');
fprintf(fid, '─────────┼───────────┼────────────\n');
fprintf(fid, '예측 합격│    TP     │     FP\n');
fprintf(fid, '예측 불합│    FN     │     TN\n');
fprintf(fid, '```\n\n');
fprintf(fid, '- **TP (True Positive)**: 합격을 합격으로 맞춤 ✅\n');
fprintf(fid, '- **TN (True Negative)**: 불합격을 불합격으로 맞춤 ✅\n');
fprintf(fid, '- **FP (False Positive)**: 불합격을 합격으로 오판 ❌\n');
fprintf(fid, '- **FN (False Negative)**: 합격을 불합격으로 오판 ❌\n\n');

fprintf(fid, '**정확도 = (TP + TN) / (TP + TN + FP + FN)**\n\n');

% 3. 분석 결과
fprintf(fid, '---\n\n');
fprintf(fid, '## 2. 분석 결과 상세\n\n');

fprintf(fid, '### 2.1 대표 시나리오: %s 선발\n\n', rep_scenario.ratio_str);

fprintf(fid, '**현재 데이터 (%d명 기준)**\n\n', n_new);
fprintf(fid, '- 선발 인원: %d명 (%s)\n', rep_scenario.k, rep_scenario.k_str);
fprintf(fid, '- 실제 합격자: %d명\n\n', pass_count);

fprintf(fid, '#### Logistic 가중치 예측 결과\n\n');
fprintf(fid, '| 지표 | 값 |\n');
fprintf(fid, '|------|----|\n');
fprintf(fid, '| **정확도** | **%.1f%%** |\n', rep_scenario.logistic.accuracy);
fprintf(fid, '| 정확히 맞춤 | %d명 |\n', rep_scenario.logistic.TP + rep_scenario.logistic.TN);
fprintf(fid, '| 오판 | %d명 |\n', rep_scenario.logistic.FP + rep_scenario.logistic.FN);
fprintf(fid, '| TP (합격→합격) | %d명 |\n', rep_scenario.logistic.TP);
fprintf(fid, '| TN (불합격→불합격) | %d명 |\n', rep_scenario.logistic.TN);
fprintf(fid, '| FP (불합격→합격 오판) | %d명 |\n', rep_scenario.logistic.FP);
fprintf(fid, '| FN (합격→불합격 오판) | %d명 |\n\n', rep_scenario.logistic.FN);

if rep_scenario.has_existing
    fprintf(fid, '#### 기존 종합점수 예측 결과\n\n');
    fprintf(fid, '| 지표 | 값 |\n');
    fprintf(fid, '|------|----|\n');
    fprintf(fid, '| **정확도** | **%.1f%%** |\n', rep_scenario.existing.accuracy);
    fprintf(fid, '| 정확히 맞춤 | %d명 |\n', rep_scenario.existing.TP + rep_scenario.existing.TN);
    fprintf(fid, '| 오판 | %d명 |\n', rep_scenario.existing.FP + rep_scenario.existing.FN);
    fprintf(fid, '| TP (합격→합격) | %d명 |\n', rep_scenario.existing.TP);
    fprintf(fid, '| TN (불합격→불합격) | %d명 |\n', rep_scenario.existing.TN);
    fprintf(fid, '| FP (불합격→합격 오판) | %d명 |\n', rep_scenario.existing.FP);
    fprintf(fid, '| FN (합격→불합격 오판) | %d명 |\n\n', rep_scenario.existing.FN);
    
    fprintf(fid, '#### 비교: 개선 효과\n\n');
    fprintf(fid, '| 지표 | 기존 | Logistic | 개선 |\n');
    fprintf(fid, '|------|------|----------|------|\n');
    fprintf(fid, '| 정확도 | %.1f%% | %.1f%% | **%+.1f%%p** |\n', ...
        rep_scenario.existing.accuracy, rep_scenario.logistic.accuracy, ...
        rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy);
    fprintf(fid, '| 정확히 맞춤 | %d명 | %d명 | **%+d명** |\n', ...
        rep_scenario.existing.TP + rep_scenario.existing.TN, ...
        rep_scenario.logistic.TP + rep_scenario.logistic.TN, ...
        (rep_scenario.logistic.TP + rep_scenario.logistic.TN) - ...
        (rep_scenario.existing.TP + rep_scenario.existing.TN));
    fprintf(fid, '| 오판 | %d명 | %d명 | **%d명 감소** |\n\n', ...
        rep_scenario.existing.FP + rep_scenario.existing.FN, ...
        rep_scenario.logistic.FP + rep_scenario.logistic.FN, ...
        (rep_scenario.existing.FP + rep_scenario.existing.FN) - ...
        (rep_scenario.logistic.FP + rep_scenario.logistic.FN));
    
    if rep_scenario.logistic.accuracy > rep_scenario.existing.accuracy
        fprintf(fid, '✅ **Logistic 가중치가 %.1f%%p 더 정확합니다!**\n\n', ...
            rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy);
    end
end

% 4. 실무 확장
fprintf(fid, '---\n\n');
fprintf(fid, '## 3. 실무 확장 시뮬레이션\n\n');

fprintf(fid, '### 3.1 만약 100명이 지원한다면?\n\n');

k_scaled = round(rep_scenario.k * 100 / n_new);
fprintf(fid, '**시나리오**: %s 선발 (%d명)\n\n', rep_scenario.ratio_str, k_scaled);

fprintf(fid, '#### Logistic 가중치 사용 시\n\n');
fprintf(fid, '- 정확도 %.1f%% 기준\n', rep_scenario.logistic.accuracy);
fprintf(fid, '- 약 **%d명 정확히 예측**\n', round(100 * rep_scenario.logistic.accuracy / 100));
fprintf(fid, '- 오판: 약 %d명\n\n', round(100 * (100 - rep_scenario.logistic.accuracy) / 100));

if rep_scenario.has_existing
    fprintf(fid, '#### 기존 종합점수 사용 시\n\n');
    fprintf(fid, '- 정확도 %.1f%% 기준\n', rep_scenario.existing.accuracy);
    fprintf(fid, '- 약 %d명 정확히 예측\n', round(100 * rep_scenario.existing.accuracy / 100));
    fprintf(fid, '- 오판: 약 %d명\n\n', round(100 * (100 - rep_scenario.existing.accuracy) / 100));
    
    improvement = round(100 * (rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy) / 100);
    if improvement > 0
        fprintf(fid, '#### 개선 효과\n\n');
        fprintf(fid, '✅ **약 %d명 더 정확하게 예측!**\n\n', improvement);
    end
end

% 5. 다양한 선발 비율별 결과
fprintf(fid, '---\n\n');
fprintf(fid, '## 4. 선발 비율별 예측 성능\n\n');

fprintf(fid, '| 선발 비율 | 선발 인원 | Logistic 정확도 | 기존 정확도 | 개선 |\n');
fprintf(fid, '|-----------|-----------|-----------------|-------------|------|\n');

for i = 1:length(selection_scenarios)
    sc = selection_scenarios{i};
    if sc.has_existing
        fprintf(fid, '| %s | %s | %.1f%% | %.1f%% | %+.1f%%p |\n', ...
            sc.ratio_str, sc.k_str, sc.logistic.accuracy, ...
            sc.existing.accuracy, sc.logistic.accuracy - sc.existing.accuracy);
    else
        fprintf(fid, '| %s | %s | %.1f%% | N/A | N/A |\n', ...
            sc.ratio_str, sc.k_str, sc.logistic.accuracy);
    end
end
fprintf(fid, '\n');

% 6. 결론 및 권장사항
fprintf(fid, '---\n\n');
fprintf(fid, '## 5. 결론 및 권장사항\n\n');

fprintf(fid, '### 5.1 핵심 결론\n\n');

if rep_scenario.has_existing && rep_scenario.logistic.accuracy > rep_scenario.existing.accuracy
    fprintf(fid, '✅ **Logistic 가중치 사용 강력 권장**\n\n');
    fprintf(fid, '**이유**:\n\n');
    fprintf(fid, '1. **더 높은 정확도**: %.1f%% vs %.1f%% (%+.1f%%p 개선)\n', ...
        rep_scenario.logistic.accuracy, rep_scenario.existing.accuracy, ...
        rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy);
    fprintf(fid, '2. **오판 감소**: %d명 → %d명\n', ...
        rep_scenario.existing.FP + rep_scenario.existing.FN, ...
        rep_scenario.logistic.FP + rep_scenario.logistic.FN);
    fprintf(fid, '3. **실무 효과**: 100명 선발 시 약 %d명 더 정확\n\n', ...
        round(100 * (rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy) / 100));
else
    fprintf(fid, 'ℹ️ **두 방법 병행 검토**\n\n');
    fprintf(fid, '- 유사한 예측 성능\n');
    fprintf(fid, '- 다른 관점의 평가 가능\n\n');
end

fprintf(fid, '### 5.2 선발 비율별 권장사항\n\n');

best_ratio_idx = 1;
best_accuracy = selection_scenarios{1}.logistic.accuracy;
for i = 2:length(selection_scenarios)
    if selection_scenarios{i}.logistic.accuracy > best_accuracy
        best_accuracy = selection_scenarios{i}.logistic.accuracy;
        best_ratio_idx = i;
    end
end

fprintf(fid, '- **최고 정확도**: %s 선발 시 %.1f%% ⭐\n', ...
    selection_scenarios{best_ratio_idx}.ratio_str, best_accuracy);

for i = 1:length(selection_scenarios)
    sc = selection_scenarios{i};
    if sc.logistic.accuracy >= 80
        fprintf(fid, '- **%s 선발**: 정확도 %.1f%% (매우 우수)\n', sc.ratio_str, sc.logistic.accuracy);
    elseif sc.logistic.accuracy >= 70
        fprintf(fid, '- **%s 선발**: 정확도 %.1f%% (우수)\n', sc.ratio_str, sc.logistic.accuracy);
    end
end
fprintf(fid, '\n');

fprintf(fid, '### 5.3 주의사항\n\n');
fprintf(fid, '1. **소규모 데이터**: %d명 분석으로 추가 검증 필요\n', n_new);
fprintf(fid, '2. **지속적 모니터링**: 실제 적용 후 성능 추적\n');
fprintf(fid, '3. **다면적 평가**: 정량 지표 외 정성적 요소 고려\n');
fprintf(fid, '4. **상황별 조정**: 선발 목적에 따라 비율 조정 필요\n\n');

% 7. 부록
fprintf(fid, '---\n\n');
fprintf(fid, '## 6. 부록\n\n');

fprintf(fid, '### 6.1 주요 용어 정리\n\n');
fprintf(fid, '- **정확도 (Accuracy)**: 전체 예측 중 맞춘 비율\n');
fprintf(fid, '- **정밀도 (Precision)**: 합격 예측 중 실제 합격 비율\n');
fprintf(fid, '- **재현율 (Recall)**: 실제 합격자 중 찾아낸 비율\n');
fprintf(fid, '- **F1 Score**: 정밀도와 재현율의 조화평균\n');
fprintf(fid, '- **TP/FP/FN/TN**: 혼동행렬 요소 (위 참조)\n\n');

fprintf(fid, '### 6.2 출력 파일\n\n');
fprintf(fid, '1. **엑셀**: `%s`\n', config.output_filename);
fprintf(fid, '   - 성능비교_요약: 선발 비율별 성능 지표\n');
fprintf(fid, '   - 혼동행렬: TP/FP/FN/TN 상세\n');
fprintf(fid, '   - 개인별_점수_예측: ID별 점수 및 예측 결과\n\n');

fprintf(fid, '2. **시각화**: `%s`\n', config.plot_filename);
fprintf(fid, '   - 정확도/정밀도/재현율 비교 그래프\n');
fprintf(fid, '   - 혼동행렬 히트맵\n');
fprintf(fid, '   - 개선 효과 시각화\n\n');

fprintf(fid, '---\n\n');
fprintf(fid, '*본 리포트는 MATLAB 자동 분석으로 생성되었습니다.*\n\n');
fprintf(fid, '*생성일시: %s*\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

fclose(fid);
fprintf('  ✓ 마크다운 리포트 저장: %s\n', config.report_filename);

%% 12) 최종 리포트 (콘솔) ------------------------------------------------
fprintf('\n');
fprintf('================================================================\n');
fprintf('              🎯 예측 성능 검증 최종 리포트 🎯\n');
fprintf('================================================================\n\n');

fprintf('【 핵심 질문 】\n');
fprintf('  "새 가중치로 예측하면 기존보다 얼마나 더 정확한가?"\n\n');

fprintf('【 분석 데이터 】\n');
fprintf('  • 총 %d명 (합격 %d명, 불합격 %d명)\n', n_new, pass_count, fail_count);
fprintf('  • 전체 합격률: %.1f%%\n\n', baseline_rate);

fprintf('【 핵심 결과 】\n\n');
fprintf('  대표 시나리오: %s 선발 (%s)\n\n', rep_scenario.ratio_str, rep_scenario.k_str);

fprintf('  ┌─────────────────┬──────────────┬──────────────┬─────────┐\n');
fprintf('  │      지표       │   Logistic   │   기존 점수  │  개선   │\n');
fprintf('  ├─────────────────┼──────────────┼──────────────┼─────────┤\n');

if rep_scenario.has_existing
    fprintf('  │ 정확도          │    %.1f%%     │    %.1f%%     │ %+.1f%%p  │\n', ...
        rep_scenario.logistic.accuracy, rep_scenario.existing.accuracy, ...
        rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy);
    fprintf('  │ 정확히 맞춤     │    %2d명      │    %2d명      │ %+2d명   │\n', ...
        rep_scenario.logistic.TP + rep_scenario.logistic.TN, ...
        rep_scenario.existing.TP + rep_scenario.existing.TN, ...
        (rep_scenario.logistic.TP + rep_scenario.logistic.TN) - ...
        (rep_scenario.existing.TP + rep_scenario.existing.TN));
    fprintf('  │ 오판            │    %2d명      │    %2d명      │ %2d명↓  │\n', ...
        rep_scenario.logistic.FP + rep_scenario.logistic.FN, ...
        rep_scenario.existing.FP + rep_scenario.existing.FN, ...
        (rep_scenario.existing.FP + rep_scenario.existing.FN) - ...
        (rep_scenario.logistic.FP + rep_scenario.logistic.FN));
else
    fprintf('  │ 정확도          │    %.1f%%     │      N/A     │   N/A   │\n', ...
        rep_scenario.logistic.accuracy);
    fprintf('  │ 정확히 맞춤     │    %2d명      │      N/A     │   N/A   │\n', ...
        rep_scenario.logistic.TP + rep_scenario.logistic.TN);
end

fprintf('  └─────────────────┴──────────────┴──────────────┴─────────┘\n\n');

fprintf('【 실무 확장 (100명 지원 시) 】\n\n');
k_scaled = round(rep_scenario.k * 100 / n_new);
fprintf('  • 선발 인원: %d명 (%s)\n', k_scaled, rep_scenario.ratio_str);
fprintf('  • Logistic: 약 %d명 정확 예측\n', round(100 * rep_scenario.logistic.accuracy / 100));

if rep_scenario.has_existing
    fprintf('  • 기존: 약 %d명 정확 예측\n', round(100 * rep_scenario.existing.accuracy / 100));
    improvement = round(100 * (rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy) / 100);
    if improvement > 0
        fprintf('  • 개선: 약 %d명 더 정확! 🎉\n\n', improvement);
    else
        fprintf('\n');
    end
end

fprintf('【 결론 】\n');

if rep_scenario.has_existing && rep_scenario.logistic.accuracy > rep_scenario.existing.accuracy
    fprintf('  ✅ Logistic 가중치가 더 정확한 예측!\n');
    fprintf('     %.1f%%p 개선 (%.1f%% → %.1f%%)\n\n', ...
        rep_scenario.logistic.accuracy - rep_scenario.existing.accuracy, ...
        rep_scenario.existing.accuracy, rep_scenario.logistic.accuracy);
else
    fprintf('  ℹ️ 두 방법이 유사한 예측 성능\n\n');
end

fprintf('【 출력 파일 】\n');
fprintf('  • 엑셀: %s\n', config.output_filename);
fprintf('  • 그래프: %s\n', config.plot_filename);
fprintf('  • 리포트: %s\n\n', config.report_filename);

fprintf('================================================================\n');
fprintf('  완료: %s\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));
fprintf('  위치: %s\n', config.output_dir);
fprintf('================================================================\n\n');

fprintf('💡 핵심 메시지:\n');
fprintf('   • 정확도가 높을수록 합격/불합격을 정확하게 예측\n');
fprintf('   • 오판이 적을수록 좋은 모델\n');
fprintf('   • 실무에서는 선발 목적에 맞는 비율 선택\n\n');