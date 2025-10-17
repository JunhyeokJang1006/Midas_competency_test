%% ========================================================================
%         원점수 vs 가중치 점수 인재유형 예측 성능 비교 분석
% =========================================================================
% 목적: 실제 데이터를 기반으로 원점수와 가중치 적용 점수의
%       인재유형 예측 정확도를 비교
%
% 입력: 역량검사_가중치적용점수_talent.xlsx
% 출력: 예측 성능 비교 보고서 및 시각화
% =========================================================================

clear; clc; close all;
rng(42, 'twister');

%% ========================================================================
%                          STEP 1: 데이터 로딩
% =========================================================================

fprintf('【STEP 1】 데이터 로딩\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 파일 경로 설정
filepath = 'D:\project\HR데이터\결과\최종\2025.10.14\역량검사_가중치적용점수_talent_2025-10-14_185545.xlsx';
output_dir = 'D:\project\HR데이터\결과\자가불소_revised_talent';

% 출력 디렉토리 생성
if ~exist(output_dir, 'dir')
    mkdir(output_dir);
end

% 데이터 로드
data = readtable(filepath, 'Sheet', 1, 'VariableNamingRule', 'preserve');

fprintf('데이터 로드 완료: %d명\n', height(data));
fprintf('컬럼: ID, 원점수, 가중치점수, 변화율, 원예측유형, 가중예측유형\n\n');

%% ========================================================================
%                    STEP 2: 데이터 전처리 및 레이블링
% =========================================================================

fprintf('【STEP 2】 데이터 전처리 및 레이블링\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 컬럼 추출
original_score = data{:, 2};  % 원점수
weighted_score = data{:, 3};  % 가중치 점수
original_type = data{:, 5};   % 원점수 기반 예측 유형 (상세)
weighted_type = data{:, 6};   % 가중치 점수 기반 예측 레이블 (간략)

% 빈 셀 처리
if iscell(original_type)
    original_type_clean = original_type;
    for i = 1:length(original_type)
        if isempty(original_type{i})
            original_type_clean{i} = '알수없음';
        end
    end
else
    original_type_clean = original_type;
end

if iscell(weighted_type)
    weighted_type_clean = weighted_type;
    for i = 1:length(weighted_type)
        if isempty(weighted_type{i})
            weighted_type_clean{i} = '알수없음';
        end
    end
else
    weighted_type_clean = weighted_type;
end

% 고성과자/저성과자 정의 (상세 유형 기준)
high_performance_types = {'성실한 가연성', '자연성', '유익한 불연성'};
low_performance_types = {'게으른 가연성', '무능한 불연성', '소화성', '위장형 소화성'};
excluded_types = {'유능한 불연성', '알수없음', ''};

% 원점수 기반 예측의 이진 레이블 생성
n_samples = length(original_type_clean);
original_prediction = zeros(n_samples, 1);  % 원점수 기반 예측

for i = 1:n_samples
    current_type = original_type_clean{i};

    if any(strcmp(current_type, high_performance_types))
        original_prediction(i) = 1;  % 고성과자
    elseif any(strcmp(current_type, low_performance_types))
        original_prediction(i) = 0;  % 저성과자
    else
        original_prediction(i) = -1;  % 중간 또는 제외 (분석 제외)
    end
end

% 가중치 점수 기반 예측의 이진 레이블 생성 (간략 레이블 사용)
% '탁월', '우수' -> 고성과자
% '저성과' -> 저성과자
% '보통' -> 중간 (분석 제외)
weighted_prediction = zeros(n_samples, 1);

for i = 1:n_samples
    current_label = weighted_type_clean{i};

    if any(strcmp(current_label, {'탁월', '우수'}))
        weighted_prediction(i) = 1;  % 고성과자
    elseif strcmp(current_label, '저성과')
        weighted_prediction(i) = 0;  % 저성과자
    else
        weighted_prediction(i) = -1;  % 중간 (분석 제외)
    end
end

% Ground Truth: 가중치 점수 기반 예측을 정답으로 사용
true_label = weighted_prediction;

% 유효한 데이터만 선택 (중간 그룹 제외)
valid_idx = (true_label ~= -1) & ~isnan(original_score) & ~isnan(weighted_score);

fprintf('필터링 전 데이터: %d명\n', n_samples);
fprintf('  고성과자 레이블: %d명\n', sum(true_label == 1));
fprintf('  저성과자 레이블: %d명\n', sum(true_label == 0));
fprintf('  중간/알수없음: %d명\n', sum(true_label == -1));
fprintf('  원점수 NaN: %d명\n', sum(isnan(original_score)));
fprintf('  가중치점수 NaN: %d명\n', sum(isnan(weighted_score)));
fprintf('  유효 인덱스: %d명\n\n', sum(valid_idx));

original_score_valid = original_score(valid_idx);
weighted_score_valid = weighted_score(valid_idx);
original_type_valid = original_type_clean(valid_idx);
weighted_type_valid = weighted_type_clean(valid_idx);
true_label_valid = true_label(valid_idx);

n_valid = length(true_label_valid);
n_high = sum(true_label_valid == 1);
n_low = sum(true_label_valid == 0);

fprintf('유효 데이터: %d명 (전체의 %.1f%%)\n', n_valid, n_valid/n_samples*100);
fprintf('  고성과자: %d명 (%.1f%%)\n', n_high, n_high/n_valid*100);
fprintf('  저성과자: %d명 (%.1f%%)\n\n', n_low, n_low/n_valid*100);

%% ========================================================================
%                STEP 3: Top-K 예측 정확도 비교 (핵심 지표)
% =========================================================================

fprintf('【STEP 3】 Top-K 예측 정확도 비교\n');
fprintf('════════════════════════════════════════════════════════════\n');

top_k_percentages = [5, 10, 20, 30, 50];

fprintf('\n┌─────────────────────────────────────────────────────────┐\n');
fprintf('│           Top-K 고성과자 선발 정확도 비교               │\n');
fprintf('├─────────────────────────────────────────────────────────┤\n');
fprintf('│ 상위 K%%  │  원점수  │ 가중치점수 │ 개선(%%) │ 개선(%%p) │\n');
fprintf('├─────────────────────────────────────────────────────────┤\n');

top_k_original_all = zeros(size(top_k_percentages));
top_k_weighted_all = zeros(size(top_k_percentages));

for idx = 1:length(top_k_percentages)
    k_pct = top_k_percentages(idx);
    k = round(n_valid * k_pct / 100);

    % 원점수 기준 상위 K명
    [~, idx_original] = sort(original_score_valid, 'descend');
    top_k_original = idx_original(1:k);
    accuracy_original = sum(true_label_valid(top_k_original) == 1) / k * 100;
    top_k_original_all(idx) = accuracy_original;

    % 가중치 점수 기준 상위 K명
    [~, idx_weighted] = sort(weighted_score_valid, 'descend');
    top_k_weighted = idx_weighted(1:k);
    accuracy_weighted = sum(true_label_valid(top_k_weighted) == 1) / k * 100;
    top_k_weighted_all(idx) = accuracy_weighted;

    improvement_pct = (accuracy_weighted - accuracy_original) / accuracy_original * 100;
    improvement_pp = accuracy_weighted - accuracy_original;

    fprintf('│ 상위 %2d%% │ %6.1f%% │  %6.1f%%   │  %+5.1f%% │  %+5.1f%%p │\n', ...
        k_pct, accuracy_original, accuracy_weighted, improvement_pct, improvement_pp);
end

fprintf('└─────────────────────────────────────────────────────────┘\n\n');

%% ========================================================================
%              STEP 4: Lift Score 분석 (랜덤 선발 대비)
% =========================================================================

fprintf('【STEP 4】 Lift Score 분석\n');
fprintf('════════════════════════════════════════════════════════════\n');

base_rate = n_high / n_valid * 100;

fprintf('전체 고성과자 비율: %.1f%%\n', base_rate);
fprintf('Lift Score = (Top-K 정확도) / (전체 고성과자 비율)\n\n');

fprintf('┌─────────────────────────────────────────────────────┐\n');
fprintf('│           Lift Score (랜덤 선발 대비)               │\n');
fprintf('├─────────────────────────────────────────────────────┤\n');
fprintf('│ 상위 K%%  │  원점수  │ 가중치점수 │  개선효과   │\n');
fprintf('├─────────────────────────────────────────────────────┤\n');

for idx = 1:length(top_k_percentages)
    k_pct = top_k_percentages(idx);
    accuracy_original = top_k_original_all(idx);
    accuracy_weighted = top_k_weighted_all(idx);

    lift_original = accuracy_original / base_rate;
    lift_weighted = accuracy_weighted / base_rate;

    fprintf('│ 상위 %2d%% │  %.2fx   │   %.2fx    │  +%.2fx     │\n', ...
        k_pct, lift_original, lift_weighted, lift_weighted - lift_original);
end

fprintf('└─────────────────────────────────────────────────────┘\n\n');

%% ========================================================================
%           STEP 5: 실제 채용 시나리오 시뮬레이션
% =========================================================================

fprintf('【STEP 5】 실제 채용 시나리오 시뮬레이션\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 채용 시나리오: 100명 중 10명 선발
n_applicants = 100;
n_select = 10;
n_simulations = 5000;

fprintf('시나리오: %d명 지원자 중 %d명 선발\n', n_applicants, n_select);
fprintf('시뮬레이션 횟수: %d회\n\n', n_simulations);

% 실제 데이터 기반 상관계수 계산
corr_original = corr(original_score_valid, double(true_label_valid), 'Type', 'Spearman');
corr_weighted = corr(weighted_score_valid, double(true_label_valid), 'Type', 'Spearman');

fprintf('점수와 고성과자 레이블의 상관계수:\n');
fprintf('  원점수:     %.4f\n', corr_original);
fprintf('  가중치점수: %.4f\n', corr_weighted);
fprintf('  상관계수 개선: %.4f (%.1f%% 증가)\n\n', ...
    corr_weighted - corr_original, (corr_weighted - corr_original) / abs(corr_original) * 100);

% 시뮬레이션
prob_high = n_high / n_valid;

original_high_count = zeros(n_simulations, 1);
weighted_high_count = zeros(n_simulations, 1);
random_high_count = zeros(n_simulations, 1);

for sim = 1:n_simulations
    % 가상 지원자 생성
    n_high_sim = round(n_applicants * prob_high);
    n_low_sim = n_applicants - n_high_sim;
    sim_labels = [ones(n_high_sim, 1); zeros(n_low_sim, 1)];

    % 가상 점수 생성 (정규분포)
    sim_original = randn(n_applicants, 1);
    sim_weighted = randn(n_applicants, 1);

    % 고성과자에게 더 높은 점수 부여 (실제 상관계수 반영)
    high_idx_sim = find(sim_labels == 1);

    effect_original = corr_original * 2.0;  % 효과 크기
    effect_weighted = corr_weighted * 2.0;

    sim_original(high_idx_sim) = sim_original(high_idx_sim) + effect_original;
    sim_weighted(high_idx_sim) = sim_weighted(high_idx_sim) + effect_weighted;

    % 상위 n_select명 선발
    [~, idx_o] = sort(sim_original, 'descend');
    [~, idx_w] = sort(sim_weighted, 'descend');

    original_high_count(sim) = sum(sim_labels(idx_o(1:n_select)) == 1);
    weighted_high_count(sim) = sum(sim_labels(idx_w(1:n_select)) == 1);

    % 랜덤 선발
    random_idx = randperm(n_applicants, n_select);
    random_high_count(sim) = sum(sim_labels(random_idx) == 1);
end

fprintf('┌──────────────────────────────────────────────────┐\n');
fprintf('│        선발된 고성과자 수 (평균 ± 표준편차)     │\n');
fprintf('├──────────────────────────────────────────────────┤\n');
fprintf('│ 랜덤 선발:      %.2f ± %.2f 명                │\n', ...
    mean(random_high_count), std(random_high_count));
fprintf('│ 원점수:         %.2f ± %.2f 명                │\n', ...
    mean(original_high_count), std(original_high_count));
fprintf('│ 가중치점수:     %.2f ± %.2f 명                │\n', ...
    mean(weighted_high_count), std(weighted_high_count));
fprintf('└──────────────────────────────────────────────────┘\n\n');

fprintf('【개선 효과】\n');
fprintf('────────────────────────────────────────────────────\n');
improvement_vs_random = (mean(weighted_high_count) - mean(random_high_count)) / mean(random_high_count) * 100;
improvement_vs_original = (mean(weighted_high_count) - mean(original_high_count)) / mean(original_high_count) * 100;

fprintf('가중치점수 vs 랜덤:   평균 +%.2f명 (%.1f%% 개선)\n', ...
    mean(weighted_high_count) - mean(random_high_count), improvement_vs_random);
fprintf('가중치점수 vs 원점수: 평균 +%.2f명 (%.1f%% 개선)\n\n', ...
    mean(weighted_high_count) - mean(original_high_count), improvement_vs_original);

%% ========================================================================
%                    STEP 6: ROC/AUC 비교
% =========================================================================

fprintf('【STEP 6】 ROC/AUC 분석\n');
fprintf('════════════════════════════════════════════════════════════\n');

[X_roc_original, Y_roc_original, ~, AUC_original] = perfcurve(true_label_valid, original_score_valid, 1);
[X_roc_weighted, Y_roc_weighted, ~, AUC_weighted] = perfcurve(true_label_valid, weighted_score_valid, 1);

fprintf('┌────────────────────────────────────────┐\n');
fprintf('│          ROC AUC 비교                  │\n');
fprintf('├────────────────────────────────────────┤\n');
fprintf('│ 원점수 AUC:        %.4f            │\n', AUC_original);
fprintf('│ 가중치점수 AUC:    %.4f            │\n', AUC_weighted);
fprintf('│ AUC 개선:          +%.4f (%.1f%%)  │\n', ...
    AUC_weighted - AUC_original, (AUC_weighted - AUC_original) / AUC_original * 100);
fprintf('└────────────────────────────────────────┘\n\n');

%% ========================================================================
%                    STEP 7: 시각화
% =========================================================================

fprintf('【STEP 7】 결과 시각화\n');
fprintf('════════════════════════════════════════════════════════════\n');

% Figure 1: 종합 비교 대시보드
fig1 = figure('Position', [50, 50, 1600, 900]);
set(gcf, 'Color', 'w');

% 서브플롯 1: Top-K 정확도 비교
subplot(2, 3, 1);
plot(top_k_percentages, top_k_original_all, '-o', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'Color', [0.3, 0.5, 0.8], 'MarkerFaceColor', [0.3, 0.5, 0.8]);
hold on;
plot(top_k_percentages, top_k_weighted_all, '-s', 'LineWidth', 2.5, 'MarkerSize', 10, ...
    'Color', [0.9, 0.3, 0.3], 'MarkerFaceColor', [0.9, 0.3, 0.3]);
yline(base_rate, '--k', sprintf('전체 고성과자 비율 (%.1f%%)', base_rate), ...
    'LineWidth', 1.5, 'FontSize', 9);
xlabel('상위 K% 선발', 'FontSize', 11, 'FontWeight', 'bold');
ylabel('고성과자 비율 (%)', 'FontSize', 11, 'FontWeight', 'bold');
title('Top-K 정확도 비교', 'FontSize', 13, 'FontWeight', 'bold');
legend('원점수', '가중치점수', 'Location', 'best', 'FontSize', 10);
grid on; box on;
set(gca, 'FontSize', 10);

% 서브플롯 2: 개선 효과
subplot(2, 3, 2);
improvement_all = top_k_weighted_all - top_k_original_all;
improvement_pct_all = (top_k_weighted_all - top_k_original_all) ./ top_k_original_all * 100;
yyaxis left;
bar(top_k_percentages, improvement_all, 'FaceColor', [0.2, 0.7, 0.4], 'EdgeColor', 'k');
ylabel('개선율 (%p)', 'FontSize', 11, 'FontWeight', 'bold');
yyaxis right;
plot(top_k_percentages, improvement_pct_all, '-o', 'LineWidth', 2, 'MarkerSize', 8, ...
    'Color', [0.8, 0.4, 0], 'MarkerFaceColor', [0.8, 0.4, 0]);
ylabel('개선율 (%)', 'FontSize', 11, 'FontWeight', 'bold');
xlabel('상위 K% 선발', 'FontSize', 11, 'FontWeight', 'bold');
title('가중치점수의 개선 효과', 'FontSize', 13, 'FontWeight', 'bold');
grid on; box on;
set(gca, 'FontSize', 10);

% 서브플롯 3: Lift Score 비교
subplot(2, 3, 3);
lift_original_all = top_k_original_all / base_rate;
lift_weighted_all = top_k_weighted_all / base_rate;
bar_data = [lift_original_all', lift_weighted_all'];
b = bar(bar_data, 'grouped');
b(1).FaceColor = [0.3, 0.5, 0.8];
b(2).FaceColor = [0.9, 0.3, 0.3];
set(gca, 'XTickLabel', arrayfun(@(x) sprintf('%d%%', x), top_k_percentages, 'UniformOutput', false));
xlabel('상위 K% 선발', 'FontSize', 11, 'FontWeight', 'bold');
ylabel('Lift Score', 'FontSize', 11, 'FontWeight', 'bold');
title('Lift Score 비교 (랜덤 선발 대비)', 'FontSize', 13, 'FontWeight', 'bold');
legend('원점수', '가중치점수', 'Location', 'best', 'FontSize', 10);
grid on; box on;
set(gca, 'FontSize', 10);

% 서브플롯 4: ROC Curve
subplot(2, 3, 4);
plot(X_roc_original, Y_roc_original, '-', 'LineWidth', 2.5, 'Color', [0.3, 0.5, 0.8]);
hold on;
plot(X_roc_weighted, Y_roc_weighted, '-', 'LineWidth', 2.5, 'Color', [0.9, 0.3, 0.3]);
plot([0, 1], [0, 1], '--k', 'LineWidth', 1.5);
xlabel('False Positive Rate', 'FontSize', 11, 'FontWeight', 'bold');
ylabel('True Positive Rate', 'FontSize', 11, 'FontWeight', 'bold');
title('ROC Curve', 'FontSize', 13, 'FontWeight', 'bold');
legend(sprintf('원점수 (AUC=%.3f)', AUC_original), ...
       sprintf('가중치점수 (AUC=%.3f)', AUC_weighted), ...
       'Random', 'Location', 'southeast', 'FontSize', 10);
grid on; box on;
set(gca, 'FontSize', 10);
axis square;

% 서브플롯 5: 채용 시뮬레이션 결과
subplot(2, 3, 5);
methods = {'랜덤', '원점수', '가중치점수'};
mean_counts = [mean(random_high_count), mean(original_high_count), mean(weighted_high_count)];
std_counts = [std(random_high_count), std(original_high_count), std(weighted_high_count)];
b = bar(mean_counts);
b.FaceColor = [0.5, 0.5, 0.8];
b.EdgeColor = 'k';
hold on;
errorbar(1:3, mean_counts, std_counts, 'k.', 'LineWidth', 2, 'CapSize', 12);
set(gca, 'XTickLabel', methods);
ylabel(sprintf('선발된 고성과자 수 (/%d명)', n_select), 'FontSize', 11, 'FontWeight', 'bold');
title(sprintf('채용 시뮬레이션 (%d명 중 %d명 선발)', n_applicants, n_select), ...
    'FontSize', 13, 'FontWeight', 'bold');
grid on; box on;
set(gca, 'FontSize', 10);
ylim([0, max(mean_counts) * 1.3]);
for i = 1:3
    text(i, mean_counts(i) + std_counts(i) + 0.15, ...
        sprintf('%.2f±%.2f', mean_counts(i), std_counts(i)), ...
        'HorizontalAlignment', 'center', 'FontSize', 9, 'FontWeight', 'bold');
end

% 서브플롯 6: 점수 분포 비교
subplot(2, 3, 6);
hold on;
% 고성과자
high_idx = (true_label_valid == 1);
histogram(original_score_valid(high_idx), 20, 'FaceColor', [0.3, 0.5, 0.8], 'FaceAlpha', 0.5, ...
    'EdgeColor', 'k', 'DisplayName', '원점수(고성과자)');
histogram(weighted_score_valid(high_idx), 20, 'FaceColor', [0.9, 0.3, 0.3], 'FaceAlpha', 0.5, ...
    'EdgeColor', 'k', 'DisplayName', '가중점수(고성과자)');
xlabel('점수', 'FontSize', 11, 'FontWeight', 'bold');
ylabel('빈도', 'FontSize', 11, 'FontWeight', 'bold');
title('고성과자 점수 분포', 'FontSize', 13, 'FontWeight', 'bold');
legend('Location', 'best', 'FontSize', 10);
grid on; box on;
set(gca, 'FontSize', 10);

sgtitle('원점수 vs 가중치점수 성과 비교 종합', 'FontSize', 16, 'FontWeight', 'bold');

% 저장
saveas(fig1, fullfile(output_dir, 'score_comparison_dashboard.png'));
fprintf('시각화 저장: score_comparison_dashboard.png\n');

%% ========================================================================
%                    STEP 8: 상세 분석 결과 저장
% =========================================================================

fprintf('\n【STEP 8】 결과 저장\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 결과 구조체 생성
results = struct();
results.sample_info = struct(...
    'total_samples', n_samples, ...
    'valid_samples', n_valid, ...
    'high_performers', n_high, ...
    'low_performers', n_low, ...
    'base_rate_pct', base_rate);

results.top_k = struct(...
    'percentages', top_k_percentages, ...
    'original_accuracy', top_k_original_all, ...
    'weighted_accuracy', top_k_weighted_all, ...
    'improvement_pp', improvement_all, ...
    'improvement_pct', improvement_pct_all);

results.lift_score = struct(...
    'original', lift_original_all, ...
    'weighted', lift_weighted_all);

results.correlation = struct(...
    'original_spearman', corr_original, ...
    'weighted_spearman', corr_weighted);

results.auc = struct(...
    'original', AUC_original, ...
    'weighted', AUC_weighted, ...
    'improvement', AUC_weighted - AUC_original);

results.simulation = struct(...
    'n_applicants', n_applicants, ...
    'n_select', n_select, ...
    'n_simulations', n_simulations, ...
    'random_mean', mean(random_high_count), ...
    'random_std', std(random_high_count), ...
    'original_mean', mean(original_high_count), ...
    'original_std', std(original_high_count), ...
    'weighted_mean', mean(weighted_high_count), ...
    'weighted_std', std(weighted_high_count), ...
    'improvement_vs_random_pct', improvement_vs_random, ...
    'improvement_vs_original_pct', improvement_vs_original);

% MAT 파일로 저장
save(fullfile(output_dir, 'score_comparison_results.mat'), 'results');
fprintf('결과 저장: score_comparison_results.mat\n');

% 텍스트 보고서 생성
report_file = fullfile(output_dir, 'score_comparison_report.txt');
fid = fopen(report_file, 'w', 'n', 'UTF-8');

fprintf(fid, '================================================================\n');
fprintf(fid, '       원점수 vs 가중치점수 인재유형 예측 성능 비교 보고서\n');
fprintf(fid, '================================================================\n');
fprintf(fid, '분석 일시: %s\n\n', datestr(now, 'yyyy-mm-dd HH:MM:SS'));

fprintf(fid, '【1. 데이터 요약】\n');
fprintf(fid, '────────────────────────────────────────────────────────\n');
fprintf(fid, '총 샘플 수:        %d명\n', n_samples);
fprintf(fid, '유효 샘플 수:      %d명 (%.1f%%)\n', n_valid, n_valid/n_samples*100);
fprintf(fid, '고성과자:          %d명 (%.1f%%)\n', n_high, n_high/n_valid*100);
fprintf(fid, '저성과자:          %d명 (%.1f%%)\n\n', n_low, n_low/n_valid*100);

fprintf(fid, '【2. Top-K 정확도 비교】\n');
fprintf(fid, '────────────────────────────────────────────────────────\n');
for idx = 1:length(top_k_percentages)
    fprintf(fid, '상위 %2d%%: 원점수 %.1f%%, 가중치 %.1f%% (+%.1f%%p, +%.1f%%)\n', ...
        top_k_percentages(idx), top_k_original_all(idx), top_k_weighted_all(idx), ...
        improvement_all(idx), improvement_pct_all(idx));
end
fprintf(fid, '\n');

fprintf(fid, '【3. Lift Score (랜덤 선발 대비)】\n');
fprintf(fid, '────────────────────────────────────────────────────────\n');
fprintf(fid, '전체 고성과자 비율: %.1f%%\n', base_rate);
for idx = 1:length(top_k_percentages)
    fprintf(fid, '상위 %2d%%: 원점수 %.2fx, 가중치 %.2fx (+%.2fx)\n', ...
        top_k_percentages(idx), lift_original_all(idx), lift_weighted_all(idx), ...
        lift_weighted_all(idx) - lift_original_all(idx));
end
fprintf(fid, '\n');

fprintf(fid, '【4. ROC AUC】\n');
fprintf(fid, '────────────────────────────────────────────────────────\n');
fprintf(fid, '원점수 AUC:        %.4f\n', AUC_original);
fprintf(fid, '가중치점수 AUC:    %.4f\n', AUC_weighted);
fprintf(fid, 'AUC 개선:          +%.4f (%.1f%%)\n\n', ...
    AUC_weighted - AUC_original, (AUC_weighted - AUC_original) / AUC_original * 100);

fprintf(fid, '【5. 채용 시뮬레이션 (%d명 중 %d명 선발)】\n', n_applicants, n_select);
fprintf(fid, '────────────────────────────────────────────────────────\n');
fprintf(fid, '랜덤 선발:         %.2f ± %.2f 명\n', mean(random_high_count), std(random_high_count));
fprintf(fid, '원점수:            %.2f ± %.2f 명\n', mean(original_high_count), std(original_high_count));
fprintf(fid, '가중치점수:        %.2f ± %.2f 명\n', mean(weighted_high_count), std(weighted_high_count));
fprintf(fid, '\n개선 효과:\n');
fprintf(fid, '  vs 랜덤:   +%.2f명 (%.1f%%)\n', ...
    mean(weighted_high_count) - mean(random_high_count), improvement_vs_random);
fprintf(fid, '  vs 원점수: +%.2f명 (%.1f%%)\n\n', ...
    mean(weighted_high_count) - mean(original_high_count), improvement_vs_original);

fprintf(fid, '【6. 상관계수】\n');
fprintf(fid, '────────────────────────────────────────────────────────\n');
fprintf(fid, '원점수 Spearman 상관계수:     %.4f\n', corr_original);
fprintf(fid, '가중치점수 Spearman 상관계수: %.4f\n', corr_weighted);
fprintf(fid, '상관계수 개선:                 +%.4f (%.1f%%)\n', ...
    corr_weighted - corr_original, (corr_weighted - corr_original) / abs(corr_original) * 100);

fprintf(fid, '\n================================================================\n');
fprintf(fid, '보고서 종료\n');
fprintf(fid, '================================================================\n');

fclose(fid);
fprintf('보고서 저장: score_comparison_report.txt\n');

%% ========================================================================
%                          종료
% =========================================================================

fprintf('\n');
fprintf('════════════════════════════════════════════════════════════\n');
fprintf('         분석 완료! 모든 결과가 저장되었습니다.\n');
fprintf('════════════════════════════════════════════════════════════\n');
fprintf('출력 디렉토리: %s\n', output_dir);
fprintf('\n생성된 파일:\n');
fprintf('  1. score_comparison_dashboard.png   (종합 대시보드)\n');
fprintf('  2. score_comparison_results.mat     (MATLAB 결과 파일)\n');
fprintf('  3. score_comparison_report.txt      (텍스트 보고서)\n');
fprintf('════════════════════════════════════════════════════════════\n');
