%% 역량검사 가중치 점수 검증: 역량진단(수평평가) 기준
%
% 목적:
%   역량검사의 "원점수"와 "가중치 점수" 중 어느 것이
%   역량진단(수평평가) 점수를 더 잘 예측하는지 비교 검증
%
% Ground Truth: 역량진단 점수 (동료 평가 기반, 독립적 데이터)
%
% 핵심 질문:
%   "가중치 점수가 원점수보다 실제 동료 평가를 더 잘 반영하는가?"
%
% 작성일: 2025-10-14

clear; clc; close all;
rng(42, 'twister'); % 재현성

cd('D:\project\HR데이터\matlab\자가불소_revised');

fprintf('========================================================\n');
fprintf('  역량검사 가중치 점수 검증 (독립적 Ground Truth 사용)\n');
fprintf('========================================================\n\n');

%% 1. 역량진단(수평평가) 데이터 로드
fprintf('[1단계] 역량진단(수평평가) 데이터 로드\n');
fprintf('--------------------------------------------------------\n');

% MAT 파일에서 로드
matPath = 'D:\project\HR데이터\matlab\competency_correlation_workspace_20250915.mat';
if ~exist(matPath, 'file')
    error('역량진단 MAT 파일을 찾을 수 없습니다: %s', matPath);
end

loadedData = load(matPath, 'consolidatedScores');
diagnosticData = loadedData.consolidatedScores;

fprintf('✓ 역량진단 데이터 로드 완료: %d명\n', height(diagnosticData));
fprintf('  - 사용 컬럼: AverageStdScore (표준화 평균 점수)\n');
fprintf('  - 유효 데이터: %d명\n', sum(~isnan(diagnosticData.AverageStdScore)));

%% 2. 역량검사 가중치 점수 데이터 로드
fprintf('\n[2단계] 역량검사 가중치 점수 데이터 로드\n');
fprintf('--------------------------------------------------------\n');

competencyPath = 'D:\project\HR데이터\결과\최종\2025.10.14\역량검사_가중치적용점수_talent_2025-10-14_185545.xlsx';
if ~exist(competencyPath, 'file')
    error('역량검사 가중치 파일을 찾을 수 없습니다: %s', competencyPath);
end

competencyData = readtable(competencyPath, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');

fprintf('✓ 역량검사 데이터 로드 완료: %d명\n', height(competencyData));
fprintf('  - 원점수 유효: %d명\n', sum(~isnan(competencyData.('원점수'))));
fprintf('  - 가중치 점수 유효: %d명\n', sum(~isnan(competencyData.('가중치 점수'))));

%% 3. 데이터 매칭
fprintf('\n[3단계] 데이터 매칭 (ID 기준)\n');
fprintf('--------------------------------------------------------\n');

% ID 표준화
diagnosticIDs = standardizeID(diagnosticData.ID);
competencyIDs = standardizeID(competencyData.ID);

% 교집합 찾기
[commonIDs, diagIdx, compIdx] = intersect(diagnosticIDs, competencyIDs);

fprintf('매칭 결과:\n');
fprintf('  - 역량진단 데이터: %d명\n', height(diagnosticData));
fprintf('  - 역량검사 데이터: %d명\n', height(competencyData));
fprintf('  - 매칭 성공: %d명 (%.1f%%)\n', length(commonIDs), ...
    100 * length(commonIDs) / min(height(diagnosticData), height(competencyData)));

if length(commonIDs) < 10
    error('매칭된 데이터가 너무 적습니다 (최소 10명 필요). ID 형식을 확인하세요.');
end

% 통합 데이터셋 생성
analysisData = table();
analysisData.ID = commonIDs;
analysisData.Diagnostic_Score = diagnosticData.AverageStdScore(diagIdx);
analysisData.Original_Score = competencyData.('원점수')(compIdx);
analysisData.Weighted_Score = competencyData.('가중치 점수')(compIdx);

% 유효 데이터만 선택
validIdx = ~isnan(analysisData.Diagnostic_Score) & ...
           ~isnan(analysisData.Original_Score) & ...
           ~isnan(analysisData.Weighted_Score);

analysisData = analysisData(validIdx, :);

fprintf('\n✓ 최종 분석 데이터: %d명\n', height(analysisData));
fprintf('  - 역량진단 평균: %.3f (±%.3f)\n', ...
    mean(analysisData.Diagnostic_Score), std(analysisData.Diagnostic_Score));
fprintf('  - 원점수 평균: %.2f (±%.2f)\n', ...
    mean(analysisData.Original_Score), std(analysisData.Original_Score));
fprintf('  - 가중치 점수 평균: %.2f (±%.2f)\n', ...
    mean(analysisData.Weighted_Score), std(analysisData.Weighted_Score));

%% 4. 상관분석
fprintf('\n[4단계] 상관분석\n');
fprintf('--------------------------------------------------------\n');

% 원점수 vs 역량진단
[r_orig, p_orig] = corrWithSignificance(analysisData.Original_Score, ...
                                        analysisData.Diagnostic_Score);

% 가중치 점수 vs 역량진단
[r_weighted, p_weighted] = corrWithSignificance(analysisData.Weighted_Score, ...
                                                analysisData.Diagnostic_Score);

fprintf('역량진단 점수와의 상관계수:\n');
fprintf('  원점수:      r = %.4f, p = %.4f %s\n', r_orig, p_orig, ...
    getSignificanceStars(p_orig));
fprintf('  가중치 점수: r = %.4f, p = %.4f %s\n', r_weighted, p_weighted, ...
    getSignificanceStars(p_weighted));

% 상관계수 차이 검정 (Fisher's Z transformation)
n = height(analysisData);
z_orig = 0.5 * log((1 + r_orig) / (1 - r_orig));
z_weighted = 0.5 * log((1 + r_weighted) / (1 - r_weighted));
z_diff = (z_weighted - z_orig) / sqrt(2 / (n - 3));
p_diff = 2 * (1 - normcdf(abs(z_diff)));

fprintf('\n상관계수 차이 검정:\n');
fprintf('  Δr = %.4f\n', r_weighted - r_orig);
fprintf('  z = %.3f, p = %.4f %s\n', z_diff, p_diff, getSignificanceStars(p_diff));

if p_diff < 0.05
    if r_weighted > r_orig
        fprintf('  ✓ 가중치 점수가 통계적으로 유의하게 더 높은 상관을 보입니다.\n');
    else
        fprintf('  ⚠ 원점수가 통계적으로 유의하게 더 높은 상관을 보입니다.\n');
    end
else
    fprintf('  → 두 점수 간 상관계수 차이가 유의하지 않습니다.\n');
end

%% 5. 회귀분석 (예측 정확도 비교)
fprintf('\n[5단계] 회귀분석 (예측 정확도)\n');
fprintf('--------------------------------------------------------\n');

% 원점수 기반 회귀
mdl_orig = fitlm(analysisData.Original_Score, analysisData.Diagnostic_Score);
pred_orig = predict(mdl_orig, analysisData.Original_Score);
residuals_orig = analysisData.Diagnostic_Score - pred_orig;

% 가중치 점수 기반 회귀
mdl_weighted = fitlm(analysisData.Weighted_Score, analysisData.Diagnostic_Score);
pred_weighted = predict(mdl_weighted, analysisData.Weighted_Score);
residuals_weighted = analysisData.Diagnostic_Score - pred_weighted;

% 예측 정확도 지표
mae_orig = mean(abs(residuals_orig));
mae_weighted = mean(abs(residuals_weighted));
rmse_orig = sqrt(mean(residuals_orig.^2));
rmse_weighted = sqrt(mean(residuals_weighted.^2));
r2_orig = mdl_orig.Rsquared.Ordinary;
r2_weighted = mdl_weighted.Rsquared.Ordinary;

fprintf('예측 정확도 비교:\n\n');
fprintf('              원점수      가중치 점수    개선도\n');
fprintf('  --------------------------------------------------------\n');
fprintf('  MAE      %.4f      %.4f      %.4f (%.1f%%)\n', ...
    mae_orig, mae_weighted, mae_orig - mae_weighted, ...
    100 * (mae_orig - mae_weighted) / mae_orig);
fprintf('  RMSE     %.4f      %.4f      %.4f (%.1f%%)\n', ...
    rmse_orig, rmse_weighted, rmse_orig - rmse_weighted, ...
    100 * (rmse_orig - rmse_weighted) / rmse_orig);
fprintf('  R²       %.4f      %.4f      %.4f (%.1f%%)\n', ...
    r2_orig, r2_weighted, r2_weighted - r2_orig, ...
    100 * (r2_weighted - r2_orig) / r2_orig);

%% 6. Top/Bottom 그룹 분류 정확도
fprintf('\n[6단계] 상위/하위 그룹 분류 정확도\n');
fprintf('--------------------------------------------------------\n');

% 역량진단 기준 상위/하위 30% 정의
percentile_30 = prctile(analysisData.Diagnostic_Score, 30);
percentile_70 = prctile(analysisData.Diagnostic_Score, 70);

true_top = analysisData.Diagnostic_Score >= percentile_70;
true_bottom = analysisData.Diagnostic_Score <= percentile_30;

fprintf('그룹 정의 (역량진단 기준):\n');
fprintf('  상위 30%%: %.3f 이상 (%d명)\n', percentile_70, sum(true_top));
fprintf('  하위 30%%: %.3f 이하 (%d명)\n', percentile_30, sum(true_bottom));

% 원점수 기준 예측
pred_top_orig = analysisData.Original_Score >= prctile(analysisData.Original_Score, 70);
pred_bottom_orig = analysisData.Original_Score <= prctile(analysisData.Original_Score, 30);

acc_top_orig = sum(true_top == pred_top_orig) / height(analysisData) * 100;
acc_bottom_orig = sum(true_bottom == pred_bottom_orig) / height(analysisData) * 100;

% 가중치 점수 기준 예측
pred_top_weighted = analysisData.Weighted_Score >= prctile(analysisData.Weighted_Score, 70);
pred_bottom_weighted = analysisData.Weighted_Score <= prctile(analysisData.Weighted_Score, 30);

acc_top_weighted = sum(true_top == pred_top_weighted) / height(analysisData) * 100;
acc_bottom_weighted = sum(true_bottom == pred_bottom_weighted) / height(analysisData) * 100;

fprintf('\n분류 정확도:\n\n');
fprintf('              원점수      가중치 점수    개선도\n');
fprintf('  --------------------------------------------------------\n');
fprintf('  상위 30%%   %.1f%%       %.1f%%         %.1f%%p\n', ...
    acc_top_orig, acc_top_weighted, acc_top_weighted - acc_top_orig);
fprintf('  하위 30%%   %.1f%%       %.1f%%         %.1f%%p\n', ...
    acc_bottom_orig, acc_bottom_weighted, acc_bottom_weighted - acc_bottom_orig);

%% 7. 시각화
fprintf('\n[7단계] 시각화\n');
fprintf('--------------------------------------------------------\n');

fig = figure('Position', [100, 100, 1400, 1000]);

% 1) 산점도 + 회귀선 (원점수)
subplot(2, 3, 1);
scatter(analysisData.Original_Score, analysisData.Diagnostic_Score, 50, 'b', 'filled', 'MarkerFaceAlpha', 0.5);
hold on;
plot(analysisData.Original_Score, pred_orig, 'r-', 'LineWidth', 2);
xlabel('역량검사 원점수');
ylabel('역량진단 점수 (수평평가)');
title(sprintf('원점수 vs 역량진단\n(r=%.3f, R²=%.3f)', r_orig, r2_orig));
grid on;
legend({'실제 데이터', '회귀선'}, 'Location', 'best');

% 2) 산점도 + 회귀선 (가중치 점수)
subplot(2, 3, 2);
scatter(analysisData.Weighted_Score, analysisData.Diagnostic_Score, 50, 'g', 'filled', 'MarkerFaceAlpha', 0.5);
hold on;
plot(analysisData.Weighted_Score, pred_weighted, 'r-', 'LineWidth', 2);
xlabel('역량검사 가중치 점수');
ylabel('역량진단 점수 (수평평가)');
title(sprintf('가중치 점수 vs 역량진단\n(r=%.3f, R²=%.3f)', r_weighted, r2_weighted));
grid on;
legend({'실제 데이터', '회귀선'}, 'Location', 'best');

% 3) 잔차 플롯 비교
subplot(2, 3, 3);
boxplot([residuals_orig, residuals_weighted], 'Labels', {'원점수', '가중치 점수'});
ylabel('잔차 (실제 - 예측)');
title('예측 오차 분포');
grid on;
yline(0, 'r--');

% 4) 상관계수 비교
subplot(2, 3, 4);
bar([r_orig, r_weighted]);
set(gca, 'XTickLabel', {'원점수', '가중치 점수'});
ylabel('상관계수 (r)');
title('역량진단 점수와의 상관계수');
ylim([0, max([r_orig, r_weighted]) * 1.2]);
grid on;
for i = 1:2
    text(i, [r_orig, r_weighted](i) + 0.02, sprintf('%.3f', [r_orig, r_weighted](i)), ...
        'HorizontalAlignment', 'center', 'FontWeight', 'bold');
end

% 5) 예측 정확도 비교 (MAE, RMSE)
subplot(2, 3, 5);
metrics = [mae_orig, mae_weighted; rmse_orig, rmse_weighted];
b = bar(metrics);
set(gca, 'XTickLabel', {'MAE', 'RMSE'});
ylabel('오차');
title('예측 오차 비교');
legend({'원점수', '가중치 점수'}, 'Location', 'best');
grid on;

% 6) 분류 정확도 비교
subplot(2, 3, 6);
acc_data = [acc_top_orig, acc_top_weighted; acc_bottom_orig, acc_bottom_weighted];
b = bar(acc_data);
set(gca, 'XTickLabel', {'상위 30%', '하위 30%'});
ylabel('정확도 (%)');
title('상위/하위 그룹 분류 정확도');
legend({'원점수', '가중치 점수'}, 'Location', 'best');
grid on;
ylim([0, 100]);

% 그림 저장
saveas(fig, 'validation_results_peer_evaluation.png');
fprintf('✓ 시각화 저장 완료: validation_results_peer_evaluation.png\n');

%% 8. 종합 리포트
fprintf('\n========================================================\n');
fprintf('              종합 결과 및 해석\n');
fprintf('========================================================\n\n');

fprintf('【 핵심 결과 】\n\n');

% 상관계수 비교 해석
if r_weighted > r_orig && p_diff < 0.05
    fprintf('✓ 가중치 점수가 원점수보다 역량진단(수평평가) 점수를\n');
    fprintf('  통계적으로 유의하게 더 잘 예측합니다.\n');
    fprintf('  (상관계수: %.3f → %.3f, p=%.4f)\n\n', r_orig, r_weighted, p_diff);
elseif r_weighted > r_orig
    fprintf('→ 가중치 점수가 원점수보다 높은 상관을 보이나,\n');
    fprintf('  통계적으로 유의한 차이는 아닙니다.\n');
    fprintf('  (상관계수: %.3f → %.3f, p=%.4f)\n\n', r_orig, r_weighted, p_diff);
else
    fprintf('⚠ 원점수가 가중치 점수와 비슷하거나 더 높은 상관을 보입니다.\n');
    fprintf('  가중치 모델의 개선 효과가 제한적입니다.\n');
    fprintf('  (상관계수: %.3f vs %.3f, p=%.4f)\n\n', r_orig, r_weighted, p_diff);
end

% 예측 정확도 해석
mae_improvement = (mae_orig - mae_weighted) / mae_orig * 100;
if mae_improvement > 5
    fprintf('✓ 가중치 점수의 예측 오차(MAE)가 %.1f%% 감소했습니다.\n', mae_improvement);
    fprintf('  (%.4f → %.4f)\n\n', mae_orig, mae_weighted);
elseif mae_improvement > 0
    fprintf('→ 가중치 점수의 예측 오차가 소폭 감소했습니다 (%.1f%%).\n\n', mae_improvement);
else
    fprintf('⚠ 가중치 점수의 예측 정확도 개선이 미미합니다.\n\n');
end

% 분류 정확도 해석
top_improvement = acc_top_weighted - acc_top_orig;
bottom_improvement = acc_bottom_weighted - acc_bottom_orig;

if top_improvement > 5
    fprintf('✓ 상위 30%% 그룹 분류 정확도가 %.1f%%p 향상되었습니다.\n', top_improvement);
elseif top_improvement > 0
    fprintf('→ 상위 30%% 그룹 분류가 소폭 개선되었습니다 (%.1f%%p).\n', top_improvement);
else
    fprintf('⚠ 상위 그룹 분류 정확도가 개선되지 않았습니다.\n');
end

if bottom_improvement > 5
    fprintf('✓ 하위 30%% 그룹 분류 정확도가 %.1f%%p 향상되었습니다.\n\n', bottom_improvement);
elseif bottom_improvement > 0
    fprintf('→ 하위 30%% 그룹 분류가 소폭 개선되었습니다 (%.1f%%p).\n\n', bottom_improvement);
else
    fprintf('⚠ 하위 그룹 분류 정확도가 개선되지 않았습니다.\n\n');
end

fprintf('--------------------------------------------------------\n\n');

fprintf('【 실무적 의미 】\n\n');

if r_weighted > r_orig && mae_improvement > 5
    fprintf('이 결과는 가중치 모델이 자가 평가의 편향을 보정하여\n');
    fprintf('실제 동료 평가(역량진단)에 더 가까운 점수를 제공함을 의미합니다.\n\n');
    fprintf('→ 채용 및 평가에서 가중치 점수 사용을 권장합니다.\n\n');
elseif r_weighted > r_orig
    fprintf('가중치 모델이 개선 효과를 보이나, 그 정도가 제한적입니다.\n');
    fprintf('추가 검증 및 모델 개선을 고려할 필요가 있습니다.\n\n');
else
    fprintf('가중치 모델의 효과가 명확하지 않습니다.\n');
    fprintf('모델 재설계 또는 다른 접근법을 검토할 필요가 있습니다.\n\n');
end

fprintf('--------------------------------------------------------\n\n');

fprintf('【 기술적 세부사항 】\n\n');
fprintf('분석 대상: %d명\n', height(analysisData));
fprintf('상관계수: %.4f → %.4f (Δr = %.4f, p = %.4f)\n', ...
    r_orig, r_weighted, r_weighted - r_orig, p_diff);
fprintf('예측 오차: MAE %.4f → %.4f (%.1f%% 개선)\n', ...
    mae_orig, mae_weighted, mae_improvement);
fprintf('결정계수: R² %.4f → %.4f\n\n', r2_orig, r2_weighted);

% 결과 테이블 저장
results = table();
results.Metric = {'Correlation'; 'MAE'; 'RMSE'; 'R_squared'; 'Top30_Accuracy'; 'Bottom30_Accuracy'};
results.Original_Score = [r_orig; mae_orig; rmse_orig; r2_orig; acc_top_orig; acc_bottom_orig];
results.Weighted_Score = [r_weighted; mae_weighted; rmse_weighted; r2_weighted; acc_top_weighted; acc_bottom_weighted];
results.Improvement = results.Weighted_Score - results.Original_Score;
results.Improvement_Percent = (results.Improvement ./ results.Original_Score) * 100;

writetable(results, 'validation_results_summary.xlsx');
writetable(analysisData, 'validation_analysis_data.xlsx');

fprintf('✓ 결과 저장 완료:\n');
fprintf('  - validation_results_summary.xlsx (요약 통계)\n');
fprintf('  - validation_analysis_data.xlsx (분석 데이터)\n');
fprintf('  - validation_results_peer_evaluation.png (시각화)\n\n');

fprintf('========================================================\n');
fprintf('분석 완료!\n');
fprintf('========================================================\n');

%% ===== 보조 함수 =====

function standardIDs = standardizeID(rawIDs)
    % ID를 문자열로 표준화
    if isnumeric(rawIDs)
        standardIDs = arrayfun(@(x) sprintf('%.0f', x), rawIDs, 'UniformOutput', false);
    elseif iscell(rawIDs)
        standardIDs = cellfun(@char, rawIDs, 'UniformOutput', false);
    elseif isstring(rawIDs)
        standardIDs = cellstr(rawIDs);
    else
        standardIDs = cellstr(rawIDs);
    end

    % 공백 제거
    standardIDs = strtrim(standardIDs);
end

function [r, p] = corrWithSignificance(x, y)
    % 상관계수와 유의확률 계산
    [r_matrix, p_matrix] = corrcoef(x, y);
    r = r_matrix(1, 2);
    p = p_matrix(1, 2);
end

function stars = getSignificanceStars(p)
    % 유의수준 별표 표시
    if p < 0.001
        stars = '***';
    elseif p < 0.01
        stars = '**';
    elseif p < 0.05
        stars = '*';
    else
        stars = '';
    end
end
