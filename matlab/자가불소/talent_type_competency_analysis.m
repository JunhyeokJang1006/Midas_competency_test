%% Talent Type Competency Analysis and Weight Calculation
% Complete analysis: Data loading, matching, statistical analysis, and weight calculation

clear; clc;

%% 1. Data Loading
fprintf('=== Starting Data Loading ===\n');

% 인적정보 데이터 (인재유형 포함)
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');

% 역량검사 데이터
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
[comp_num, comp_txt, comp_raw] = xlsread(comp_file, 1);

fprintf('인적정보 데이터: %d행 x %d열\n', height(hr_data), width(hr_data));
fprintf('역량검사 데이터: %d행 x %d열\n', size(comp_raw, 1), size(comp_raw, 2));

%% 2. 인재유형 데이터 추출
fprintf('\n=== 인재유형 데이터 추출 ===\n');

% 인재유형이 있는 행만 필터링
hr_clean = hr_data(~cellfun(@isempty, hr_data.('인재유형')), :);
fprintf('인재유형 데이터가 있는 직원 수: %d명\n', height(hr_clean));

% 인재유형별 분포 확인
talent_types = hr_clean.('인재유형');
unique_types = unique(talent_types(~cellfun(@isempty, talent_types)));

fprintf('\n인재유형별 분포:\n');
for i = 1:length(unique_types)
    count = sum(strcmp(talent_types, unique_types{i}));
    fprintf('%s: %d명\n', unique_types{i}, count);
end

%% 3. 역량검사 데이터 구조 파악
fprintf('\n=== 역량검사 데이터 구조 파악 ===\n');

% 헤더 행 (3행)과 데이터 시작점 확인
headers = comp_raw(3, :);
data_start_row = 4;

% 역량 항목 열 찾기 (10열부터 주요 역량 점수들이 있음)
competency_headers = {};
competency_cols = [];

for j = 10:min(50, size(comp_raw, 2))
    header_val = comp_raw{3, j};
    if ischar(header_val) && ~isempty(header_val)
        competency_headers{end+1} = header_val;
        competency_cols = [competency_cols, j];
    end
end

fprintf('발견된 역량 항목 수: %d개\n', length(competency_headers));
fprintf('주요 역량 항목들:\n');
for i = 1:min(10, length(competency_headers))
    fprintf('  %d. %s\n', i, competency_headers{i});
end

%% 4. 역량검사 데이터에서 유효한 데이터 추출
fprintf('\n=== 역량검사 유효 데이터 추출 ===\n');

% 유효한 ID와 점수가 있는 행들 찾기
valid_comp_data = [];
valid_ids = [];

for i = data_start_row:size(comp_raw, 1)
    id_val = comp_raw{i, 1};
    if isnumeric(id_val) && id_val > 1000000
        % 이 행의 역량 점수들 추출
        scores = [];
        for j = 1:length(competency_cols)
            col_idx = competency_cols(j);
            val = comp_raw{i, col_idx};
            if isnumeric(val) && ~isnan(val) && val >= 1 && val <= 100
                scores = [scores, val];
            else
                scores = [scores, NaN];
            end
        end

        % 최소 5개 이상의 유효한 점수가 있으면 포함
        if sum(~isnan(scores)) >= 5
            valid_ids = [valid_ids; id_val];
            valid_comp_data = [valid_comp_data; scores];
        end
    end
end

fprintf('유효한 역량검사 데이터: %d명\n', length(valid_ids));

%% 5. 인재유형과 역량검사 데이터 매칭
fprintf('\n=== 데이터 매칭 ===\n');

matched_data = [];
matched_talent_types = {};
matched_ids = [];

for i = 1:height(hr_clean)
    hr_id = hr_clean.ID(i);

    % 역량검사 데이터에서 해당 ID 찾기
    comp_idx = find(valid_ids == hr_id);

    if ~isempty(comp_idx) && ~isempty(hr_clean.('인재유형'){i})
        matched_ids = [matched_ids; hr_id];
        matched_talent_types{end+1} = hr_clean.('인재유형'){i};
        matched_data = [matched_data; valid_comp_data(comp_idx, :)];
    end
end

fprintf('매칭된 데이터: %d명\n', length(matched_ids));

% 매칭된 인재유형별 분포
fprintf('\n매칭된 인재유형별 분포:\n');
unique_matched_types = unique(matched_talent_types);
for i = 1:length(unique_matched_types)
    count = sum(strcmp(matched_talent_types, unique_matched_types{i}));
    fprintf('%s: %d명\n', unique_matched_types{i}, count);
end

%% 6. 인재유형별 성과 순위 정의
fprintf('\n=== 인재유형 성과 순위 정의 ===\n');

% 사용자가 제공한 성과 순위 (높은 순서대로)
performance_ranking = {
    '자연성', 8;
    '성실한 가연성', 7;
    '유익한 불연성', 6;
    '유능한 불연성', 5;
    '게으른 가연성', 4;
    '무능한 불연성', 3;
    '위장형 소화성', 2;
    '소화성', 1
};

% 성과 점수 매핑
performance_scores = containers.Map();
for i = 1:size(performance_ranking, 1)
    performance_scores(performance_ranking{i, 1}) = performance_ranking{i, 2};
end

% 매칭된 데이터에 성과 점수 부여
matched_performance = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    talent_type = matched_talent_types{i};
    if performance_scores.isKey(talent_type)
        matched_performance(i) = performance_scores(talent_type);
    else
        fprintf('경고: 알 수 없는 인재유형 - %s\n', talent_type);
        matched_performance(i) = 0;
    end
end

fprintf('성과 점수가 부여된 데이터: %d명\n', sum(matched_performance > 0));

%% 7. 데이터 저장 (첫 번째 단계)
fprintf('\n=== 데이터 저장 ===\n');

% 결과를 구조체에 저장
analysis_data = struct();
analysis_data.matched_ids = matched_ids;
analysis_data.matched_talent_types = matched_talent_types;
analysis_data.matched_data = matched_data;
analysis_data.matched_performance = matched_performance;
analysis_data.competency_headers = competency_headers;
analysis_data.performance_ranking = performance_ranking;

% 파일로 저장
save('talent_competency_merged_data.mat', 'analysis_data');
fprintf('병합된 데이터가 저장되었습니다: talent_competency_merged_data.mat\n');

%% 8. 통계 분석 시작
fprintf('\n\n=== 통계 분석 시작 ===\n');

% 데이터 품질 분석
fprintf('\n=== 데이터 품질 분석 ===\n');

missing_ratio = sum(isnan(matched_data), 1) ./ size(matched_data, 1);
fprintf('결측값 비율 (상위 10개 역량):\n');
for i = 1:min(10, length(competency_headers))
    fprintf('  %s: %.2f%%\n', competency_headers{i}, missing_ratio(i)*100);
end

% 50% 이상 결측값이 있는 역량 제거
valid_comp_idx = missing_ratio < 0.5;
valid_competencies = competency_headers(valid_comp_idx);
valid_data = matched_data(:, valid_comp_idx);

fprintf('\n유효한 역량 (결측률 < 50%%): %d개\n', sum(valid_comp_idx));

%% 9. 인재유형별 성과 분석
fprintf('\n=== 인재유형별 성과 분석 ===\n');

performance_by_type = [];
competency_means_by_type = [];
competency_stds_by_type = [];

fprintf('성과별 인재유형 분포:\n');
for i = 1:length(unique_matched_types)
    talent_type = unique_matched_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);

    % 성과 통계
    type_performance = matched_performance(type_idx);
    mean_performance = mean(type_performance);
    count = sum(type_idx);

    fprintf('  %s: %.2f (n=%d)\n', talent_type, mean_performance, count);
    performance_by_type = [performance_by_type; mean_performance];

    % 이 인재유형의 역량 통계
    type_data = valid_data(type_idx, :);
    type_means = nanmean(type_data, 1);
    type_stds = nanstd(type_data, 1);

    competency_means_by_type = [competency_means_by_type; type_means];
    competency_stds_by_type = [competency_stds_by_type; type_stds];
end

%% 10. 성과 수준별 역량 분석
fprintf('\n=== 성과 수준별 역량 분석 ===\n');

% 고성과 vs 저성과 그룹으로 구분
high_perf_threshold = median(matched_performance);
high_perf_idx = matched_performance > high_perf_threshold;
low_perf_idx = matched_performance <= high_perf_threshold;

fprintf('고성과 그룹: %d명 (성과 > %.1f)\n', sum(high_perf_idx), high_perf_threshold);
fprintf('저성과 그룹: %d명 (성과 <= %.1f)\n', sum(low_perf_idx), high_perf_threshold);

% 고성과 vs 저성과 그룹 간 역량 점수 비교
high_perf_data = valid_data(high_perf_idx, :);
low_perf_data = valid_data(low_perf_idx, :);

high_perf_means = nanmean(high_perf_data, 1);
low_perf_means = nanmean(low_perf_data, 1);

competency_differences = high_perf_means - low_perf_means;

%% 11. 역량별 상관관계 및 효과 크기 계산
fprintf('\n=== 역량별 상관관계 및 효과 크기 계산 ===\n');

correlations = [];
effect_sizes = [];
p_values = [];

for i = 1:size(valid_data, 2)
    comp_scores = valid_data(:, i);
    valid_pairs_idx = ~isnan(comp_scores) & ~isnan(matched_performance);

    if sum(valid_pairs_idx) >= 10  % 최소 10개 유효 쌍
        % 상관계수 계산 (수동)
        x_valid = comp_scores(valid_pairs_idx);
        y_valid = matched_performance(valid_pairs_idx);

        if var(x_valid) > 0 && var(y_valid) > 0
            corr_coef = sum((x_valid - mean(x_valid)) .* (y_valid - mean(y_valid))) / ...
                       (sqrt(sum((x_valid - mean(x_valid)).^2)) * sqrt(sum((y_valid - mean(y_valid)).^2)));
        else
            corr_coef = 0;
        end
        correlations = [correlations, corr_coef];

        % 효과 크기 계산
        effect_size = abs(competency_differences(i)) / (nanstd(comp_scores) + eps);
        effect_sizes = [effect_sizes, effect_size];

        % t-test (간단한 근사)
        high_scores = comp_scores(valid_pairs_idx & high_perf_idx);
        low_scores = comp_scores(valid_pairs_idx & low_perf_idx);

        if length(high_scores) >= 3 && length(low_scores) >= 3
            pooled_std = sqrt(((length(high_scores)-1)*var(high_scores) + (length(low_scores)-1)*var(low_scores)) / ...
                             (length(high_scores) + length(low_scores) - 2));
            t_stat = (mean(high_scores) - mean(low_scores)) / ...
                    (pooled_std * sqrt(1/length(high_scores) + 1/length(low_scores)));
            % 자유도 근사 p-value (간단한 근사)
            df = length(high_scores) + length(low_scores) - 2;
            if abs(t_stat) > 2.0  % 대략적인 p < 0.05 기준
                p_val = 0.04;
            else
                p_val = 0.2;
            end
        else
            p_val = 1.0;
        end
        p_values = [p_values, p_val];
    else
        correlations = [correlations, NaN];
        effect_sizes = [effect_sizes, NaN];
        p_values = [p_values, NaN];
    end
end

%% 12. 종합 가중치 계산
fprintf('\n=== 종합 가중치 계산 ===\n');

% 지표 정규화
norm_correlation = abs(correlations) / (max(abs(correlations)) + eps);
norm_effect_size = effect_sizes / (max(effect_sizes) + eps);
significance_weight = 1 - p_values;  % p값이 낮을수록 높은 가중치

% 결측값 처리
norm_correlation(isnan(norm_correlation)) = 0;
norm_effect_size(isnan(norm_effect_size)) = 0;
significance_weight(isnan(significance_weight)) = 0;

% 종합 가중치 (가중평균)
weight_corr = 0.4;      % 40% 상관관계
weight_effect = 0.3;    % 30% 효과크기
weight_sig = 0.3;       % 30% 통계적 유의성

composite_weights = weight_corr * norm_correlation + ...
                   weight_effect * norm_effect_size + ...
                   weight_sig * significance_weight;

%% 13. 결과 정렬 및 표시
fprintf('\n=== 성과 예측을 위한 TOP 15 역량 ===\n');

% 종합 가중치로 정렬
[sorted_weights, sort_idx] = sort(composite_weights, 'descend');

fprintf('순위 | 역량명                    | 종합가중치 | 상관계수 | 차이값 | 효과크기 | p-value\n');
fprintf('-----|---------------------------|------------|----------|--------|----------|---------\n');

top_competencies = {};
top_weights = [];

for i = 1:min(15, length(sort_idx))
    idx = sort_idx(i);
    if ~isnan(sorted_weights(i)) && sorted_weights(i) > 0
        fprintf('%4d | %-25s | %10.3f | %8.3f | %6.2f | %8.3f | %7.3f\n', ...
                i, valid_competencies{idx}, ...
                composite_weights(idx), ...
                correlations(idx), ...
                competency_differences(idx), ...
                effect_sizes(idx), ...
                p_values(idx));

        top_competencies{end+1} = valid_competencies{idx};
        top_weights = [top_weights; composite_weights(idx)];
    end
end

%% 14. TOP 5 핵심 역량 가중치
fprintf('\n=== TOP 5 핵심 역량 가중치 ===\n');

top_5_count = min(5, length(top_competencies));
top_5_total = sum(top_weights(1:top_5_count));

fprintf('성과 예측을 위한 핵심 역량 가중치:\n');
for i = 1:top_5_count
    weight_percent = (top_weights(i) / top_5_total) * 100;
    fprintf('  %d. %s: %.1f%%\n', i, top_competencies{i}, weight_percent);
end

%% 15. 최종 결과 저장
fprintf('\n=== 최종 결과 저장 ===\n');

% 통계 분석 결과
statistical_results = struct();
statistical_results.unique_talent_types = unique_matched_types;
statistical_results.performance_by_type = performance_by_type;
statistical_results.valid_competencies = valid_competencies;
statistical_results.correlations = correlations;
statistical_results.effect_sizes = effect_sizes;
statistical_results.p_values = p_values;
statistical_results.composite_weights = composite_weights;
statistical_results.sort_idx = sort_idx;

% 최종 추천사항
recommendations = struct();
recommendations.top_competencies = top_competencies(1:top_5_count);
recommendations.top_weights_normalized = top_weights(1:top_5_count) / top_5_total;
recommendations.analysis_summary = sprintf('총 %d명의 데이터로 %d개 역량 분석', ...
    length(matched_ids), length(valid_competencies));

% 파일 저장
save('statistical_analysis_results.mat', 'statistical_results');
save('final_recommendations.mat', 'recommendations');

% 엑셀 결과 저장
results_table = table();
for i = 1:length(sort_idx)
    idx = sort_idx(i);
    results_table.Rank(i) = i;
    results_table.Competency{i} = valid_competencies{idx};
    results_table.Composite_Weight(i) = composite_weights(idx);
    results_table.Correlation(i) = correlations(idx);
    results_table.Effect_Size(i) = effect_sizes(idx);
    results_table.P_Value(i) = p_values(idx);
end

try
    writetable(results_table, 'competency_analysis_complete.xlsx');
    fprintf('엑셀 파일 저장: competency_analysis_complete.xlsx\n');
catch
    fprintf('엑셀 파일 저장 실패 (Excel이 설치되지 않았을 수 있음)\n');
end

fprintf('\n저장된 파일들:\n');
fprintf('  - talent_competency_merged_data.mat (병합 데이터)\n');
fprintf('  - statistical_analysis_results.mat (통계 분석 결과)\n');
fprintf('  - final_recommendations.mat (최종 추천사항)\n');
fprintf('  - competency_analysis_complete.xlsx (전체 결과표)\n');

%% 16. 최종 요약
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 60));
fprintf('                   분석 완료 요약\n');
fprintf('%s\n', repmat('=', 1, 60));
fprintf('• 분석 대상자: %d명\n', length(matched_ids));
fprintf('• 분석 역량: %d개\n', length(valid_competencies));
fprintf('• 핵심 예측 역량: %d개\n', top_5_count);
fprintf('\n인재유형별 평균 성과 점수:\n');
for i = 1:length(unique_matched_types)
    type_idx = strcmp(matched_talent_types, unique_matched_types{i});
    fprintf('  %s: %.1f점\n', unique_matched_types{i}, mean(matched_performance(type_idx)));
end
fprintf('\n성과 예측 핵심 역량 (상위 5개):\n');
for i = 1:top_5_count
    weight_percent = (top_weights(i) / top_5_total) * 100;
    fprintf('  %d. %s (%.1f%%)\n', i, top_competencies{i}, weight_percent);
end
fprintf('%s\n', repmat('=', 1, 60));
fprintf('분석이 성공적으로 완료되었습니다!\n');