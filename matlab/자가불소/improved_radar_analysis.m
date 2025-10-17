%% Improved Talent Type Analysis with Performance-Ordered Radar Charts
% 성과순서대로 정렬된 레이더 차트 생성 및 데이터 사용량 최적화

clear; clc; close all;

%% Global Settings
set(0, 'DefaultAxesFontName', 'Malgun Gothic');
set(0, 'DefaultTextFontName', 'Malgun Gothic');
set(0, 'DefaultAxesFontSize', 11);
set(0, 'DefaultTextFontSize', 11);

%% 1. Data Loading and Analysis
fprintf('========================================\n');
fprintf('📊 개선된 인재유형 분석 (성과순서 레이더 차트)\n');
fprintf('========================================\n\n');

% File paths
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

% Load data
fprintf('▶ 데이터 로딩...\n');
hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
comp_upper = readtable(comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
comp_total = readtable(comp_file, 'Sheet', '역량검사_종합점수', 'VariableNamingRule', 'preserve');

fprintf('  ✓ HR 데이터: %d명\n', height(hr_data));
fprintf('  ✓ 역량검사 상위항목: %d명\n', height(comp_upper));
fprintf('  ✓ 역량검사 종합점수: %d명\n', height(comp_total));

%% 2. Data Usage Analysis and Optimization
fprintf('\n=== 데이터 사용량 분석 ===\n');

% 인재유형이 있는 데이터만 필터링
talent_data = hr_data.인재유형;
valid_talent_idx = ~cellfun(@isempty, talent_data) & ~ismissing(talent_data);
hr_clean = hr_data(valid_talent_idx, :);

fprintf('전체 HR 데이터: %d명\n', height(hr_data));
fprintf('인재유형 보유자: %d명 (%.1f%%)\n', height(hr_clean), height(hr_clean)/height(hr_data)*100);

% ID 매칭
hr_clean_ids = hr_clean.ID;
comp_upper_ids = comp_upper.ID;
comp_total_ids = comp_total.ID;

[matched_ids, hr_idx, comp_idx] = intersect(hr_clean_ids, comp_upper_ids);
[total_matched_ids, hr_total_idx, comp_total_idx] = intersect(hr_clean_ids, comp_total_ids);

fprintf('HR ↔ 상위항목 매칭: %d명 (%.1f%%)\n', length(matched_ids), length(matched_ids)/height(hr_clean)*100);
fprintf('HR ↔ 종합점수 매칭: %d명 (%.1f%%)\n', length(total_matched_ids), length(total_matched_ids)/height(hr_clean)*100);

% 최종 분석 데이터
matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, :);
matched_total = comp_total(comp_total_idx, :);

matched_talent_types = matched_hr.인재유형;
total_scores = matched_total{:, end}; % 마지막 컬럼이 종합점수

%% 3. 역량 항목 추출 (개선된 방법)
fprintf('\n=== 역량 항목 추출 ===\n');

% 상위항목 역량 컬럼 찾기 (6번째 컬럼부터)
valid_comp_cols = {};
valid_comp_indices = [];

for i = 6:width(matched_comp)
    col_name = matched_comp.Properties.VariableNames{i};
    col_data = matched_comp{:, i};

    % 숫자 데이터이고 결측값이 50% 미만인 컬럼만 선택
    if isnumeric(col_data) && sum(~isnan(col_data)) >= length(col_data) * 0.5
        valid_comp_cols{end+1} = col_name;
        valid_comp_indices = [valid_comp_indices, i];
    end
end

fprintf('유효한 상위항목 역량: %d개\n', length(valid_comp_cols));
for i = 1:length(valid_comp_cols)
    fprintf('  %d. %s\n', i, valid_comp_cols{i});
end

%% 4. 성과 순위 정의
performance_ranking = containers.Map();
performance_ranking('자연성') = 8;
performance_ranking('성실한 가연성') = 7;
performance_ranking('유익한 불연성') = 6;
performance_ranking('유능한 불연성') = 5;
performance_ranking('게으른 가연성') = 4;
performance_ranking('무능한 불연성') = 3;
performance_ranking('위장형 소화성') = 2;
performance_ranking('소화성') = 1;

%% 5. 인재유형별 통계 및 성과 분석
fprintf('\n=== 인재유형별 분석 ===\n');

unique_types = unique(matched_talent_types);
type_stats = [];

fprintf('인재유형별 상세 분석:\n');
fprintf('인재유형            | 인원 | 역량평균 | 종합점수 | 성과순위 | 데이터품질\n');
fprintf('-------------------|------|----------|----------|----------|----------\n');

for i = 1:length(unique_types)
    talent_type = unique_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);

    % 기본 통계
    count = sum(type_idx);

    % 역량 점수 통계
    type_comp_data = matched_comp{type_idx, valid_comp_indices};
    comp_mean = nanmean(type_comp_data(:));

    % 종합점수 통계
    type_total_scores = total_scores(type_idx);
    total_mean = nanmean(type_total_scores);

    % 성과 순위
    if performance_ranking.isKey(talent_type)
        perf_rank = performance_ranking(talent_type);
    else
        perf_rank = 0;
    end

    % 데이터 품질 (결측값 비율)
    data_quality = sum(~isnan(type_comp_data(:))) / numel(type_comp_data) * 100;

    fprintf('%-18s | %4d | %8.1f | %8.1f | %8d | %7.1f%%\n', ...
            talent_type, count, comp_mean, total_mean, perf_rank, data_quality);

    type_stats = [type_stats; {talent_type, count, comp_mean, total_mean, perf_rank, data_quality}];
end

%% 6. 성과 예측 가중치 계산
fprintf('\n=== 성과 예측 가중치 계산 ===\n');

% 성과 점수 매핑
performance_scores = zeros(length(matched_talent_types), 1);
for i = 1:length(matched_talent_types)
    if performance_ranking.isKey(matched_talent_types{i})
        performance_scores(i) = performance_ranking(matched_talent_types{i});
    end
end

% 각 역량별 상관분석
comp_importance = [];
comp_correlations = [];

fprintf('역량별 성과 예측력:\n');
fprintf('역량항목        | 상관계수 | 고성과평균 | 저성과평균 | 차이값 | 중요도\n');
fprintf('----------------|----------|------------|------------|--------|--------\n');

for j = 1:length(valid_comp_cols)
    comp_name = valid_comp_cols{j};
    comp_scores = matched_comp{:, valid_comp_indices(j)};

    valid_idx = ~isnan(comp_scores) & performance_scores > 0;

    if sum(valid_idx) >= 10
        valid_comp_scores = comp_scores(valid_idx);
        valid_perf_scores = performance_scores(valid_idx);

        % 상관계수 계산
        if var(valid_comp_scores) > 0 && var(valid_perf_scores) > 0
            correlation = corr(valid_comp_scores, valid_perf_scores);
        else
            correlation = 0;
        end

        % 고성과 vs 저성과 비교
        high_perf_threshold = median(valid_perf_scores);
        high_idx = valid_perf_scores > high_perf_threshold;
        low_idx = valid_perf_scores <= high_perf_threshold;

        high_mean = mean(valid_comp_scores(high_idx));
        low_mean = mean(valid_comp_scores(low_idx));
        difference = high_mean - low_mean;

        % 중요도 점수 (상관계수 + 효과크기)
        importance = abs(correlation) * 0.7 + abs(difference/std(valid_comp_scores)) * 0.3;

        fprintf('%-15s | %8.3f | %10.1f | %10.1f | %6.1f | %6.3f\n', ...
                comp_name, correlation, high_mean, low_mean, difference, importance);

        comp_correlations = [comp_correlations; correlation];
        comp_importance = [comp_importance; importance];
    else
        comp_correlations = [comp_correlations; 0];
        comp_importance = [comp_importance; 0];
    end
end

% 상위 5개 역량 선정
[~, top_idx] = sort(comp_importance, 'descend');
top_5_idx = top_idx(1:min(5, length(top_idx)));
top_5_competencies = valid_comp_cols(top_5_idx);
top_5_weights = comp_importance(top_5_idx);
top_5_weights_normalized = top_5_weights / sum(top_5_weights);

fprintf('\n상위 5개 핵심 역량:\n');
for i = 1:length(top_5_competencies)
    fprintf('  %d. %s: %.1f%%\n', i, top_5_competencies{i}, top_5_weights_normalized(i)*100);
end

%% 7. 성과순서대로 정렬된 레이더 차트 생성
fprintf('\n=== 성과순서 레이더 차트 생성 ===\n');

% 성과순위로 인재유형 정렬
type_performance = [];
for i = 1:length(unique_types)
    if performance_ranking.isKey(unique_types{i})
        type_performance = [type_performance; performance_ranking(unique_types{i})];
    else
        type_performance = [type_performance; 0];
    end
end

[~, perf_sort_idx] = sort(type_performance, 'descend');
sorted_types = unique_types(perf_sort_idx);
sorted_performance = type_performance(perf_sort_idx);

% 컬러 팔레트 (성과 높은 순서대로)
performance_colors = [
    0.0000 0.4470 0.7410;  % 진한 파랑 (최고성과)
    0.0940 0.6940 0.1250;  % 진한 초록
    0.4940 0.1840 0.5560;  % 진한 보라
    0.8500 0.3250 0.0980;  % 진한 주황
    0.6350 0.0780 0.1840;  % 진한 빨강
    0.4660 0.6740 0.1880;  % 올리브
    0.3010 0.7450 0.9330;  % 하늘색
    0.9290 0.6940 0.1250;  % 노랑 (최저성과)
];

% 전체 평균 계산 (기준선)
overall_means = nanmean(matched_comp{:, valid_comp_indices}, 1);

% 메인 레이더 차트 생성
figure('Position', [50 50 1600 1200], 'Color', 'white', 'Name', '인재유형별 역량 프로필 (성과순서)');

n_types = length(sorted_types);
n_cols = 3;
n_rows = ceil(n_types / n_cols);

for i = 1:n_types
    subplot(n_rows, n_cols, i);

    talent_type = sorted_types{i};
    perf_rank = sorted_performance(i);
    type_idx = strcmp(matched_talent_types, talent_type);

    % 해당 인재유형의 역량 평균 계산
    type_comp_data = matched_comp{type_idx, valid_comp_indices};
    type_means = nanmean(type_comp_data, 1);

    % 상위 5개 역량만 사용
    radar_data = type_means(top_5_idx);
    radar_baseline = overall_means(top_5_idx);

    % 0-100 스케일로 정규화
    radar_data = max(0, min(100, radar_data));
    radar_baseline = max(0, min(100, radar_baseline));

    % 레이더 차트 그리기
    create_performance_radar(radar_data, radar_baseline, top_5_competencies, ...
                           talent_type, perf_rank, performance_colors(i, :), sum(type_idx));
end

sgtitle('인재유형별 핵심 역량 프로필 (성과순서: 높음→낮음)', 'FontSize', 18, 'FontWeight', 'bold');

%% 8. 결과 저장
fprintf('\n=== 결과 저장 ===\n');

% 성과순서 요약 테이블
performance_summary = table();
performance_summary.순위 = (1:length(sorted_types))';
performance_summary.인재유형 = sorted_types;
performance_summary.성과점수 = sorted_performance;

for i = 1:length(sorted_types)
    talent_type = sorted_types{i};
    type_idx = strcmp(matched_talent_types, talent_type);
    performance_summary.인원수(i) = sum(type_idx);

    type_comp_data = matched_comp{type_idx, valid_comp_indices};
    performance_summary.평균역량점수(i) = nanmean(type_comp_data(:));

    type_total = total_scores(type_idx);
    performance_summary.평균종합점수(i) = nanmean(type_total);
end

% 엑셀 저장
try
    writetable(performance_summary, 'performance_ordered_analysis.xlsx', 'Sheet', '성과순서요약');

    top_comp_table = table();
    top_comp_table.순위 = (1:length(top_5_competencies))';
    top_comp_table.역량항목 = top_5_competencies';
    top_comp_table.가중치 = top_5_weights_normalized;
    top_comp_table.가중치_퍼센트 = top_5_weights_normalized * 100;
    top_comp_table.상관계수 = comp_correlations(top_5_idx);

    writetable(top_comp_table, 'performance_ordered_analysis.xlsx', 'Sheet', '핵심역량가중치');

    fprintf('✓ 엑셀 파일 저장: performance_ordered_analysis.xlsx\n');
catch
    fprintf('⚠ 엑셀 저장 실패\n');
end

% 매트랩 결과 저장
results = struct();
results.performance_summary = performance_summary;
results.top_competencies = top_5_competencies;
results.competency_weights = top_5_weights_normalized;
results.sorted_types = sorted_types;
results.data_usage_analysis = sprintf('전체 %d명 중 %d명 분석 (%.1f%%)', ...
    height(hr_data), length(matched_ids), length(matched_ids)/height(hr_data)*100);

save('performance_ordered_results.mat', 'results');
fprintf('✓ 매트랩 파일 저장: performance_ordered_results.mat\n');

%% 9. 최종 요약
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 80));
fprintf('                     성과순서 인재유형 분석 완료\n');
fprintf('%s\n', repmat('=', 1, 80));

fprintf('\n📊 데이터 사용량 분석 결과:\n');
fprintf('  • 전체 HR 데이터: %d명\n', height(hr_data));
fprintf('  • 인재유형 보유자: %d명 (%.1f%%)\n', height(hr_clean), height(hr_clean)/height(hr_data)*100);
fprintf('  • 최종 분석 대상: %d명 (%.1f%%)\n', length(matched_ids), length(matched_ids)/height(hr_data)*100);
fprintf('  • 데이터 손실 원인: 인재유형 누락(%d명), 역량검사 미실시(%d명)\n', ...
        height(hr_data)-height(hr_clean), height(hr_clean)-length(matched_ids));

fprintf('\n🏆 성과순위별 인재유형:\n');
for i = 1:length(sorted_types)
    type_count = sum(strcmp(matched_talent_types, sorted_types{i}));
    fprintf('  %d위. %s (%d명, 성과점수: %d)\n', i, sorted_types{i}, type_count, sorted_performance(i));
end

fprintf('\n🎯 핵심 역량 가중치 (상위 5개):\n');
for i = 1:length(top_5_competencies)
    fprintf('  %d. %s: %.1f%%\n', i, top_5_competencies{i}, top_5_weights_normalized(i)*100);
end

fprintf('\n✅ 성과순서 레이더 차트가 생성되었습니다!\n');
fprintf('%s\n', repmat('=', 1, 80));

%% Helper Function: Performance Radar Chart
function create_performance_radar(data, baseline, labels, title_text, perf_rank, color, count)
    % 각도 계산
    N = length(data);
    theta = linspace(0, 2*pi, N+1);

    % 데이터 준비 (원형으로 닫기)
    data_plot = [data, data(1)];
    baseline_plot = [baseline, baseline(1)];

    % 극좌표 변환
    [x_data, y_data] = pol2cart(theta, data_plot);
    [x_base, y_base] = pol2cart(theta, baseline_plot);

    % 플롯
    hold on;

    % 기준선 (전체 평균)
    plot(x_base, y_base, '--', 'Color', [0.5 0.5 0.5], 'LineWidth', 1.5);
    fill(x_base, y_base, [0.9 0.9 0.9], 'FaceAlpha', 0.3, 'EdgeColor', 'none');

    % 해당 인재유형 데이터
    plot(x_data, y_data, '-', 'Color', color, 'LineWidth', 3);
    fill(x_data, y_data, color, 'FaceAlpha', 0.4, 'EdgeColor', color, 'LineWidth', 2);

    % 축 설정
    axis equal;
    max_val = max([max(data), max(baseline)]) * 1.1;
    xlim([-max_val, max_val]);
    ylim([-max_val, max_val]);

    % 격자 그리기
    for r = 20:20:100
        [x_circle, y_circle] = pol2cart(theta, repmat(r, size(theta)));
        plot(x_circle, y_circle, ':', 'Color', [0.7 0.7 0.7], 'LineWidth', 0.5);
    end

    % 축선 그리기
    for i = 1:N
        plot([0, max_val*cos(theta(i))], [0, max_val*sin(theta(i))], ':', 'Color', [0.7 0.7 0.7]);

        % 라벨 추가
        label_r = max_val * 1.15;
        text(label_r*cos(theta(i)), label_r*sin(theta(i)), labels{i}, ...
             'HorizontalAlignment', 'center', 'VerticalAlignment', 'middle', ...
             'FontSize', 9, 'FontWeight', 'bold');
    end

    % 제목
    title(sprintf('%s\n(성과순위: %d위, %d명)', title_text, perf_rank, count), ...
          'FontSize', 12, 'FontWeight', 'bold', 'Color', color);

    % 범례
    legend({'전체평균', '', sprintf('%s', title_text), ''}, ...
           'Location', 'best', 'FontSize', 8);

    axis off;
    hold off;
end