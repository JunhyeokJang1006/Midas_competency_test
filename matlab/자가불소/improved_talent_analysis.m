%% Improved Talent Type Competency Analysis
% Using readtable and proper sheet selection for better efficiency

clear; clc;

%% 1. Data Loading - 개선된 방식
fprintf('=== 개선된 데이터 로딩 ===\n');

% 인적정보 데이터
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보.xlsx';
hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');

% 역량검사 데이터 - 상위항목 (3번째 시트)
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';
comp_upper = readtable(comp_file, 'Sheet', 3, 'VariableNamingRule', 'preserve');

% 역량검사 데이터 - 종합점수 (4번째 시트)
comp_total = readtable(comp_file, 'Sheet', 4, 'VariableNamingRule', 'preserve');

fprintf('인적정보 데이터: %d행 x %d열\n', height(hr_data), width(hr_data));
fprintf('상위항목 데이터: %d행 x %d열\n', height(comp_upper), width(comp_upper));
fprintf('종합점수 데이터: %d행 x %d열\n', height(comp_total), width(comp_total));

%% 2. 인재유형 데이터 정리
fprintf('\n=== 인재유형 데이터 정리 ===\n');

% 인재유형이 있는 데이터만 필터링
hr_clean = hr_data(~cellfun(@isempty, hr_data.('인재유형')), :);
fprintf('인재유형 데이터가 있는 직원: %d명\n', height(hr_clean));

% 인재유형별 분포
talent_types = hr_clean.('인재유형');
unique_types = unique(talent_types(~cellfun(@isempty, talent_types)));

fprintf('\n전체 인재유형 분포:\n');
for i = 1:length(unique_types)
    count = sum(strcmp(talent_types, unique_types{i}));
    fprintf('  %s: %d명\n', unique_types{i}, count);
end

%% 3. 상위항목 역량 데이터 분석
fprintf('\n=== 상위항목 역량 데이터 분석 ===\n');

% ID로 매칭 가능한 데이터 찾기
hr_ids = hr_clean.ID;
comp_upper_ids = comp_upper.ID;

% 매칭되는 ID 찾기
[matched_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_upper_ids);
fprintf('상위항목과 매칭된 ID: %d개\n', length(matched_ids));

if ~isempty(matched_ids)
    % 매칭된 데이터 추출
    matched_hr = hr_clean(hr_idx, :);
    matched_comp_upper = comp_upper(comp_idx, :);

    % 상위항목 역량 컬럼들 찾기 (숫자 데이터가 있는 컬럼)
    comp_cols = {};
    comp_col_indices = [];

    for i = 6:width(matched_comp_upper)  % 6번째 컬럼부터 역량 점수
        col_name = matched_comp_upper.Properties.VariableNames{i};
        col_data = matched_comp_upper{:, i};

        % 숫자 데이터이고 결측값이 50% 미만인 컬럼만 선택
        if isnumeric(col_data) && sum(~isnan(col_data)) >= length(col_data) * 0.5
            comp_cols{end+1} = col_name;
            comp_col_indices = [comp_col_indices, i];
        end
    end

    fprintf('유효한 상위항목 역량: %d개\n', length(comp_cols));
    fprintf('상위항목 역량들:\n');
    for i = 1:length(comp_cols)
        fprintf('  %d. %s\n', i, comp_cols{i});
    end

    % 상위항목 점수 데이터 추출
    upper_scores = matched_comp_upper{:, comp_col_indices};

    %% 4. 종합점수 분석
    fprintf('\n=== 종합점수 분석 ===\n');

    % 종합점수와 매칭
    comp_total_ids = comp_total.ID;
    [total_matched_ids, hr_total_idx, comp_total_idx] = intersect(hr_ids, comp_total_ids);
    fprintf('종합점수와 매칭된 ID: %d개\n', length(total_matched_ids));

    if ~isempty(total_matched_ids)
        matched_hr_total = hr_clean(hr_total_idx, :);
        matched_comp_total = comp_total(comp_total_idx, :);

        % 종합점수 컬럼 찾기 (보통 마지막 숫자 컬럼)
        total_score_col = width(matched_comp_total);
        total_scores = matched_comp_total{:, total_score_col};

        fprintf('종합점수 컬럼: %s\n', matched_comp_total.Properties.VariableNames{total_score_col});

        %% 5. 인재유형별 종합점수 분석
        fprintf('\n=== 인재유형별 종합점수 분석 ===\n');

        talent_types_total = matched_hr_total.('인재유형');
        unique_types_total = unique(talent_types_total(~cellfun(@isempty, talent_types_total)));

        fprintf('인재유형별 종합점수 통계:\n');
        fprintf('인재유형                | 평균점수 | 표준편차 | 최고점수 | 최저점수 | 인원수\n');
        fprintf('------------------------|----------|----------|----------|----------|-------\n');

        type_stats = [];
        for i = 1:length(unique_types_total)
            talent_type = unique_types_total{i};
            type_idx = strcmp(talent_types_total, talent_type) & ~isnan(total_scores);

            if sum(type_idx) > 0
                type_scores = total_scores(type_idx);

                mean_score = mean(type_scores);
                std_score = std(type_scores);
                max_score = max(type_scores);
                min_score = min(type_scores);
                count = length(type_scores);

                fprintf('%-22s | %8.2f | %8.2f | %8.2f | %8.2f | %6d\n', ...
                        talent_type, mean_score, std_score, max_score, min_score, count);

                type_stats = [type_stats; {talent_type, mean_score, std_score, max_score, min_score, count}];
            end
        end
    end

    %% 6. 상위항목별 인재유형 특성 분석
    fprintf('\n=== 상위항목별 인재유형 특성 분석 ===\n');

    % 인재유형 성과 순위 정의
    performance_ranking = containers.Map({...
        '자연성', '성실한 가연성', '유익한 불연성', '유능한 불연성', ...
        '게으른 가연성', '무능한 불연성', '위장형 소화성', '소화성'}, ...
        {8, 7, 6, 5, 4, 3, 2, 1});

    % 매칭된 인재유형과 상위항목 점수
    talent_types_upper = matched_hr.('인재유형');
    performance_scores = zeros(length(talent_types_upper), 1);

    for i = 1:length(talent_types_upper)
        talent_type = talent_types_upper{i};
        if performance_ranking.isKey(talent_type)
            performance_scores(i) = performance_ranking(talent_type);
        end
    end

    % 각 상위항목과 성과의 상관관계 계산
    fprintf('상위항목별 성과 예측력 분석:\n');
    fprintf('상위항목        | 상관계수 | 고성과평균 | 저성과평균 | 차이값 | 효과크기\n');
    fprintf('----------------|----------|------------|------------|--------|----------\n');

    correlations = [];
    competency_importance = [];

    for i = 1:length(comp_cols)
        comp_name = comp_cols{i};
        comp_scores = upper_scores(:, i);
        valid_idx = ~isnan(comp_scores) & performance_scores > 0;

        if sum(valid_idx) >= 10
            valid_comp_scores = comp_scores(valid_idx);
            valid_perf_scores = performance_scores(valid_idx);

            % 상관계수 계산
            if var(valid_comp_scores) > 0 && var(valid_perf_scores) > 0
                correlation = sum((valid_comp_scores - mean(valid_comp_scores)) .* ...
                                (valid_perf_scores - mean(valid_perf_scores))) / ...
                             (sqrt(sum((valid_comp_scores - mean(valid_comp_scores)).^2)) * ...
                              sqrt(sum((valid_perf_scores - mean(valid_perf_scores)).^2)));
            else
                correlation = 0;
            end

            % 고성과 vs 저성과 그룹 비교
            high_perf_threshold = median(valid_perf_scores);
            high_idx = valid_perf_scores > high_perf_threshold;
            low_idx = valid_perf_scores <= high_perf_threshold;

            high_mean = mean(valid_comp_scores(high_idx));
            low_mean = mean(valid_comp_scores(low_idx));
            difference = high_mean - low_mean;

            % 효과 크기 (Cohen's d)
            pooled_std = sqrt(((sum(high_idx)-1)*var(valid_comp_scores(high_idx)) + ...
                              (sum(low_idx)-1)*var(valid_comp_scores(low_idx))) / ...
                             (sum(high_idx) + sum(low_idx) - 2));
            effect_size = difference / (pooled_std + eps);

            fprintf('%-15s | %8.3f | %10.2f | %10.2f | %6.2f | %8.3f\n', ...
                    comp_name, correlation, high_mean, low_mean, difference, effect_size);

            correlations = [correlations; correlation];
            competency_importance = [competency_importance; abs(correlation) * 0.6 + abs(effect_size) * 0.4];
        end
    end

    %% 7. 최종 가중치 및 추천사항
    fprintf('\n=== 성과 예측을 위한 상위항목 가중치 ===\n');

    % 중요도 순으로 정렬
    [sorted_importance, sort_idx] = sort(competency_importance, 'descend');

    fprintf('순위 | 상위항목        | 가중치 | 상관계수 | 효과크기\n');
    fprintf('-----|-----------------|--------|----------|----------\n');

    top_competencies = {};
    top_weights = [];

    for i = 1:min(5, length(sort_idx))
        idx = sort_idx(i);
        comp_name = comp_cols{idx};
        weight = sorted_importance(i);
        correlation = correlations(idx);

        % 효과크기 재계산
        comp_scores = upper_scores(:, idx);
        valid_idx = ~isnan(comp_scores) & performance_scores > 0;
        valid_comp_scores = comp_scores(valid_idx);
        valid_perf_scores = performance_scores(valid_idx);

        high_perf_threshold = median(valid_perf_scores);
        high_idx = valid_perf_scores > high_perf_threshold;
        low_idx = valid_perf_scores <= high_perf_threshold;

        high_mean = mean(valid_comp_scores(high_idx));
        low_mean = mean(valid_comp_scores(low_idx));
        difference = high_mean - low_mean;

        pooled_std = sqrt(((sum(high_idx)-1)*var(valid_comp_scores(high_idx)) + ...
                          (sum(low_idx)-1)*var(valid_comp_scores(low_idx))) / ...
                         (sum(high_idx) + sum(low_idx) - 2));
        effect_size = difference / (pooled_std + eps);

        fprintf('%4d | %-15s | %6.3f | %8.3f | %8.3f\n', ...
                i, comp_name, weight, correlation, effect_size);

        top_competencies{end+1} = comp_name;
        top_weights = [top_weights; weight];
    end

    % 가중치 정규화 (합이 1이 되도록)
    if ~isempty(top_weights)
        top_weights_normalized = top_weights / sum(top_weights);

        fprintf('\n=== 최종 성과 예측 가중치 (정규화) ===\n');
        for i = 1:length(top_competencies)
            fprintf('%d. %s: %.1f%%\n', i, top_competencies{i}, top_weights_normalized(i) * 100);
        end
    end

    %% 8. 결과 저장
    fprintf('\n=== 결과 저장 ===\n');

    % 결과 구조체 생성
    results = struct();
    results.type_stats = type_stats;
    results.top_competencies = top_competencies;
    results.top_weights_normalized = top_weights_normalized;
    results.correlations = correlations;
    results.comp_cols = comp_cols;
    results.analysis_summary = sprintf('총 %d명 분석, %d개 상위항목 역량', ...
        length(matched_ids), length(comp_cols));

    % 파일 저장
    save('improved_talent_analysis_results.mat', 'results');

    % 엑셀로도 저장
    try
        % 인재유형별 통계 테이블
        type_table = cell2table(type_stats, 'VariableNames', ...
            {'인재유형', '평균점수', '표준편차', '최고점수', '최저점수', '인원수'});
        writetable(type_table, 'talent_type_statistics.xlsx', 'Sheet', '인재유형별통계');

        % 상위항목 가중치 테이블
        weight_table = table();
        weight_table.순위 = (1:length(top_competencies))';
        weight_table.상위항목 = top_competencies';
        weight_table.가중치 = top_weights_normalized;
        weight_table.가중치_퍼센트 = top_weights_normalized * 100;
        writetable(weight_table, 'talent_type_statistics.xlsx', 'Sheet', '상위항목가중치');

        fprintf('엑셀 파일 저장 완료: talent_type_statistics.xlsx\n');
    catch
        fprintf('엑셀 저장 실패 (Excel이 없을 수 있음)\n');
    end

    fprintf('MATLAB 파일 저장 완료: improved_talent_analysis_results.mat\n');

else
    fprintf('경고: 매칭되는 데이터가 없습니다.\n');
end

%% 9. 최종 요약
fprintf('\n');
fprintf('%s\n', repmat('=', 1, 70));
fprintf('                    개선된 인재유형 분석 완료\n');
fprintf('%s\n', repmat('=', 1, 70));

if exist('type_stats', 'var') && ~isempty(type_stats)
    fprintf('\n📊 인재유형별 종합점수 순위:\n');
    % 평균점수로 정렬
    type_scores = cell2mat(type_stats(:, 2));
    [~, rank_idx] = sort(type_scores, 'descend');

    for i = 1:length(rank_idx)
        idx = rank_idx(i);
        fprintf('  %d. %s: %.2f점 (%d명)\n', i, type_stats{idx, 1}, ...
                type_stats{idx, 2}, type_stats{idx, 6});
    end
end

if exist('top_competencies', 'var') && ~isempty(top_competencies)
    fprintf('\n🎯 성과 예측 핵심 상위항목:\n');
    for i = 1:length(top_competencies)
        fprintf('  %d. %s: %.1f%%\n', i, top_competencies{i}, ...
                top_weights_normalized(i) * 100);
    end
end

fprintf('\n✅ 분석 완료! readtable과 시트별 접근으로 더 효율적인 분석이 되었습니다.\n');
fprintf('%s\n', repmat('=', 1, 70));