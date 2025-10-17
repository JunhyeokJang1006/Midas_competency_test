% 간단한 데이터 손실 추적 스크립트
% 원본 코드의 로직을 그대로 따라가면서 각 단계별 손실을 기록

clc; clear; close all;

fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
fprintf('║           단계별 데이터 손실 추적 (간단 버전)             ║\n');
fprintf('╚════════════════════════════════════════════════════════════╝\n\n');

% 파일 경로
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사_개발자추가.xlsx';

%% 단계 1: 초기 로딩
fprintf('【단계 1】 초기 데이터 로딩\n');
fprintf('════════════════════════════════════════════════════════════\n');

hr_data = readtable(hr_file, 'Sheet', 1, 'VariableNamingRule', 'preserve');
comp_upper = readtable(comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');

fprintf('HR 데이터: %d명\n', height(hr_data));
fprintf('역량검사 데이터: %d명\n\n', height(comp_upper));

%% 단계 2: 신뢰가능성 필터링
fprintf('【단계 2】 신뢰가능성 필터링\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 원본 코드 방식: '신뢰불가' 문자열 찾기
reliability_col_idx = find(contains(comp_upper.Properties.VariableNames, '신뢰가능성'), 1);
if ~isempty(reliability_col_idx)
    reliability_data = comp_upper{:, reliability_col_idx};
    if iscell(reliability_data)
        unreliable_idx = strcmp(reliability_data, '신뢰불가');
    else
        unreliable_idx = false(height(comp_upper), 1);
    end
    comp_upper = comp_upper(~unreliable_idx, :);
    fprintf('신뢰가능: %d명 (제외: %d명)\n\n', height(comp_upper), sum(unreliable_idx));
else
    fprintf('신뢰가능성 컬럼 없음\n\n');
end

%% 단계 3: 인재유형 추출 및 정제
fprintf('【단계 3】 인재유형 정제 (위장형 소화성 제외)\n');
fprintf('════════════════════════════════════════════════════════════\n');

talent_col_idx = find(contains(hr_data.Properties.VariableNames, {'인재유형'}), 1);
hr_clean = hr_data(~cellfun(@isempty, hr_data{:, talent_col_idx}), :);
excluded_mask = strcmp(hr_clean{:, talent_col_idx}, '위장형 소화성');
hr_clean = hr_clean(~excluded_mask, :);

talent_types = hr_clean{:, talent_col_idx};
[unique_types, ~, type_indices] = unique(talent_types);
type_counts_stage1 = accumarray(type_indices, 1);

fprintf('인재유형별 초기 분포:\n');
fprintf('  %-28s: %4s\n', '인재유형', '인원');
fprintf('  %s\n', repmat('-', 1, 38));
for i = 1:length(unique_types)
    fprintf('  %-28s: %4d명\n', unique_types{i}, type_counts_stage1(i));
end
fprintf('  %s\n', repmat('-', 1, 38));
fprintf('  %-28s: %4d명\n\n', '【전체】', height(hr_clean));

%% 단계 4: ID 매칭
fprintf('【단계 4】 ID 매칭\n');
fprintf('════════════════════════════════════════════════════════════\n');

comp_id_col = find(contains(lower(comp_upper.Properties.VariableNames), {'id', '사번'}), 1);

% ID 표준화 (원본 코드 방식)
hr_ids_str = arrayfun(@(x) sprintf('%.0f', x), hr_clean.ID, 'UniformOutput', false);
comp_ids_str = arrayfun(@(x) sprintf('%.0f', x), comp_upper{:, comp_id_col}, 'UniformOutput', false);

% 교집합 찾기
[matched_ids, hr_idx, comp_idx] = intersect(hr_ids_str, comp_ids_str);

matched_hr = hr_clean(hr_idx, :);
matched_comp = comp_upper(comp_idx, :);
matched_talent_types = matched_hr{:, talent_col_idx};

fprintf('매칭 성공: %d명\n\n', length(matched_ids));

% 매칭 후 유형별 분포
[~, ~, type_indices_matched] = unique(matched_talent_types);
type_counts_stage2 = accumarray(type_indices_matched, 1);
unique_types_matched = unique(matched_talent_types);

fprintf('인재유형별 변화:\n');
fprintf('  %-28s | %4s | %6s | %4s | %7s\n', '인재유형', '초기', '매칭후', '손실', '손실률');
fprintf('  %s\n', repmat('-', 1, 68));

for i = 1:length(unique_types)
    type_name = unique_types{i};
    count_before = type_counts_stage1(i);

    % 매칭 후 카운트
    match_idx = find(strcmp(unique_types_matched, type_name));
    if isempty(match_idx)
        count_after = 0;
    else
        count_after = type_counts_stage2(match_idx);
    end

    loss = count_before - count_after;
    loss_rate = loss / count_before * 100;

    fprintf('  %-28s | %4d | %6d | %4d | %6.1f%%\n', type_name, count_before, ...
        count_after, loss, loss_rate);
end

total_before = height(hr_clean);
total_after = length(matched_ids);
total_loss = total_before - total_after;
total_loss_rate = total_loss / total_before * 100;

fprintf('  %s\n', repmat('-', 1, 68));
fprintf('  %-28s | %4d | %6d | %4d | %6.1f%%\n\n', '【전체】', total_before, ...
    total_after, total_loss, total_loss_rate);

%% 단계 5: 결측값 분석
fprintf('【단계 5】 결측값 분석\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 역량 컬럼만 추출
valid_comp_indices = [];
for i = 6:width(matched_comp)
    col_data = matched_comp{:, i};
    if isnumeric(col_data) && ~all(isnan(col_data))
        valid_data = col_data(~isnan(col_data));
        if length(valid_data) >= 5
            valid_comp_indices(end+1) = i;
        end
    end
end

X_raw = matched_comp{:, valid_comp_indices};
num_competencies = size(X_raw, 2);
missing_count_per_sample = sum(isnan(X_raw), 2);  % 각 사람별 결측 개수

fprintf('역량 항목 수: %d개\n', num_competencies);
fprintf('전체 결측률: %.1f%%\n\n', sum(isnan(X_raw(:)))/numel(X_raw)*100);

% 결측 개수별 분포 (인재유형별로 세분화)
fprintf('【결측 개수별 분포】\n');
fprintf('  %-28s | 결측0 | 결측1 | 결측2 | 결측3+ | 합계\n', '인재유형');
fprintf('  %s\n', repmat('-', 1, 80));

% 각 인재유형별 결측 분포
for i = 1:length(unique_types_matched)
    type_name = unique_types_matched{i};
    type_mask = strcmp(matched_talent_types, type_name);

    missing_0 = sum(type_mask & missing_count_per_sample == 0);
    missing_1 = sum(type_mask & missing_count_per_sample == 1);
    missing_2 = sum(type_mask & missing_count_per_sample == 2);
    missing_3plus = sum(type_mask & missing_count_per_sample >= 3);
    total = sum(type_mask);

    fprintf('  %-28s | %5d | %5d | %5d | %6d | %4d\n', ...
        type_name, missing_0, missing_1, missing_2, missing_3plus, total);
end

% 전체 합계
total_missing_0 = sum(missing_count_per_sample == 0);
total_missing_1 = sum(missing_count_per_sample == 1);
total_missing_2 = sum(missing_count_per_sample == 2);
total_missing_3plus = sum(missing_count_per_sample >= 3);
total_all = length(missing_count_per_sample);

fprintf('  %s\n', repmat('-', 1, 80));
fprintf('  %-28s | %5d | %5d | %5d | %6d | %4d\n', ...
    '【전체】', total_missing_0, total_missing_1, total_missing_2, total_missing_3plus, total_all);
fprintf('\n');

% 결측 개수별 상세 분포 (0개부터 10개까지 모두)
fprintf('【실제 결측 분포 (역량 %d개 기준)】\n', num_competencies);
fprintf('  결측 개수 | 인원 | 비율  | 결측률 | 상태\n');
fprintf('  %s\n', repmat('-', 1, 55));

for missing_cnt = 0:num_competencies
    count = sum(missing_count_per_sample == missing_cnt);
    if count > 0 || missing_cnt <= 5  % 0-5개는 무조건 표시, 그 이상은 있을 때만
        missing_pct = missing_cnt / num_competencies * 100;

        % 상태 표시
        if missing_cnt == 0
            status = '✓ 완전응답';
        elseif missing_cnt < 3
            status = '✓ 유지';
        elseif missing_cnt == 3
            status = '⚠ 경계선';
        else
            status = '✗ 제거대상';
        end

        fprintf('  결측 %2d개 | %4d | %5.1f%% | %5.1f%% | %s\n', ...
            missing_cnt, count, count/total_all*100, missing_pct, status);
    end
end
fprintf('  %s\n', repmat('-', 1, 55));
fprintf('  합계      | %4d | 100.0%%\n\n', total_all);

% 제거 대상 분석
threshold_30pct = ceil(num_competencies * 0.3);  % 30% 기준 (올림)
will_be_removed = missing_count_per_sample >= threshold_30pct;

fprintf('【30%% 제거 기준 상세】\n');
fprintf('  제거 기준: 결측 %d개 이상 (%.1f%%)\n', threshold_30pct, 30.0);
fprintf('  제거 예정: %d명\n', sum(will_be_removed));
fprintf('  유지 예정: %d명\n\n', sum(~will_be_removed));

% 역량별 결측 분석
fprintf('【역량 항목별 결측 통계】\n');
comp_names = matched_comp.Properties.VariableNames(valid_comp_indices);
missing_per_competency = sum(isnan(X_raw), 1);

fprintf('  %-20s | 결측수 | 결측률\n', '역량 항목');
fprintf('  %s\n', repmat('-', 1, 45));
for i = 1:length(comp_names)
    fprintf('  %-20s | %6d | %6.1f%%\n', ...
        comp_names{i}, missing_per_competency(i), missing_per_competency(i)/size(X_raw,1)*100);
end
fprintf('  %s\n', repmat('-', 1, 45));
fprintf('  %-20s | %6d | %6.1f%%\n\n', '평균', ...
    mean(missing_per_competency), mean(missing_per_competency)/size(X_raw,1)*100);

% 제거될 사람들의 상세 정보
if sum(will_be_removed) > 0 && sum(will_be_removed) <= 30
    fprintf('【제거 대상자 상세 (결측 3개 이상)】\n');
    fprintf('  번호 | %-28s | 결측수 | 결측 역량\n', '인재유형');
    fprintf('  %s\n', repmat('-', 1, 100));

    removed_indices = find(will_be_removed);
    for idx = 1:length(removed_indices)
        person_idx = removed_indices(idx);
        type_name = matched_talent_types{person_idx};
        missing_cnt = missing_count_per_sample(person_idx);

        % 이 사람의 결측 역량 찾기
        missing_comps = find(isnan(X_raw(person_idx, :)));
        missing_comp_names = comp_names(missing_comps);

        fprintf('  %4d | %-28s | %6d | %s\n', ...
            idx, type_name, missing_cnt, strjoin(missing_comp_names, ', '));
    end
    fprintf('\n');
end

%% 단계 6: 결측값 제거 (30% 임계값)
fprintf('【단계 6】 결측값 제거 (2단계 프로세스)\n');
fprintf('════════════════════════════════════════════════════════════\n');

% Step 1: 결측률 30% 이상 제거
missing_per_sample = missing_count_per_sample / num_competencies;
high_quality_idx = missing_per_sample < 0.3;

fprintf('Step 1: 고품질 샘플 선택 (결측률 < 30%%)\n');
fprintf('  제거: %d명\n', sum(~high_quality_idx));
fprintf('  유지: %d명\n\n', sum(high_quality_idx));

X_filtered = X_raw(high_quality_idx, :);
filtered_talent_types = matched_talent_types(high_quality_idx);

% 제거된 사람들의 인재유형 분석
fprintf('  제거된 인원의 인재유형 분포:\n');
for i = 1:length(unique_types_matched)
    type_name = unique_types_matched{i};
    removed_count = sum(strcmp(matched_talent_types(~high_quality_idx), type_name));
    if removed_count > 0
        fprintf('    %-28s: %d명\n', type_name, removed_count);
    end
end
fprintf('\n');

% Step 2: 완전한 케이스만 사용
complete_idx = ~any(isnan(X_filtered), 2);
fprintf('Step 2: 완전한 케이스만 사용 (결측 0개)\n');
fprintf('  추가 제거: %d명\n', sum(~complete_idx));

if sum(~complete_idx) > 0
    fprintf('  추가 제거된 인원의 인재유형:\n');
    for i = 1:length(unique_types_matched)
        type_name = unique_types_matched{i};
        additional_removed = sum(strcmp(filtered_talent_types(~complete_idx), type_name));
        if additional_removed > 0
            fprintf('    %-28s: %d명\n', type_name, additional_removed);
        end
    end
end

final_talent_types = filtered_talent_types(complete_idx);
final_count = length(final_talent_types);

fprintf('  최종 샘플: %d명\n\n', final_count);

% 최종 유형별 분포
[unique_types_final, ~, type_indices_final] = unique(final_talent_types);
type_counts_stage3 = accumarray(type_indices_final, 1);

fprintf('인재유형별 최종 분포:\n');
fprintf('  %-28s | %4s | %6s | %8s | %4s | %7s\n', '인재유형', '초기', '매칭후', ...
    '결측제거', '손실', '손실률');
fprintf('  %s\n', repmat('-', 1, 80));

for i = 1:length(unique_types)
    type_name = unique_types{i};
    count_stage1 = type_counts_stage1(i);

    % 매칭 후 카운트
    match_idx = find(strcmp(unique_types_matched, type_name));
    if isempty(match_idx)
        count_stage2 = 0;
    else
        count_stage2 = type_counts_stage2(match_idx);
    end

    % 최종 카운트
    final_idx = find(strcmp(unique_types_final, type_name));
    if isempty(final_idx)
        count_stage3 = 0;
    else
        count_stage3 = type_counts_stage3(final_idx);
    end

    loss = count_stage1 - count_stage3;
    loss_rate = loss / count_stage1 * 100;

    fprintf('  %-28s | %4d | %6d | %8d | %4d | %6.1f%%\n', type_name, ...
        count_stage1, count_stage2, count_stage3, loss, loss_rate);
end

total_loss_final = total_before - final_count;
total_loss_rate_final = total_loss_final / total_before * 100;

fprintf('  %s\n', repmat('-', 1, 80));
fprintf('  %-28s | %4d | %6d | %8d | %4d | %6.1f%%\n\n', '【전체】', ...
    total_before, total_after, final_count, total_loss_final, total_loss_rate_final);

%% 단계 7: 고성과자 그룹 추적
fprintf('【단계 7】 고성과자/저성과자 추적\n');
fprintf('════════════════════════════════════════════════════════════\n');

high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
low_performers = {'무능한 불연성', '소화성', '게으른 가연성'};

% 초기 고성과자
high_stage1 = 0;
for i = 1:length(high_performers)
    idx = find(strcmp(unique_types, high_performers{i}));
    if ~isempty(idx)
        high_stage1 = high_stage1 + type_counts_stage1(idx);
    end
end

% 매칭 후 고성과자
high_stage2 = 0;
for i = 1:length(high_performers)
    idx = find(strcmp(unique_types_matched, high_performers{i}));
    if ~isempty(idx)
        high_stage2 = high_stage2 + type_counts_stage2(idx);
    end
end

% 최종 고성과자
high_stage3 = 0;
for i = 1:length(high_performers)
    idx = find(strcmp(unique_types_final, high_performers{i}));
    if ~isempty(idx)
        high_stage3 = high_stage3 + type_counts_stage3(idx);
    end
end

fprintf('【고성과자】\n');
fprintf('  초기: %d명 → 매칭 후: %d명 → 최종: %d명\n', high_stage1, high_stage2, high_stage3);
fprintf('  총 손실: %d명 (%.1f%%)\n\n', high_stage1 - high_stage3, ...
    (high_stage1 - high_stage3) / high_stage1 * 100);

% 저성과자
low_stage1 = 0;
for i = 1:length(low_performers)
    idx = find(strcmp(unique_types, low_performers{i}));
    if ~isempty(idx)
        low_stage1 = low_stage1 + type_counts_stage1(idx);
    end
end

low_stage2 = 0;
for i = 1:length(low_performers)
    idx = find(strcmp(unique_types_matched, low_performers{i}));
    if ~isempty(idx)
        low_stage2 = low_stage2 + type_counts_stage2(idx);
    end
end

low_stage3 = 0;
for i = 1:length(low_performers)
    idx = find(strcmp(unique_types_final, low_performers{i}));
    if ~isempty(idx)
        low_stage3 = low_stage3 + type_counts_stage3(idx);
    end
end

fprintf('【저성과자】\n');
fprintf('  초기: %d명 → 매칭 후: %d명 → 최종: %d명\n', low_stage1, low_stage2, low_stage3);
fprintf('  총 손실: %d명 (%.1f%%)\n\n', low_stage1 - low_stage3, ...
    (low_stage1 - low_stage3) / low_stage1 * 100);

fprintf('최종 분석 데이터:\n');
fprintf('  고성과자: %d명 (%.1f%%)\n', high_stage3, high_stage3/final_count*100);
fprintf('  저성과자: %d명 (%.1f%%)\n', low_stage3, low_stage3/final_count*100);
fprintf('  클래스 비율: %.2f:1\n\n', high_stage3/low_stage3);

fprintf('✅ 추적 완료\n\n');
