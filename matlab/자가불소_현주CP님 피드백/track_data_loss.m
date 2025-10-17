% 단계별 데이터 손실 추적 스크립트
% 각 전처리 단계에서 어떤 인재유형이 얼마나 제외되는지 추적

clc; clear; close all;

% 파일 경로 설정
config = struct();
config.hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
config.comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

fprintf('\n╔════════════════════════════════════════════════════════════╗\n');
fprintf('║           단계별 데이터 손실 추적 분석                    ║\n');
fprintf('╚════════════════════════════════════════════════════════════╝\n\n');

%% ========================================================================
%                           단계 1: 초기 데이터 로딩
% =========================================================================
fprintf('【단계 1】 초기 데이터 로딩\n');
fprintf('════════════════════════════════════════════════════════════\n');

% HR 데이터 로딩
hr_data = readtable(config.hr_file, 'VariableNamingRule', 'preserve');
fprintf('✓ HR 데이터 로드: %d명\n', height(hr_data));

% 역량검사 데이터 로딩
comp_data_original = readtable(config.comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');
fprintf('✓ 역량검사 데이터 로드: %d명\n\n', height(comp_data_original));

% 인재유형 초기 분포
talent_types = hr_data.("인재유형");
valid_talent_idx = ~cellfun(@isempty, talent_types) & ~strcmp(talent_types, '');

% 위장형 소화성 제외 (원본 코드 방식)
excluded_mask = strcmp(talent_types, '위장형 소화성');
hr_data_clean = hr_data(valid_talent_idx & ~excluded_mask, :);
talent_types_clean = hr_data_clean.("인재유형");

unique_types = unique(talent_types_clean);

stage1_counts = containers.Map();
fprintf('인재유형별 초기 분포:\n');
fprintf('  %-28s: %4s\n', '인재유형', '인원');
fprintf('  %s\n', repmat('-', 1, 38));

for i = 1:length(unique_types)
    type_name = unique_types{i};
    count = sum(strcmp(talent_types_clean, type_name));
    stage1_counts(type_name) = count;
    fprintf('  %-28s: %4d명\n', type_name, count);
end
fprintf('  %s\n', repmat('-', 1, 38));
fprintf('  %-28s: %4d명\n\n', '【전체】', height(hr_data_clean));

%% ========================================================================
%                           단계 2: ID 매칭 (신뢰가능성 필터링 전)
% =========================================================================
fprintf('【단계 2】 ID 매칭 (신뢰가능성 필터링 전)\n');
fprintf('════════════════════════════════════════════════════════════\n');

% ID 컬럼명 확인 및 자동 매칭
hr_cols = hr_data_clean.Properties.VariableNames;
comp_cols = comp_data_original.Properties.VariableNames;

% ID 컬럼 찾기
hr_id_col = '';
comp_id_col = '';

for i = 1:length(hr_cols)
    if contains(lower(hr_cols{i}), 'id')
        hr_id_col = hr_cols{i};
        break;
    end
end

for i = 1:length(comp_cols)
    if contains(lower(comp_cols{i}), 'id')
        comp_id_col = comp_cols{i};
        break;
    end
end

fprintf('HR ID 컬럼: %s\n', hr_id_col);
fprintf('역량검사 ID 컬럼: %s\n', comp_id_col);

% ID를 문자열로 변환하여 매칭 (원본 코드 방식)
% hr_data_clean 사용 (위장형 소화성 제외 후 데이터)
hr_ids_raw = hr_data_clean.(hr_id_col);
comp_ids_raw = comp_data_original.(comp_id_col);

% 숫자를 문자열로 변환
if isnumeric(hr_ids_raw)
    hr_ids = arrayfun(@(x) sprintf('%.0f', x), hr_ids_raw, 'UniformOutput', false);
else
    hr_ids = hr_ids_raw;
end

if isnumeric(comp_ids_raw)
    comp_ids = arrayfun(@(x) sprintf('%.0f', x), comp_ids_raw, 'UniformOutput', false);
else
    comp_ids = comp_ids_raw;
end

[matched_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_ids);
fprintf('✓ 매칭 성공: %d명\n', length(matched_ids));
fprintf('  (HR에서 %d명, 역량검사에서 %d명)\n\n', length(hr_idx), length(comp_idx));

% 매칭된 데이터
matched_hr = hr_data_clean(hr_idx, :);
matched_comp_with_unreliable = comp_data_original(comp_idx, :);
matched_talent_types = matched_hr.("인재유형");

% 매칭 후 분포
stage2_counts = containers.Map();
fprintf('인재유형별 변화:\n');
fprintf('  %-28s | %4s | %6s | %4s | %7s\n', '인재유형', '초기', '매칭후', '손실', '손실률');
fprintf('  %s\n', repmat('-', 1, 68));

for i = 1:length(unique_types)
    type_name = unique_types{i};
    count_before = stage1_counts(type_name);
    count_after = sum(strcmp(matched_talent_types, type_name));
    stage2_counts(type_name) = count_after;
    loss = count_before - count_after;
    loss_rate = loss / count_before * 100;

    fprintf('  %-28s | %4d | %6d | %4d | %6.1f%%\n', type_name, count_before, ...
        count_after, loss, loss_rate);
end

total_before = height(hr_data_clean);
total_after = length(matched_ids);
total_loss = total_before - total_after;
total_loss_rate = total_loss / total_before * 100;

fprintf('  %s\n', repmat('-', 1, 68));
fprintf('  %-28s | %4d | %6d | %4d | %6.1f%%\n\n', '【전체】', total_before, ...
    total_after, total_loss, total_loss_rate);

%% ========================================================================
%                        단계 4: 역량 데이터 결측값 확인
% =========================================================================
fprintf('【단계 4】 역량 데이터 결측값 분석\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 역량 컬럼 추출
comp_cols = matched_comp.Properties.VariableNames;
exclude_cols = {'ID', '대상자ID', '코드', '인재유형', '성별', '나이', '개발자여부', '신뢰가능성', '코호트'};
comp_col_idx = ~ismember(comp_cols, exclude_cols);
comp_feature_names = comp_cols(comp_col_idx);

fprintf('역량 컬럼 수: %d개\n', length(comp_feature_names));
fprintf('분석 대상: %s\n\n', strjoin(comp_feature_names, ', '));

% 역량 데이터만 추출 (숫자형으로 변환)
X_raw = zeros(height(matched_comp), length(comp_feature_names));
for i = 1:length(comp_feature_names)
    col_data = matched_comp.(comp_feature_names{i});
    if iscell(col_data)
        % 셀 배열을 숫자로 변환
        X_raw(:, i) = cellfun(@(x) str2double(x), col_data);
    elseif isnumeric(col_data)
        X_raw(:, i) = col_data;
    else
        X_raw(:, i) = double(col_data);
    end
end

% 샘플별 결측값 비율 계산
missing_per_sample = sum(isnan(X_raw), 2) / size(X_raw, 2);
fprintf('결측값 통계:\n');
fprintf('  전체 결측률: %.1f%%\n', sum(isnan(X_raw(:)))/numel(X_raw)*100);
fprintf('  완전한 케이스: %d명\n', sum(missing_per_sample == 0));
fprintf('  결측 있는 케이스: %d명\n\n', sum(missing_per_sample > 0));

% 결측률별 분포
missing_thresholds = [0, 0.1, 0.3, 0.5, 1.0];
fprintf('결측률별 샘플 분포:\n');
for i = 1:length(missing_thresholds)-1
    lower = missing_thresholds(i);
    upper = missing_thresholds(i+1);
    count = sum(missing_per_sample > lower & missing_per_sample <= upper);
    fprintf('  %3.0f%% < 결측률 ≤ %3.0f%%: %3d명\n', lower*100, upper*100, count);
end
fprintf('\n');

%% ========================================================================
%                    단계 5: 결측값 제거 (30% 임계값)
% =========================================================================
fprintf('【단계 5】 결측값 제거 (30%% 임계값)\n');
fprintf('════════════════════════════════════════════════════════════\n');

% 결측률 30% 이상 제거
high_quality_idx = missing_per_sample < 0.3;
fprintf('결측률 30%% 이상 제거: %d명\n', sum(~high_quality_idx));

filtered_talent_types = matched_talent_types(high_quality_idx);
filtered_comp = matched_comp(high_quality_idx, :);
X_filtered = X_raw(high_quality_idx, :);

% 완전한 케이스만 사용
complete_idx = ~any(isnan(X_filtered), 2);
fprintf('추가 결측값 제거: %d명\n', sum(~complete_idx));
fprintf('최종 샘플: %d명\n\n', sum(complete_idx));

final_talent_types = filtered_talent_types(complete_idx);

% 최종 분포
stage3_counts = containers.Map();
fprintf('인재유형별 최종 분포:\n');
fprintf('  %-28s | %4s | %6s | %6s | %4s | %7s\n', '인재유형', '초기', '매칭후', ...
    '결측제거', '손실', '손실률');
fprintf('  %s\n', repmat('-', 1, 80));

for i = 1:length(unique_types)
    type_name = unique_types{i};
    count_stage1 = stage1_counts(type_name);
    count_stage2 = stage2_counts(type_name);
    count_stage3 = sum(strcmp(final_talent_types, type_name));
    stage3_counts(type_name) = count_stage3;
    loss = count_stage1 - count_stage3;
    loss_rate = loss / count_stage1 * 100;

    fprintf('  %-28s | %4d | %6d | %8d | %4d | %6.1f%%\n', type_name, ...
        count_stage1, count_stage2, count_stage3, loss, loss_rate);
end

final_count = sum(complete_idx);
total_loss_final = total_before - final_count;
total_loss_rate_final = total_loss_final / total_before * 100;

fprintf('  %s\n', repmat('-', 1, 80));
fprintf('  %-28s | %4d | %6d | %8d | %4d | %6.1f%%\n\n', '【전체】', ...
    total_before, total_after, final_count, total_loss_final, total_loss_rate_final);

%% ========================================================================
%                      단계 6: 고성과자 그룹 추적
% =========================================================================
fprintf('【단계 6】 고성과자 그룹 변화 추적\n');
fprintf('════════════════════════════════════════════════════════════\n');

high_performers = {'성실한 가연성', '자연성', '유익한 불연성'};
low_performers = {'무능한 불연성', '소화성', '게으른 가연성'};
excluded_group = {'유능한 불연성'};

% 고성과자 추적
fprintf('【고성과자】 (성실한 가연성 + 자연성 + 유익한 불연성)\n');
fprintf('  %-28s | %4s | %6s | %8s\n', '단계', '인원', '손실', '손실률');
fprintf('  %s\n', repmat('-', 1, 55));

high_stage1 = 0;
high_stage2 = 0;
high_stage3 = 0;

for i = 1:length(high_performers)
    type_name = high_performers{i};
    if stage1_counts.isKey(type_name)
        high_stage1 = high_stage1 + stage1_counts(type_name);
    end
    if stage2_counts.isKey(type_name)
        high_stage2 = high_stage2 + stage2_counts(type_name);
    end
    if stage3_counts.isKey(type_name)
        high_stage3 = high_stage3 + stage3_counts(type_name);
    end
end

fprintf('  %-28s | %4d |      - |       -\n', '1. 초기 데이터', high_stage1);
fprintf('  %-28s | %4d | %6d | %6.1f%%\n', '2. 매칭 후', high_stage2, ...
    high_stage1 - high_stage2, (high_stage1 - high_stage2)/high_stage1*100);
fprintf('  %-28s | %4d | %6d | %6.1f%%\n\n', '3. 결측값 제거 후', high_stage3, ...
    high_stage1 - high_stage3, (high_stage1 - high_stage3)/high_stage1*100);

% 저성과자 추적
fprintf('【저성과자】 (무능한 불연성 + 소화성 + 게으른 가연성)\n');
fprintf('  %-28s | %4s | %6s | %8s\n', '단계', '인원', '손실', '손실률');
fprintf('  %s\n', repmat('-', 1, 55));

low_stage1 = 0;
low_stage2 = 0;
low_stage3 = 0;

for i = 1:length(low_performers)
    type_name = low_performers{i};
    if stage1_counts.isKey(type_name)
        low_stage1 = low_stage1 + stage1_counts(type_name);
    end
    if stage2_counts.isKey(type_name)
        low_stage2 = low_stage2 + stage2_counts(type_name);
    end
    if stage3_counts.isKey(type_name)
        low_stage3 = low_stage3 + stage3_counts(type_name);
    end
end

fprintf('  %-28s | %4d |      - |       -\n', '1. 초기 데이터', low_stage1);
fprintf('  %-28s | %4d | %6d | %6.1f%%\n', '2. 매칭 후', low_stage2, ...
    low_stage1 - low_stage2, (low_stage1 - low_stage2)/low_stage1*100);
fprintf('  %-28s | %4d | %6d | %6.1f%%\n\n', '3. 결측값 제거 후', low_stage3, ...
    low_stage1 - low_stage3, (low_stage1 - low_stage3)/low_stage1*100);

% 제외 그룹 추적
fprintf('【분석 제외】 (유능한 불연성)\n');
fprintf('  %-28s | %4s\n', '단계', '인원');
fprintf('  %s\n', repmat('-', 1, 38));

excluded_stage1 = 0;
excluded_stage2 = 0;
excluded_stage3 = 0;

for i = 1:length(excluded_group)
    type_name = excluded_group{i};
    if stage1_counts.isKey(type_name)
        excluded_stage1 = excluded_stage1 + stage1_counts(type_name);
    end
    if stage2_counts.isKey(type_name)
        excluded_stage2 = excluded_stage2 + stage2_counts(type_name);
    end
    if stage3_counts.isKey(type_name)
        excluded_stage3 = excluded_stage3 + stage3_counts(type_name);
    end
end

fprintf('  %-28s | %4d\n', '1. 초기 데이터', excluded_stage1);
fprintf('  %-28s | %4d\n', '2. 매칭 후', excluded_stage2);
fprintf('  %-28s | %4d\n\n', '3. 결측값 제거 후', excluded_stage3);

%% ========================================================================
%                           단계 7: 최종 요약
% =========================================================================
fprintf('【단계 7】 최종 요약\n');
fprintf('════════════════════════════════════════════════════════════\n');

fprintf('전체 데이터 흐름:\n');
fprintf('  초기 데이터        : %3d명\n', total_before);
fprintf('  → 매칭 실패        : -%2d명\n', total_loss);
fprintf('  → 매칭 성공        : %3d명\n', total_after);
fprintf('  → 결측값 제거      : -%2d명\n', total_after - final_count);
fprintf('  → 최종 분석 데이터 : %3d명\n\n', final_count);

fprintf('고성과자/저성과자 비율:\n');
fprintf('  고성과자 : %2d명 (%.1f%%)\n', high_stage3, high_stage3/final_count*100);
fprintf('  저성과자 : %2d명 (%.1f%%)\n', low_stage3, low_stage3/final_count*100);
fprintf('  불균형 비율: %.2f:1\n\n', high_stage3/low_stage3);

fprintf('✅ 단계별 손실 추적 완료\n\n');
