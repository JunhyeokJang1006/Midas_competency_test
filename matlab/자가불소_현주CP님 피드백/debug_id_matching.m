% ID 매칭 디버깅 스크립트
clc; clear;

% 파일 경로
hr_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
comp_file = 'D:\project\HR데이터\데이터\역량검사 요청 정보\23-25년 역량검사.xlsx';

% 데이터 로드
fprintf('=== ID 매칭 디버깅 ===\n\n');
fprintf('1. 데이터 로딩\n');
hr_data = readtable(hr_file, 'VariableNamingRule', 'preserve');
comp_data = readtable(comp_file, 'Sheet', '역량검사_상위항목', 'VariableNamingRule', 'preserve');

fprintf('  HR 데이터: %d행\n', height(hr_data));
fprintf('  역량검사 데이터: %d행\n\n', height(comp_data));

% 컬럼명 확인
fprintf('2. HR 데이터 컬럼 (%d개)\n', width(hr_data));
hr_cols = hr_data.Properties.VariableNames;
for i = 1:min(10, length(hr_cols))
    fprintf('  %2d. %s\n', i, hr_cols{i});
end
if length(hr_cols) > 10
    fprintf('  ... 외 %d개\n', length(hr_cols) - 10);
end
fprintf('\n');

fprintf('3. 역량검사 데이터 컬럼 (%d개)\n', width(comp_data));
comp_cols = comp_data.Properties.VariableNames;
for i = 1:min(10, length(comp_cols))
    fprintf('  %2d. %s\n', i, comp_cols{i});
end
if length(comp_cols) > 10
    fprintf('  ... 외 %d개\n', length(comp_cols) - 10);
end
fprintf('\n');

% ID 컬럼 찾기
fprintf('4. ID 컬럼 찾기\n');
hr_id_col = find(contains(lower(hr_cols), 'id'), 1);
comp_id_col = find(contains(lower(comp_cols), 'id'), 1);

if isempty(hr_id_col)
    fprintf('  ⚠ HR 데이터에 ID 컬럼 없음!\n');
    return;
end

if isempty(comp_id_col)
    fprintf('  ⚠ 역량검사 데이터에 ID 컬럼 없음!\n');
    return;
end

fprintf('  HR ID 컬럼: [%d] %s\n', hr_id_col, hr_cols{hr_id_col});
fprintf('  역량검사 ID 컬럼: [%d] %s\n\n', comp_id_col, comp_cols{comp_id_col});

% ID 데이터 샘플 확인
fprintf('5. ID 데이터 샘플 확인\n');
hr_ids_raw = hr_data.(hr_cols{hr_id_col});
comp_ids_raw = comp_data.(comp_cols{comp_id_col});

fprintf('  HR ID 타입: %s\n', class(hr_ids_raw));
fprintf('  역량검사 ID 타입: %s\n\n', class(comp_ids_raw));

fprintf('  HR ID 샘플 (첫 5개):\n');
for i = 1:min(5, length(hr_ids_raw))
    if isnumeric(hr_ids_raw)
        fprintf('    [%d] %.0f\n', i, hr_ids_raw(i));
    else
        fprintf('    [%d] %s\n', i, hr_ids_raw{i});
    end
end
fprintf('\n');

fprintf('  역량검사 ID 샘플 (첫 5개):\n');
for i = 1:min(5, length(comp_ids_raw))
    if isnumeric(comp_ids_raw)
        fprintf('    [%d] %.0f\n', i, comp_ids_raw(i));
    else
        fprintf('    [%d] %s\n', i, comp_ids_raw{i});
    end
end
fprintf('\n');

% 문자열 변환
fprintf('6. 문자열 변환 시도\n');
if isnumeric(hr_ids_raw)
    hr_ids = arrayfun(@(x) sprintf('%.0f', x), hr_ids_raw, 'UniformOutput', false);
    fprintf('  HR ID 변환 완료: %.0f → "%s"\n', hr_ids_raw(1), hr_ids{1});
else
    hr_ids = hr_ids_raw;
    fprintf('  HR ID 이미 문자열: "%s"\n', hr_ids{1});
end

if isnumeric(comp_ids_raw)
    comp_ids = arrayfun(@(x) sprintf('%.0f', x), comp_ids_raw, 'UniformOutput', false);
    fprintf('  역량검사 ID 변환 완료: %.0f → "%s"\n', comp_ids_raw(1), comp_ids{1});
else
    comp_ids = comp_ids_raw;
    fprintf('  역량검사 ID 이미 문자열: "%s"\n', comp_ids{1});
end
fprintf('\n');

% 교집합 확인
fprintf('7. 교집합 계산\n');
[matched_ids, hr_idx, comp_idx] = intersect(hr_ids, comp_ids);
fprintf('  매칭 성공: %d개\n\n', length(matched_ids));

if length(matched_ids) == 0
    fprintf('  ⚠ 매칭 실패! 원인 분석:\n\n');

    % HR ID 고유값
    hr_unique = unique(hr_ids);
    fprintf('    HR 고유 ID 수: %d개\n', length(hr_unique));
    fprintf('    HR ID 샘플 (첫 10개):\n');
    for i = 1:min(10, length(hr_unique))
        fprintf('      "%s"\n', hr_unique{i});
    end
    fprintf('\n');

    % 역량검사 ID 고유값
    comp_unique = unique(comp_ids);
    fprintf('    역량검사 고유 ID 수: %d개\n', length(comp_unique));
    fprintf('    역량검사 ID 샘플 (첫 10개):\n');
    for i = 1:min(10, length(comp_unique));
        fprintf('      "%s"\n', comp_unique{i});
    end
    fprintf('\n');

    % NaN 체크
    hr_nan_count = sum(cellfun(@isempty, hr_ids) | strcmp(hr_ids, 'NaN'));
    comp_nan_count = sum(cellfun(@isempty, comp_ids) | strcmp(comp_ids, 'NaN'));
    fprintf('    HR ID에 NaN: %d개\n', hr_nan_count);
    fprintf('    역량검사 ID에 NaN: %d개\n\n', comp_nan_count);

    % 공백 문자 체크
    fprintf('    특수 문자/공백 체크:\n');
    fprintf('      HR ID 첫번째: "%s" (길이: %d)\n', hr_unique{1}, length(hr_unique{1}));
    fprintf('      역량검사 ID 첫번째: "%s" (길이: %d)\n', comp_unique{1}, length(comp_unique{1}));
else
    fprintf('  ✓ 매칭 성공!\n');
    fprintf('    매칭된 ID 샘플 (첫 5개):\n');
    for i = 1:min(5, length(matched_ids))
        fprintf('      %s\n', matched_ids{i});
    end
end

fprintf('\n✅ 디버깅 완료\n');
