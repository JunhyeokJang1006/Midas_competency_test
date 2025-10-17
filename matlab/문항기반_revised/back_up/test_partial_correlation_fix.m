%% Partial Correlation ID 매칭 테스트
clear; clc;

fprintf('========================================\n');
fprintf('Partial Correlation ID 매칭 테스트\n');
fprintf('========================================\n');

% 1. 나이 데이터 로드
agePath = 'D:\project\HR데이터\데이터\역량검사 요청 정보\최근 3년 입사자_인적정보_cleaned.xlsx';
ageData = readtable(agePath, 'Sheet', 1, 'VariableNamingRule', 'preserve');

fprintf('나이 데이터 로드 완료: %d명\n', height(ageData));
fprintf('ID 타입: %s\n', class(ageData.ID));
if ismember('만 나이', ageData.Properties.VariableNames)
    ageColStatus = '있음';
else
    ageColStatus = '없음';
end
fprintf('나이 컬럼: %s\n', ageColStatus);

% 2. 테스트용 matchedIDs 생성 (실제 데이터에서 일부 추출)
testIDs = ageData.ID(1:10);  % 처음 10개 ID 사용
fprintf('\n테스트 ID 샘플:\n');
for i = 1:min(5, length(testIDs))
    fprintf('  %d: %s (타입: %s)\n', i, string(testIDs(i)), class(testIDs(i)));
end

% 3. ID 매칭 테스트
fprintf('\nID 매칭 테스트 시작...\n');
matchedAgeValues = nan(length(testIDs), 1);

try
    for i = 1:length(testIDs)
        currentID = testIDs(i);

        % 안전한 ID 매칭
        if isnumeric(currentID)
            ageIdx = find(ageData.ID == currentID, 1);
        else
            % 문자열인 경우 숫자로 변환
            numericID = str2double(string(currentID));
            if ~isnan(numericID)
                ageIdx = find(ageData.ID == numericID, 1);
            else
                ageIdx = [];
            end
        end

        if ~isempty(ageIdx) && ageIdx <= height(ageData)
            matchedAgeValues(i) = ageData.("만 나이")(ageIdx);
            fprintf('  ID %s -> 나이 %d세 (인덱스 %d)\n', string(currentID), matchedAgeValues(i), ageIdx);
        else
            fprintf('  ID %s -> 매칭 실패\n', string(currentID));
        end
    end

    fprintf('\n✓ ID 매칭 완료\n');

    % 결과 확인
    validAgeIdx = ~isnan(matchedAgeValues);
    numValidAge = sum(validAgeIdx);

    fprintf('매칭 결과:\n');
    fprintf('  - 총 테스트 ID: %d개\n', length(testIDs));
    fprintf('  - 매칭 성공: %d개 (%.1f%%)\n', numValidAge, (numValidAge/length(testIDs))*100);

    if numValidAge > 0
        validAgeValues = matchedAgeValues(validAgeIdx);
        fprintf('  - 나이 범위: %.1f ~ %.1f세\n', min(validAgeValues), max(validAgeValues));
        fprintf('  - 평균 나이: %.1f세\n', mean(validAgeValues));
    end

    fprintf('\n✅ ID 매칭 테스트 성공!\n');

catch ME
    fprintf('❌ ID 매칭 중 오류: %s\n', ME.message);
    fprintf('스택 트레이스:\n');
    for i = 1:length(ME.stack)
        fprintf('  %s (줄 %d)\n', ME.stack(i).name, ME.stack(i).line);
    end
end

fprintf('========================================\n');