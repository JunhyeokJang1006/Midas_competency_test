function [matchedData, matchedUpperScores, matchedIdx] = matchDataByID(periodData, upperCategoryResults)
% ID를 기반으로 Period 데이터와 상위항목 데이터를 매칭하는 함수
%
% 입력:
%   periodData - Period별 역량진단 데이터 (테이블)
%   upperCategoryResults - 상위항목 분석 결과 구조체
%
% 출력:
%   matchedData - 매칭된 Period 데이터
%   matchedUpperScores - 매칭된 상위항목 점수 매트릭스
%   matchedIdx - 매칭된 인덱스

    matchedData = [];
    matchedUpperScores = [];
    matchedIdx = [];

    try
        % Period 데이터에서 ID 컬럼 찾기
        periodIDCol = findIDColumn(periodData);
        if isempty(periodIDCol)
            fprintf('  ✗ Period 데이터에서 ID 컬럼을 찾을 수 없습니다\n');
            return;
        end

        % Period 데이터의 ID 추출 및 표준화
        periodIDs = extractAndStandardizeIDs(periodData{:, periodIDCol});

        % 상위항목 결과에서 ID 가져오기
        if ~isfield(upperCategoryResults, 'matchedIDs') || isempty(upperCategoryResults.matchedIDs)
            fprintf('  ✗ 상위항목 결과에서 ID를 찾을 수 없습니다\n');
            return;
        end

        upperIDs = upperCategoryResults.matchedIDs;
        upperScoreMatrix = upperCategoryResults.upperScoreMatrix;

        % ID 매칭
        [commonIDs, periodIdx, upperIdx] = intersect(periodIDs, upperIDs);

        if length(commonIDs) < 5
            fprintf('  ✗ 매칭된 ID가 부족합니다 (%d개)\n', length(commonIDs));
            return;
        end

        % 매칭된 데이터 추출
        matchedData = periodData(periodIdx, :);
        matchedUpperScores = upperScoreMatrix(upperIdx, :);
        matchedIdx = periodIdx;

        fprintf('  ✓ ID 매칭 성공: %d명 (전체 Period: %d명, 상위항목: %d명)\n', ...
                length(commonIDs), length(periodIDs), length(upperIDs));

    catch ME
        fprintf('  ✗ ID 매칭 중 오류 발생: %s\n', ME.message);
        matchedData = [];
        matchedUpperScores = [];
        matchedIdx = [];
    end
end