function [score, prediction, probability] = calculateTalentScore(competency_scores)
% 역량 점수를 입력받아 종합점수와 예측 결과를 반환하는 함수
% Input: competency_scores - 역량 점수 벡터 (1x9)
% Output: score - 종합점수, prediction - 예측결과, probability - 확률

    % 가중치 정보 (상위 10개 핵심 역량)
    weights = [21.8011, 17.6877, 15.3024, 12.8947, 11.3993, 9.6754, 6.5399, 2.7415, 1.9581];
    feature_names = {'전략성', '정체성', '성실성', '사회성', '긍정성', '적극성', '관계성', '비활성', '과활성'};
    threshold = 0.2733;

    % 입력 검증
    if length(competency_scores) ~= length(weights)
        error('입력 역량 점수 개수가 맞지 않습니다. %d개 필요', length(weights));
    end

    % 표준화 (Z-score)
    normalized_scores = (competency_scores - mean(competency_scores)) / std(competency_scores);

    % 가중 점수 계산
    score = sum(normalized_scores .* (weights/100));

    % 예측 및 확률 계산
    if score > threshold
        prediction = '고성과자 가능성 높음';
    else
        prediction = '일반 성과자';
    end
    
    % 시그모이드 함수로 확률 계산
    probability = 1 / (1 + exp(-(score - threshold)));
end

% 사용 예시:
% new_scores = [75, 82, 68, 90, 77, 85, 70, 88, 79, 81];  % 10개 역량 점수
% [score, pred, prob] = calculateTalentScore(new_scores);
% fprintf('종합점수: %.2f, 예측: %s, 확률: %.1f%%\n', score, pred, prob*100);
