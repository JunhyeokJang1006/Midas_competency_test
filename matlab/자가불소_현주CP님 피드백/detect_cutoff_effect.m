function [has_cutoff, cutoff_info] = detect_cutoff_effect(data, theoretical_range, varargin)
%DETECT_CUTOFF_EFFECT 통계적으로 정교한 Cut-off 효과 탐지
%
% 입력:
%   data - 분석할 데이터 벡터
%   theoretical_range - [min, max] 이론적 범위
%   varargin - 선택적 매개변수
%     'alpha' - 유의수준 (기본값: 0.05)
%     'bins' - 히스토그램 구간 수 (기본값: 10)
%     'verbose' - 상세 출력 여부 (기본값: false)
%
% 출력:
%   has_cutoff - Cut-off 효과 존재 여부 (논리값)
%   cutoff_info - 상세 분석 결과 구조체

% 입력 매개변수 파싱
p = inputParser;
addRequired(p, 'data', @isnumeric);
addRequired(p, 'theoretical_range', @(x) isnumeric(x) && length(x) == 2);
addParameter(p, 'alpha', 0.05, @(x) isnumeric(x) && x > 0 && x < 1);
addParameter(p, 'bins', 10, @(x) isnumeric(x) && x > 0);
addParameter(p, 'verbose', false, @islogical);
parse(p, data, theoretical_range, varargin{:});

alpha = p.Results.alpha;
n_bins = p.Results.bins;
verbose = p.Results.verbose;

% 결과 구조체 초기화
cutoff_info = struct();
cutoff_info.sample_size = length(data);
cutoff_info.theoretical_min = theoretical_range(1);
cutoff_info.theoretical_max = theoretical_range(2);
cutoff_info.observed_min = min(data);
cutoff_info.observed_max = max(data);
cutoff_info.theoretical_range_span = theoretical_range(2) - theoretical_range(1);

% 1. 기본 분포 통계
cutoff_info.mean = mean(data);
cutoff_info.std = std(data);
cutoff_info.skewness = skewness(data);
cutoff_info.kurtosis = kurtosis(data);

% 2. 정규성 검정 (Shapiro-Wilk 또는 Kolmogorov-Smirnov)
if length(data) <= 5000
    try
        % Shapiro-Wilk 검정 (작은 샘플에 적합)
        [h_sw, p_sw] = swtest(data, alpha);
        cutoff_info.normality_test = 'Shapiro-Wilk';
        cutoff_info.normality_h = h_sw;
        cutoff_info.normality_p = p_sw;
    catch
        % Shapiro-Wilk가 없으면 Kolmogorov-Smirnov 사용
        [h_ks, p_ks] = kstest(zscore(data));
        cutoff_info.normality_test = 'Kolmogorov-Smirnov';
        cutoff_info.normality_h = h_ks;
        cutoff_info.normality_p = p_ks;
    end
else
    % 큰 샘플에는 Kolmogorov-Smirnov 사용
    [h_ks, p_ks] = kstest(zscore(data));
    cutoff_info.normality_test = 'Kolmogorov-Smirnov';
    cutoff_info.normality_h = h_ks;
    cutoff_info.normality_p = p_ks;
end

% 3. 구간별 빈도 분석
bin_edges = linspace(theoretical_range(1), theoretical_range(2), n_bins + 1);
bin_counts = histcounts(data, bin_edges);
bin_frequencies = bin_counts / length(data);
cutoff_info.bin_frequencies = bin_frequencies;
cutoff_info.bin_edges = bin_edges;

% 이론적 균등분포와 비교 (카이제곱 검정)
expected_freq = length(data) / n_bins;
expected_counts = repmat(expected_freq, 1, n_bins);
[h_chi2, p_chi2] = chi2test(bin_counts, expected_counts);
cutoff_info.uniformity_h = h_chi2;
cutoff_info.uniformity_p = p_chi2;

% 4. Cut-off 효과 유형별 검정

% 4.1 Lower Truncation (하위 절단)
lower_20pct_bins = 1:max(1, floor(n_bins * 0.2));
lower_freq_sum = sum(bin_frequencies(lower_20pct_bins));
cutoff_info.lower_truncation = lower_freq_sum < 0.05; % 하위 20% 구간에 5% 미만
cutoff_info.lower_freq_sum = lower_freq_sum;

% 4.2 Upper Truncation (상위 절단)
upper_20pct_bins = ceil(n_bins * 0.8):n_bins;
upper_freq_sum = sum(bin_frequencies(upper_20pct_bins));
cutoff_info.upper_truncation = upper_freq_sum < 0.05; % 상위 20% 구간에 5% 미만
cutoff_info.upper_freq_sum = upper_freq_sum;

% 4.3 Floor Effect (바닥 효과)
floor_threshold = cutoff_info.observed_min + 5;
floor_data_count = sum(data <= floor_threshold);
floor_proportion = floor_data_count / length(data);
cutoff_info.floor_effect = floor_proportion > 0.15; % 최소값+5점 이내에 15% 이상
cutoff_info.floor_proportion = floor_proportion;
cutoff_info.floor_threshold = floor_threshold;

% 4.4 Ceiling Effect (천장 효과)
ceiling_threshold = cutoff_info.observed_max - 5;
ceiling_data_count = sum(data >= ceiling_threshold);
ceiling_proportion = ceiling_data_count / length(data);
cutoff_info.ceiling_effect = ceiling_proportion > 0.15; % 최대값-5점 이내에 15% 이상
cutoff_info.ceiling_proportion = ceiling_proportion;
cutoff_info.ceiling_threshold = ceiling_threshold;

% 4.5 강한 왜도를 통한 Cut-off 검정
cutoff_info.strong_left_skew = cutoff_info.skewness < -1.0; % 좌편향 (상위 절단 시사)
cutoff_info.strong_right_skew = cutoff_info.skewness > 1.0; % 우편향 (하위 절단 시사)

% 4.6 범위 축소 검정
observed_range = cutoff_info.observed_max - cutoff_info.observed_min;
range_utilization = observed_range / cutoff_info.theoretical_range_span;
cutoff_info.range_utilization = range_utilization;
cutoff_info.limited_range = range_utilization < 0.6; % 이론적 범위의 60% 미만 사용

% 5. 전체 Cut-off 효과 종합 판단
cutoff_effects = [
    cutoff_info.lower_truncation,
    cutoff_info.upper_truncation,
    cutoff_info.floor_effect,
    cutoff_info.ceiling_effect,
    cutoff_info.strong_left_skew,
    cutoff_info.strong_right_skew,
    cutoff_info.limited_range
];

has_cutoff = any(cutoff_effects);
cutoff_info.has_cutoff = has_cutoff;

% 6. Cut-off 유형 분류
cutoff_types = {};
if cutoff_info.lower_truncation || cutoff_info.strong_right_skew
    cutoff_types{end+1} = 'Lower_Truncation';
end
if cutoff_info.upper_truncation || cutoff_info.strong_left_skew
    cutoff_types{end+1} = 'Upper_Truncation';
end
if cutoff_info.floor_effect
    cutoff_types{end+1} = 'Floor_Effect';
end
if cutoff_info.ceiling_effect
    cutoff_types{end+1} = 'Ceiling_Effect';
end
if cutoff_info.limited_range
    cutoff_types{end+1} = 'Limited_Range';
end

if isempty(cutoff_types)
    cutoff_types{1} = 'None';
end

cutoff_info.cutoff_types = cutoff_types;
cutoff_info.cutoff_type_string = strjoin(cutoff_types, ', ');

% 7. Cut-off 강도 계산 (0-1 척도)
severity_scores = [];
if cutoff_info.lower_truncation
    severity_scores(end+1) = 1 - cutoff_info.lower_freq_sum / 0.2; % 예상 20%에서 얼마나 부족한지
end
if cutoff_info.upper_truncation
    severity_scores(end+1) = 1 - cutoff_info.upper_freq_sum / 0.2;
end
if cutoff_info.floor_effect
    severity_scores(end+1) = min(1, cutoff_info.floor_proportion / 0.15 - 1);
end
if cutoff_info.ceiling_effect
    severity_scores(end+1) = min(1, cutoff_info.ceiling_proportion / 0.15 - 1);
end
if cutoff_info.strong_left_skew
    severity_scores(end+1) = min(1, abs(cutoff_info.skewness) / 3);
end
if cutoff_info.strong_right_skew
    severity_scores(end+1) = min(1, abs(cutoff_info.skewness) / 3);
end
if cutoff_info.limited_range
    severity_scores(end+1) = 1 - cutoff_info.range_utilization;
end

if isempty(severity_scores)
    cutoff_info.severity_score = 0;
else
    cutoff_info.severity_score = mean(severity_scores);
end

% 8. 상세 출력 (옵션)
if verbose
    fprintf('\n=== Cut-off 효과 분석 결과 ===\n');
    fprintf('샘플 크기: %d\n', cutoff_info.sample_size);
    fprintf('이론적 범위: [%.1f, %.1f]\n', cutoff_info.theoretical_min, cutoff_info.theoretical_max);
    fprintf('관찰된 범위: [%.1f, %.1f]\n', cutoff_info.observed_min, cutoff_info.observed_max);
    fprintf('범위 활용도: %.1f%%\n', cutoff_info.range_utilization * 100);
    fprintf('왜도: %.3f, 첨도: %.3f\n', cutoff_info.skewness, cutoff_info.kurtosis);
    fprintf('정규성 검정 (%s): p=%.4f\n', cutoff_info.normality_test, cutoff_info.normality_p);
    fprintf('Cut-off 효과: %s\n', cutoff_info.has_cutoff);
    if cutoff_info.has_cutoff
        fprintf('Cut-off 유형: %s\n', cutoff_info.cutoff_type_string);
        fprintf('심각도 점수: %.3f\n', cutoff_info.severity_score);
    end
    fprintf('================================\n');
end

end

% 보조 함수: 카이제곱 검정
function [h, p] = chi2test(observed, expected)
    % 간단한 카이제곱 적합도 검정
    chi2_stat = sum((observed - expected).^2 ./ expected);
    df = length(observed) - 1;
    p = 1 - chi2cdf(chi2_stat, df);
    h = p < 0.05;
end

% 보조 함수: Shapiro-Wilk 검정 (MATLAB Statistics Toolbox 없을 경우 대체)
function [h, p] = swtest(x, alpha)
    % 간단한 Shapiro-Wilk 검정 구현
    % 실제로는 Statistics Toolbox의 swtest가 더 정확함
    try
        % Statistics Toolbox가 있으면 사용
        [h, p] = swtest(x, alpha);
    catch
        % 없으면 Kolmogorov-Smirnov로 대체
        [h, p] = kstest(zscore(x));
    end
end