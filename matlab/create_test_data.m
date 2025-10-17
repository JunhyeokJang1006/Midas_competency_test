%% ====================== 모의데이터 생성 (테스트용) ======================
% 이 블록을 '분석 스크립트' 맨 위에 붙여서 먼저 실행하세요.
% 생성 후, 바로 아래의 기존 분석 코드가 동일 경로/파일을 읽어서 동작합니다.

clear; clc; close all;

% 경로/파일명 정의 (기존 스크립트와 동일해야 함)
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods  = {'23년_하반기','24년_상반기','24년_하반기','25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

% 폴더 없으면 생성
if ~exist(dataPath, 'dir'); mkdir(dataPath); end

% ---------------- 파라미터 ----------------
nMasterAll   = 200;   % 전체 마스터 인원 풀
nPerPeriod   = 120;   % 각 시점 자가진단 응답자 수
nCommonItems = 15;    % 모든 시점 공통 문항 수 (Q01~Q15)
nExtraItems  = 5;     % 시점별 추가 문항 수 (교집합에 안 들어가도 OK)

% 마스터 ID 전체 풀 (문자형 '사번')
masterPool = compose("E%04d", 1:nMasterAll).';

% 공통 문항 코드/문항명
commonQs   = compose("Q%02d", 1:nCommonItems);
extraQsTpl = compose("QX%02d", 1:nExtraItems); % 시점별로 이름 바꿔서 고유로

% 문항 설명(간단)
mkItemText = @(q) "문항("+q+") : 업무 수행 및 목표 달성 관련 자기평가";

rng(42); % 재현성

for p = 1:numel(periods)
    % ---- 기준인원 검토(마스터) 시트 ----
    % 각 시점 기준인원은 전체 200명 중 160명 정도로 가정
    idxMaster = sort(randsample(nMasterAll, 160));
    masterIDs = masterPool(idxMaster);
    T_master = table(masterIDs, 'VariableNames', {'사번'});
    
    % ---- 자가 진단 시트 ----
    % 응답자는 기준인원 중 nPerPeriod명 추출
    idxResp   = sort(randsample(numel(masterIDs), nPerPeriod));
    selfIDs   = masterIDs(idxResp);
    
    % 공통문항 + 시점별 추가문항
    extraQs = erase(extraQsTpl, "X") + "_" + string(p); % QX01 -> Q01_1 처럼 변경
    % 실제 컬럼명은 전부 'Q'로 시작하게 맞춥니다 (교집합 탐색 로직과 호환)
    extraQs = replace(extraQs, "Q0", "Q9"); % 예: Q01_1 -> Q91_1 (문항코드 충돌 방지)
    
    itemNames = [commonQs, extraQs];
    nItems    = numel(itemNames);
    
    % 리커트(1~5) 모의 응답: 시점별 평균 차/분산 차, 약간의 결측 포함
    resp = nan(nPerPeriod, nItems);
    periodShift = [-0.3, 0, +0.2, +0.1];   % 시점별 수준 차 가정
    for j = 1:nItems
        base = 3 + 0.7*randn(nPerPeriod,1) + periodShift(p);
        lik  = round(min(max(base,1),5));   % 1~5로 절단
        % 5% 결측
        missIdx = rand(nPerPeriod,1) < 0.05;
        lik(missIdx) = NaN;
        resp(:,j) = lik;
    end
    
    % 자가진단 테이블 구성 (ID 컬럼명은 '사번'으로 제공)
    T_self = table(selfIDs, 'VariableNames', {'사번'});
    for j=1:nItems
        T_self.(itemNames{j}) = resp(:,j);
    end
    
    % ---- 문항 정보 시트 ----
    % 공통문항 + 추가문항의 코드/내용
    qcodes = string(itemNames(:));                         % n×1 string
qtext  = arrayfun(mkItemText, qcodes, 'UniformOutput', false); 
qtext  = qtext(:);                                     % n×1 cell

T_qinfo = table(qcodes, qtext, 'VariableNames', {'문항코드','문항내용'});
    
    % ---- 엑셀로 기록 ----
    fn = fullfile(dataPath, fileNames{p});
    if exist(fn, 'file'); delete(fn); end
    
    writetable(T_master, fn, 'Sheet','기준인원 검토');
    writetable(T_self,   fn, 'Sheet','자가 진단');
    writetable(T_qinfo,  fn, 'Sheet','문항 정보_자가진단');
    
    fprintf('모의파일 생성: %s | 마스터 %d명, 응답 %d명, 문항 %d개(공통 %d)\n', ...
        fileNames{p}, height(T_master), height(T_self), width(T_self)-1, nCommonItems);
end

fprintf('모의데이터 생성 완료. 이제 아래의 기존 분석 코드를 바로 실행하세요.\n');
%% ==================== 모의데이터 생성 끝 ====================
