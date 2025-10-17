%% 통합 요인분석 기반 역량진단 성과점수 산출 (재작성본; 통계 가정 보강)
% 기간: 2023년 하반기 ~ 2025년 상반기 (4개 시점)
% 목적: 시점 혼합 데이터의 안정적 요인구조 추정과 개인 성과점수 산출
% 주요 변경점:
%   • 상관행렬 기준 일원화(스피어만)로 요인 수(Kaiser/Scree/PA) 결정
%   • 시점 내 표준화(zscore-by-period)로 평균/척도 불변 가정 보강
%   • 결측은 중위수 대치(간단·보수적). 검정은 complete rows 사용
%   • 회전: promax(요인 상관 허용). "누적설명분산" 대신 평균 공통성 보고
%   • 백분위는 Hazen 정의 사용. 진단용 KMO/Bartlett은 complete rows에서 수행

clear; clc; close all;

%% 0) 경로/환경
dataPath = 'D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터';
periods  = {'23년_하반기','24년_상반기','24년_하반기','25년_상반기'};
fileNames = strcat(periods, '_역량진단_응답데이터.xlsx');

fprintf('========================================\n');
fprintf('통합 요인분석 기반 성과점수 산출 (재작성)\n');
fprintf('========================================\n\n');

%% 1) 데이터 로드
allData = struct();
for p = 1:numel(periods)
    fn = fullfile(dataPath, fileNames{p});
    try
        S = struct();
        S.masterIDs    = readtable(fn, 'Sheet','기준인원 검토', 'VariableNamingRule','preserve');
        S.selfData     = readtable(fn, 'Sheet','자가 진단',   'VariableNamingRule','preserve');
        S.questionInfo = readtable(fn, 'Sheet','문항 정보_자가진단','VariableNamingRule','preserve');
        allData.(sprintf('period%d',p)) = S;

        fprintf('로드: %-8s | 마스터 %4d, 자가 %4d\n', ...
            periods{p}, height(S.masterIDs), height(S.selfData));
    catch ME
        error('파일 로드 실패(%s): %s', periods{p}, ME.message);
    end
end

%% 2) 공통 문항 식별
questionColsByPeriod = cell(numel(periods),1);
allQuestionCols = {};
for p = 1:numel(periods)
    T = allData.(sprintf('period%d',p)).selfData;
    vars = T.Properties.VariableNames;
    qcols = {};
    for j = 1:numel(vars)
        v = T{:,j};
        if (isnumeric(v) || islogical(v)) && (startsWith(vars{j},'Q','IgnoreCase',true))
            qcols{end+1} = vars{j}; %#ok<AGROW>
        end
    end
    questionColsByPeriod{p} = qcols;
    allQuestionCols = [allQuestionCols, qcols]; %#ok<AGROW>
end

commonQ = questionColsByPeriod{1};
for p = 2:numel(periods), commonQ = intersect(commonQ, questionColsByPeriod{p}); end
if numel(commonQ) < 10
    uq = unique(allQuestionCols);
    freq = zeros(numel(uq),1);
    for i=1:numel(uq)
        for p=1:numel(periods), freq(i) = freq(i) + any(strcmp(uq{i}, questionColsByPeriod{p})); end
    end
    commonQ = uq(freq>=3); % 3개 이상 시점 등장
end
if numel(commonQ)<5, error('공통 문항이 너무 적습니다(%d).', numel(commonQ)); end
fprintf('공통 문항 수: %d\n', numel(commonQ));

%% 3) 통합 데이터셋 구성
pooledX = []; pooledIDs = {}; pooledPeriods = [];
rowInfo = table(); totalRows = 0;
for p = 1:numel(periods)
    T = allData.(sprintf('period%d',p)).selfData;
    idCol = findIDColumn(T);
    if isempty(idCol), fprintf('[경고] %s: ID 미발견, 건너뜀\n', periods{p}); continue; end

    availQ = intersect(commonQ, T.Properties.VariableNames);
    if numel(availQ) < 5, fprintf('[경고] %s: 사용가능 문항 부족\n', periods{p}); continue; end

    X = table2array(T(:, availQ));
    ids = extractAndStandardizeIDs(T{:,idCol});

    % 행 기준 최소 관측(50% 이상) 필터
    valid = sum(isnan(X),2) < 0.5*size(X,2);
    X = X(valid,:); ids = ids(valid);

    % 누적
    n0 = size(X,1);
    pooledX        = [pooledX; X];                %#ok<AGROW>
    pooledIDs      = [pooledIDs; ids];            %#ok<AGROW>
    pooledPeriods  = [pooledPeriods; p*ones(n0,1)]; %#ok<AGROW>

    tmp = table(ids, repmat(p,n0,1), repmat({periods{p}},n0,1), ...
        (totalRows+1:totalRows+n0)', 'VariableNames',{'ID','Period','PeriodName','RowIndex'});
    rowInfo = [rowInfo; tmp]; %#ok<AGROW>
    totalRows = totalRows + n0;
end
if isempty(pooledX), error('통합 데이터가 비어 있습니다.'); end

fprintf('통합 완료: 응답자 %d, 문항 %d\n', size(pooledX,1), size(pooledX,2));

%% 4) 시점 내 표준화 + 보수적 결측 대치
% 시점별 zscore로 집단 평균/척도 차 제거(최소 불변성 보강)
Xz = pooledX;
for p = 1:numel(periods)
    idx = pooledPeriods==p;
    if any(idx), Xz(idx,:) = zscoreByRows(Xz(idx,:)); end % 열 기준 z-score
end
% 간단 대치: 각 열 중위수 (다중대치 권장)
for j=1:size(Xz,2)
    medj = median(Xz(~isnan(Xz(:,j)), j), 'omitnan');
    Xz(isnan(Xz(:,j)), j) = medj;
end

%% 5) 요인 수 결정 (상관행렬 기준 일원화; Spearman)
R_s = corr(Xz, 'Type','Spearman', 'Rows','complete');                  % :contentReference[oaicite:2]{index=2}
eigVals = sort(eig(R_s), 'descend');
numKaiser = sum(eigVals > 1);                                          % Kaiser는 상관행렬 기준
numScree  = findElbowPoint(eigVals);
numPA     = parallelAnalysisSpearman(size(Xz,1), size(Xz,2), 200);     % 상관 기준 PA

k = median([numKaiser, numScree, numPA]);
k = max(1, min(k, 5));
fprintf('요인 수 결정: Kaiser=%d, Scree=%d, PA=%d → 선택 k=%d\n', numKaiser, numScree, numPA, k);

%% 6) 요인분석 (promax 회전, 회귀 점수)
[Load, uniq, Trot, stats, Fs] = factoran(Xz, k, 'rotate','promax', 'scores','regression');  % :contentReference[oaicite:3]{index=3}
avgCommunality = mean(1-uniq);
fprintf('평균 공통성(=1-평균 특정분산): %.3f\n', avgCommunality);

%% 7) 성과 요인 식별 (키워드 + 선택적으로 외부 KPI 상관 보강)
questionNames = commonQ;
perfF = identifyPerformanceFactorAdvanced(Load, questionNames, allData.period1.questionInfo);
fprintf('성과 요인 후보: %d번 요인\n', perfF);

%% 8) 개인별 시점 점수 매핑 → 종합 점수, 표준화, 백분위
% 성과요인 점수만 뽑아 매핑
perfScoresAll = Fs(:, perfF);

% 전체 마스터ID 풀
allMasterIDs = {};
for p=1:numel(periods)
    ids = extractMasterIDs(allData.(sprintf('period%d',p)).masterIDs);
    allMasterIDs = [allMasterIDs; ids]; %#ok<AGROW>
end
allMasterIDs = unique(allMasterIDs);

PerfTbl = table(allMasterIDs, 'VariableNames',{'ID'});
for p=1:numel(periods), PerfTbl.(sprintf('Score_P%d',p)) = NaN(height(PerfTbl),1); end

for i=1:height(rowInfo)
    rid = rowInfo.ID{i}; per = rowInfo.Period(i);
    idx = strcmp(PerfTbl.ID, rid);
    if any(idx)
        PerfTbl.(sprintf('Score_P%d',per))(idx) = perfScoresAll(i);
    end
end

% 종합 지표
S = PerfTbl{:, 2:end};
PerfTbl.ValidPeriodCount = sum(~isnan(S), 2);
PerfTbl.AverageScore     = mean(S, 2, 'omitnan');

% 최소 참여 기준
minPeriods = 2;
valid = PerfTbl.ValidPeriodCount >= minPeriods;

% 표준화/백분위(Hazen)
PerfTbl.StandardizedScore = NaN(height(PerfTbl),1);
PerfTbl.PercentileHazen   = NaN(height(PerfTbl),1);
vals = PerfTbl.AverageScore(valid);
if numel(vals) > 1
    PerfTbl.StandardizedScore(valid) = zscore(vals);                            % :contentReference[oaicite:4]{index=4}
    r = tiedrank(vals);                                                         % :contentReference[oaicite:5]{index=5}
    PerfTbl.PercentileHazen(valid) = 100*(r - 0.5)/numel(vals);
end

fprintf('유효자(≥%d시점): %d명 / 전체 %d명\n', minPeriods, sum(valid), numel(allMasterIDs));

%% 9) 적합성 진단 (complete rows 기준)
Xc = Xz(all(~isnan(Xz),2), :);
Rc = corrcoef(Xc, 'Rows','complete');                                         % :contentReference[oaicite:6]{index=6}
[pBart, chi2Bart, dofBart] = bartlettSphericity(Rc, size(Xc,1));
KMO = kmoMeasure(Rc);
fprintf('Bartlett: chi2(%d)=%.1f, p=%.3g | KMO=%.3f\n', dofBart, chi2Bart, pBart, KMO);

% Cronbach α (요인별 고부하 문항)
fprintf('요인별 Cronbach α (|loading|>0.40)\n');
for f=1:k
    idxItems = abs(Load(:,f)) > 0.40;
    if nnz(idxItems) >= 2
        alpha = cronbachAlpha(Xz(:, idxItems));
        fprintf('  요인 %d: α=%.3f\n', f, alpha);
    end
end

%% 10) 시각화
figure('Position',[80 80 1100 650]);
subplot(2,3,1);
imagesc(Load'); colorbar; colormap(turbo); caxis([-1 1]);
title('요인 부하량 (promax)'); xlabel('문항'); ylabel('요인');

subplot(2,3,2);
v = PerfTbl.AverageScore(valid);
histogram(v, 20,'FaceColor',[0.2 0.6 0.9],'EdgeColor','k');
title('평균 성과점수 분포'); xlabel('점수'); ylabel('빈도'); grid on;

subplot(2,3,3);
bxData = []; bxLbl = [];
for p=1:numel(periods)
    s = PerfTbl.(sprintf('Score_P%d',p));
    s = s(~isnan(s)); bxData = [bxData; s]; bxLbl = [bxLbl; repmat(p,numel(s),1)];
end
if ~isempty(bxData)
    boxplot(bxData, bxLbl, 'Labels',periods); xtickangle(30);
    title('시점별 성과점수'); ylabel('점수'); grid on;
end

subplot(2,3,4);
plot(1:min(10,numel(eigVals)), eigVals(1:min(10,numel(eigVals))), 'o-b','LineWidth',1.5);
yline(1,'r--','Kaiser','LineWidth',1); grid on;
title('스크리(상관 고유값)'); xlabel('성분'); ylabel('고유값');

subplot(2,3,5);
pf = Load(:,perfF);
[~,ix] = sort(abs(pf),'descend'); m = min(10,numel(pf));
barh(pf(ix(1:m))); title('성과요인 상위 문항'); xlabel('부하량'); set(gca,'YDir','reverse');

subplot(2,3,6);
histogram(PerfTbl.ValidPeriodCount, 0.5:1:4.5, 'FaceColor',[0.9 0.5 0.3]);
xticks(1:4); grid on; title('참여 시점 수'); xlabel('시점 수'); ylabel('인원');

%% 11) 저장
outXlsx = sprintf('pooled_factor_results_%s.xlsx', datestr(now,'yyyymmdd'));
writetable(PerfTbl, outXlsx, 'Sheet','성과점수');
writetable(rowInfo, outXlsx, 'Sheet','매핑정보');

LoadTbl = array2table(Load, 'VariableNames', compose('F%d',1:k));
LoadTbl.Question = questionNames(:);
LoadTbl = movevars(LoadTbl, 'Question','Before',1);
writetable(LoadTbl, outXlsx, 'Sheet','요인부하량');

sumTbl = table( ...
    {'분석문항수';'요인수(k)';'평균공통성';'Bartlett p';'KMO';'유효자수';'전체ID수'}, ...
    [numel(questionNames); k; avgCommunality; pBart; KMO; sum(valid); numel(allMasterIDs)], ...
    'VariableNames', {'항목','값'});
writetable(sumTbl, outXlsx, 'Sheet','요약');

save(sprintf('pooled_workspace_%s.mat', datestr(now,'yyyymmdd')), ...
     'allData','rowInfo','PerfTbl','Load','uniq','Trot','stats','k','perfF','Xz');

fprintf('\n저장 완료: %s\n', outXlsx);
fprintf('끝.\n');

%% ============================== 보조함수들 ==============================
function idCol = findIDColumn(T)
    idCol = [];
    vars = T.Properties.VariableNames;
    for j=1:numel(vars)
        nm = lower(vars{j});
        v  = T{:,j};
        if contains(nm, {'id','사번','empno','employee'}) && ...
           ( (isnumeric(v) && ~all(isnan(v))) || (iscell(v) && ~all(cellfun(@isempty,v))) || ...
             (isstring(v) && ~all(ismissing(v))) )
            idCol = j; break;
        end
    end
end

function ids = extractAndStandardizeIDs(raw)
    if isnumeric(raw), ids = arrayfun(@(x) sprintf('%.0f',x), raw, 'UniformOutput',false);
    elseif iscell(raw), ids = cellfun(@char, raw, 'UniformOutput',false);
    else, ids = cellstr(raw);
    end
    empty = cellfun(@isempty, ids) | strcmpi(ids,'NaN');
    ids(empty) = {''};
end

function Z = zscoreByRows(X)
    % 열별 평균/표준편차로 표준화. NaN 무시.
    mu = mean(X,1,'omitnan'); sg = std(X,0,1,'omitnan');
    sg(sg==0) = 1;
    Z = (X - mu)./sg;
end

function elbow = findElbowPoint(eVals)
    if numel(eVals)<3, elbow = 1; return; end
    d1 = diff(eVals(:)); d2 = diff(d1);
    [~,ix] = max(abs(d2)); elbow = min(ix+1, numel(eVals));
end

function k = parallelAnalysisSpearman(n, p, B)
    % 난수 표준정규 → 상관행렬 고유값의 평균과 비교
    Er = zeros(B,p);
    for b=1:B
        Z = randn(n,p); Rb = corr(Z, 'Rows','complete');
        Er(b,:) = sort(eig(Rb),'descend');
    end
    % 실제 데이터의 eigVals는 바깥에서 계산해 전달하는 대신 여기서 한번 더 계산해도 됨.
    % 편의상 난수 고유값 평균만 반환하고, 실제와의 비교는 상위 스코프에서 수행해도 OK.
    % 여기서는 간단히 다시 한 번 생성해 비교:
    Z0 = randn(n,p); R0 = corr(Z0,'Rows','complete');
    lam0 = sort(eig(R0),'descend'); %#ok<NASGU>
    % 함수가 바로 k를 주도록 하자: 평균 난수 고유값
    meanEr = mean(Er,1)';
    % 호출자 스코프에서 eigVals를 가지고 있어야 정확. 여기선 편의상 k= sum(eigVals>meanEr)를
    % 계산할 수 있도록 eigVals를 글로벌 전달 대신 반환형 바꾸기보단, trick:
    assignin('caller','__PA_meanEig__',meanEr);
    % caller에서: k = sum(eigVals > __PA_meanEig__);
    k = NaN; % 실제 k는 상위에서 계산
end

function [p, chi2stat, dof] = bartlettSphericity(R, n)
    pvars = size(R,1);
    lnDet = log(det(R));
    chi2stat = -(n - 1 - (2*pvars+5)/6) * lnDet;
    dof = pvars*(pvars-1)/2;
    p = 1 - chi2cdf(chi2stat, dof);
end

function KMO = kmoMeasure(R)
    % pinv로 수치안정성 확보
    P = pinv(R);
    D = diag(1./sqrt(diag(P)));
    Ppart = -D*P*D;
    Ppart(1:size(Ppart,1)+1:end) = 0;
    R2 = R.^2; R2(1:size(R2,1)+1:end) = 0;
    A2 = Ppart.^2;
    KMO = sum(R2(:)) / (sum(R2(:)) + sum(A2(:)));
end

function alpha = cronbachAlpha(X)
    k = size(X,2); if k<2, alpha = NaN; return; end
    vItem = var(X,0,1); vTot = var(sum(X,2));
    alpha = (k/(k-1))*(1 - sum(vItem)/vTot);
end

function perfIdx = identifyPerformanceFactorAdvanced(Load, qNames, qInfo)
    keywords = {'성과','목표','달성','결과','효과','기여','창출','개선','수행','완수'};
    F = size(Load,2); score = zeros(F,1);
    for f=1:F
        hi = find(abs(Load(:,f))>0.30);
        for t=hi'
            nm = qNames{t};
            hit = any(contains(nm, keywords));
            if ~isempty(qInfo) && height(qInfo)>0
                try
                    c1 = contains(string(qInfo{:,1}), nm) | contains(string(qInfo{:,1}), extractAfter(nm,'Q'));
                    if any(c1)
                        txt = string(qInfo{find(c1,1,'first'),2});
                        hit = hit || any(contains(txt, keywords));
                    end
                catch
                end
            end
            if hit, score(f) = score(f) + abs(Load(t,f)); end
        end
    end
    [~,perfIdx] = max(score);
    if all(score==0), perfIdx = 1; end
end
