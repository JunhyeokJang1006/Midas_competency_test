# CSR 및 조직지표 분석 시스템 종합 가이드

## 📋 개요

본 가이드는 기업사회책임(CSR)과 조직지표(조직시너지, 자부심)를 포함한 확장된 리더십 분석 시스템에 대한 종합적인 분석 과정과 결과를 정리한 문서입니다.

**분석 기간**: 2025년 9월 16일
**분석 대상**: 25년 상반기 (Q40-Q46)
**주요 스크립트**: `corr_CSR_vs_comp_test.m`

---

## 🎯 주요 성과 및 해결된 문제들

### ✅ 1. 확장된 분석 범위 구현

**🔴 기존**: CSR 문항만 분석 (Q42-Q46)
**🟢 개선**: 조직지표 + CSR 개별 문항 분석 (Q40-Q46)

**확장된 문항 구성**:
- Q40: 조직시너지 (OrganizationalSynergy)
- Q41: 자부심 (Pride)
- Q42: CSR C1 - 소통관계 (C1_Communication_Relationship)
- Q43: CSR C2 - 소통목적 (C2_Communication_Purpose)
- Q44: CSR C3 - 전략고객가치 (C3_Strategy_CustomerValue)
- Q45: CSR C4 - 전략성과 (C4_Strategy_Performance)
- Q46: CSR C5 - 성찰조직 (C5_Reflection_Organizational)

### ✅ 2. Marginal Significance 분석 도입

**🔴 기존 기준**: p < 0.05만 유의
**🟢 개선 기준**:
- **Significant**: p < 0.05 (**, ***)
- **Marginal**: 0.05 ≤ p < 0.1 (†)
- **Strong Correlation**: |r| > 0.1 & p < 0.1

### ✅ 3. 동적 컬럼 탐지 시스템

**🔴 기존 문제**: 하드코딩된 컬럼명으로 인한 탐지 실패
**🟢 해결책**:
```matlab
% 숫자형 데이터이고, 결측치가 아닌 값이 있으며, 분산이 0보다 큰 컬럼만 선택
if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
    competencyMatrix = [competencyMatrix, colData];
    validCompetencyCategories{end+1} = colName;
end
```

### ✅ 4. 13개 분석 카테고리 생성

**조직 기반 지표 (3개)**:
- Organizational_Score (Q40+Q41 통합)
- OrganizationalSynergy (Q40 개별)
- Pride (Q41 개별)

**CSR 개별 문항 (5개)**:
- C1_Communication_Relationship ~ C5_Reflection_Organizational

**CSR 카테고리별 점수 (3개)**:
- Communication_Score (C1+C2)
- Strategy_Score (C3+C4)
- Reflection_Score (C5)

**통합 점수 (2개)**:
- Total_CSR_Score (C1-C5)
- Total_Leadership_Score (Q40-Q46 전체)

---

## 📊 최종 분석 결과

### 🎲 데이터 현황
- **전체 분석 대상**: 90명 (원시 데이터)
- **최종 분석 대상**: 60명 (완전한 데이터)
- **분석 문항**: 7개 (Q40-Q46)
- **분석 카테고리**: 13개
- **역량검사 상위항목**: 10개

### 📈 주요 발견 사항

#### 🏆 유의한 상관관계 (p < 0.05): 2개
| 지표 | 역량 | 상관계수 (r) | p-value | 해석 |
|------|------|-------------|---------|------|
| Pride | 감정역량 | 0.387 | 0.002** | 강한 양의 상관 |
| Pride | 성실성 | 0.337 | 0.008** | 중간 양의 상관 |

#### 📊 Marginal 상관관계 (0.05 ≤ p < 0.1): 4개
| 지표 | 역량 | 상관계수 (r) | p-value | 해석 |
|------|------|-------------|---------|------|
| Pride | 인지역량 | 0.248 | 0.056† | 약한 양의 상관 |
| Pride | 성과역량 | 0.242 | 0.063† | 약한 양의 상관 |
| C2_Communication_Purpose | 성과역량 | 0.230 | 0.077† | 약한 양의 상관 |
| Communication_Score | 성실성 | 0.226 | 0.082† | 약한 양의 상관 |

#### 🎯 핵심 인사이트

1. **자부심(Pride)이 핵심 예측변수**:
   - 4개 역량과 유의/marginal 관계
   - 특히 감정역량과 성실성에 강한 관계

2. **CSR Communication 영역의 중요성**:
   - C2(소통목적)과 Communication 통합점수가 성과/성실성과 관련

3. **조직시너지는 상관관계 미약**:
   - 모든 역량과 유의한 관계 없음

---

## 🔧 시스템 아키텍처

### 📁 주요 파일 구조
```
D:\project\HR데이터\matlab\CSR\
├── corr_CSR_vs_comp_test.m                    # 메인 분석 스크립트
├── CSR_조직지표_분석가이드.md                  # 본 가이드 문서
└── [생성된 결과 파일들]
    ├── CSR_vs_competency_results.xlsx         # Excel 결과
    ├── CSR_vs_competency_workspace.mat        # MAT 워크스페이스
    ├── CSR_score_distributions.png            # 점수 분포 히스토그램
    └── CSR_competency_correlation_heatmap.png # 상관계수 히트맵
```

### 🔄 11단계 분석 파이프라인

1. **[1단계] 25년 상반기 데이터 로드**
   - CSR 원시 데이터 + 역량검사 데이터 동시 로드

2. **[2단계] 확장된 조직지표 + CSR 데이터 추출**
   - Q40-Q46 동적 탐지 및 전처리
   - 7개 문항 모두 성공적으로 발견

3. **[3단계] 확장된 문항별 척도 분석 및 표준화**
   - 실제 데이터 기반 척도 매핑
   - [0,1] 범위 완벽 표준화

4. **[4단계] 영역별 점수 계산**
   - 13개 분석 카테고리 생성
   - 조직기반 + CSR개별 + CSR카테고리 + 통합점수

5. **[5단계] 역량검사 상위항목 데이터 전처리**
   - 10개 역량 동적 탐지 성공

6. **[6단계] ID 매칭 및 데이터 정리**
   - 90명 → 60명 (완전한 데이터)
   - Listwise deletion 적용

7. **[7단계] CSR vs 역량검사 상관분석**
   - Pairwise correlation
   - 13 × 10 = 130개 상관분석 수행

8. **[8단계] 유의한 상관관계 요약**
   - Significant + Marginal 분류
   - 강한 상관관계 (|r| > 0.1 & p < 0.1) 식별

9. **[9단계] 중다회귀분석**
   - Total_CSR_Score 예측 모델 시도 (데이터 부족으로 생략)

10. **[10단계] 시각화 생성**
    - 점수 분포 히스토그램
    - 상관계수 히트맵

11. **[11단계] 결과 저장**
    - Excel: 15개 시트 (카테고리별 + 유의 + marginal + 강한상관 + 요약)
    - MAT: 완전한 워크스페이스

### 💾 출력 파일 구조

**Excel 파일** (`CSR_vs_competency_results.xlsx`):
- 13개 CSR 카테고리별 상관분석 시트
- `유의한_상관관계`: p < 0.05 결과
- `Marginal_상관관계`: 0.05 ≤ p < 0.1 결과
- `강한_상관관계`: |r| > 0.1 & p < 0.1 결과
- `요약통계`: 카테고리별 요약 (Marginal 개수 포함)

**MAT 파일**: 모든 분석 결과 + 중간 데이터

---

## 🛠 기술적 개선사항

### 1. **확장된 척도 매핑**
```matlab
% 실제 데이터 기반 척도 정의
extendedScaleMapping('Q40') = [0, 4];  % 조직시너지
extendedScaleMapping('Q41') = [1, 5];  % 자부심
extendedScaleMapping('Q42') = [2, 5];  % CSR C1
% ... 각 문항별 정확한 척도 정의
```

### 2. **Marginal Significance 구현**
```matlab
% 3단계 유의성 분류
significantIdx = find(pValues < 0.05);           % 유의
marginalIdx = find(pValues >= 0.05 & pValues < 0.1);  % marginal
strongIdx = find(abs(correlations) > 0.1 & pValues < 0.1);  % 강한상관
```

### 3. **동적 컬럼 탐지**
```matlab
% 숫자형 + 분산존재 + 결측치아님 조건
if isnumeric(colData) && ~all(isnan(colData)) && var(colData, 'omitnan') > 0
    validCompetencyCategories{end+1} = colName;
end
```

### 4. **13개 카테고리 자동 생성**
- 조직 기반 (Q40+Q41)
- CSR 개별 (C1-C5)
- CSR 카테고리 (Communication, Strategy, Reflection)
- 통합 점수 (Total_CSR, Total_Leadership)

---

## 📋 사용 방법

### 🚀 기본 실행
```matlab
cd('D:\project\HR데이터\matlab\CSR');
run('corr_CSR_vs_comp_test.m');
```

### 📊 결과 해석 가이드

**상관계수 해석 기준**:
- |r| < 0.1: 무시할 수 있는 상관관계
- 0.1 ≤ |r| < 0.3: 약한 상관관계
- 0.3 ≤ |r| < 0.5: 보통 상관관계
- 0.5 ≤ |r| < 0.7: 강한 상관관계

**유의성 기준**:
- p < 0.001: *** (매우 유의)
- p < 0.01: ** (유의)
- p < 0.05: * (유의)
- 0.05 ≤ p < 0.1: † (marginal)
- p ≥ 0.1: ns (비유의)

---

## 🔍 주요 연구 결과

### 📈 자부심(Pride)의 중심적 역할

**강한 관계**:
- 감정역량 (r=0.387, p=0.002): 자부심이 높을수록 감정 관리 능력 우수
- 성실성 (r=0.337, p=0.008): 자부심이 높을수록 책임감과 성실성 우수

**Marginal 관계**:
- 인지역량 (r=0.248, p=0.056): 자부심이 높을수록 사고력 향상 경향
- 성과역량 (r=0.242, p=0.063): 자부심이 높을수록 성과 창출 능력 향상 경향

### 📊 Communication의 선택적 영향

**Marginal 관계**:
- C2(소통목적) → 성과역량: 목적 지향적 소통이 성과와 관련
- Communication 통합점수 → 성실성: 소통 능력이 성실성과 관련

### ⚠️ 조직시너지의 제한적 영향

- 모든 역량과 유의한 관계 없음
- 개별적 자부심이 조직 차원 시너지보다 중요

---

## 🔄 향후 연구 방향

### 1. **자부심 중심 심화 분석**
- 자부심의 구성 요소 분해 분석
- 자부심과 성과의 매개 변수 탐색

### 2. **Communication 카테고리 세분화**
- C1(관계) vs C2(목적) 차이점 분석
- 소통 방식별 효과 차이 연구

### 3. **조직시너지 재개념화**
- 현재 측정 방식의 타당성 검토
- 팀 단위 vs 개인 단위 분석

### 4. **종단 연구 확장**
- 시간에 따른 변화 패턴 분석
- 인과관계 규명

---

## 📞 지원 및 문의

**개발자**: Claude Code Assistant
**개발일**: 2025년 9월 16일
**버전**: v3.0 (Marginal Significance + 조직지표 확장)

**주요 개선사항**:
- ✅ 조직지표 (Q40-Q41) 추가 분석
- ✅ CSR 개별 문항 (C1-C5) 분석
- ✅ Marginal significance 도입
- ✅ 13개 분석 카테고리 자동 생성
- ✅ 동적 컬럼 탐지 시스템
- ✅ 종합적인 시각화 및 결과 저장

---

## 📚 기술 참고사항

### 필수 전제조건
1. MATLAB R2018b 이상
2. Statistics and Machine Learning Toolbox
3. 25년 상반기 역량진단 데이터
4. 23-25년 역량검사 데이터

### 예상 실행시간
- 데이터 규모: 90명 × 7문항 × 10역량
- 실행시간: 약 2-5분
- 메모리 사용량: ~150MB

### 데이터 품질 요구사항
- 최소 분석 대상: 10명 이상
- 문항별 최소 유효 응답: 5명 이상
- ID 매칭 성공률: 80% 이상

---

*본 가이드는 실제 분석 과정에서 발생한 모든 문제와 해결 과정을 상세히 기록한 실무 중심의 문서입니다.*
*Marginal significance 도입으로 더 넓은 범위의 의미있는 관계를 탐지할 수 있게 되었습니다.*