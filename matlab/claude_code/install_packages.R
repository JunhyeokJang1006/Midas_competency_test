# R 패키지 설치 스크립트
# factor_analysis_by_period_peer_cc.R 실행에 필요한 패키지들을 설치합니다.

# 필요한 패키지 목록
required_packages <- c(
  "readxl",      # Excel 파일 읽기
  "dplyr",       # 데이터 처리
  "tidyr",       # 데이터 정리
  "psych",       # 요인분석, KMO, Cronbach's alpha
  "corpcor",     # 상관행렬 처리
  "openxlsx",    # Excel 파일 쓰기
  "ggplot2",     # 시각화
  "corrplot",    # 상관행렬 시각화
  "VIM",         # 결측값 처리
  "parallel"     # 병렬 처리
)

# 설치되지 않은 패키지 확인
missing_packages <- required_packages[!required_packages %in% installed.packages()[,"Package"]]

# 패키지 설치
if (length(missing_packages) > 0) {
  cat("다음 패키지들을 설치합니다:\n")
  cat(paste(missing_packages, collapse = ", "), "\n\n")
  
  install.packages(missing_packages, dependencies = TRUE)
  
  cat("패키지 설치가 완료되었습니다!\n")
} else {
  cat("모든 필요한 패키지가 이미 설치되어 있습니다.\n")
}

# 패키지 로드 테스트
cat("\n패키지 로드 테스트 중...\n")
test_result <- tryCatch({
  suppressPackageStartupMessages({
    library(readxl)
    library(dplyr)
    library(tidyr)
    library(psych)
    library(corpcor)
    library(openxlsx)
    library(ggplot2)
    library(corrplot)
    library(VIM)
    library(parallel)
  })
  cat("✅ 모든 패키지가 성공적으로 로드되었습니다!\n")
  cat("이제 factor_analysis_by_period_peer_cc.R 스크립트를 실행할 수 있습니다.\n")
}, error = function(e) {
  cat("❌ 패키지 로드 실패:", e$message, "\n")
  cat("패키지 설치를 다시 시도해주세요.\n")
})

# R 버전 확인
cat("\nR 버전 정보:\n")
cat(R.version.string, "\n")

# 권장 R 버전 (4.0 이상)
r_version <- as.numeric(paste(R.version$major, R.version$minor, sep = "."))
if (r_version < 4.0) {
  cat("⚠️  권장 R 버전: 4.0 이상 (현재:", r_version, ")\n")
  cat("일부 패키지에서 문제가 발생할 수 있습니다.\n")
} else {
  cat("✅ R 버전이 적절합니다.\n")
}
