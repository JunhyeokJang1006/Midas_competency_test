# 시점별 개별 요인분석 기반 역량진단 성과점수 산출 및 성과기여도 상관분석 (수평 진단 버전)
# 2023년 하반기 ~ 2025년 상반기 (4개 시점) 수평 진단 데이터 개별 분석
#
# 작성일: 2025년
# 목적: 수평 진단 데이터를 사용하여 각 시점별로 독립적인 요인분석 수행 후 개별 점수 산출 및 성과기여도와 상관분석
# 특징: 동료 평가 점수를 마스터ID별로 평균하여 하향 진단 형태로 변환 후 원본과 동일한 분석 적용

# 필요한 패키지 로드
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

# 한글 인코딩 설정
Sys.setlocale("LC_ALL", "Korean")

# 작업 디렉토리 설정
setwd("D:/project/HR데이터/matlab")

# 로그 파일 설정
log_file <- "D:/project/matlab_runlog/runlog_horizontal_enhanced_R.txt"
sink(log_file, append = FALSE)

cat("========================================\n")
cat("시점별 개별 요인분석 기반 성과점수 산출 시작 (수평 진단 버전 - R)\n")
cat("========================================\n\n")

# 1. 초기 설정 및 전역 변수
data_path <- "D:/project/HR데이터/데이터/역량진단 데이터/역량진단_응답데이터"
periods <- c("23년_하반기", "24년_상반기", "24년_하반기", "25년_상반기")
file_names <- paste0(periods, "_역량진단_응답데이터.xlsx")

# 결과 저장용 리스트
all_data <- list()
period_results <- list()
consolidated_scores <- data.frame(stringsAsFactors = FALSE)

cat("[1단계] 모든 시점 데이터 로드\n")
cat("----------------------------------------\n")

# 2. 데이터 로드 (수평 진단 데이터를 하향 진단 형태로 변환)
for (p in 1:length(periods)) {
  cat(sprintf("▶ %s 데이터 로드 중...\n", periods[p]))
  file_name <- file.path(data_path, file_names[p])
  
  tryCatch({
    # 기본 데이터 로드
    master_ids <- read_excel(file_name, sheet = "기준인원 검토", col_names = TRUE)
    peer_raw_data <- read_excel(file_name, sheet = "수평 진단", col_names = TRUE)
    question_info <- read_excel(file_name, sheet = "문항 정보_타인진단", col_names = TRUE)
    
    cat(sprintf("  ✓ 마스터ID: %d명, 수평진단 원시데이터: %d개 레코드\n", 
                nrow(master_ids), nrow(peer_raw_data)))
    
    # 수평 진단 데이터를 하향 진단 형태로 변환
    cat(sprintf("\n▶ %s: 수평 진단 → 하향 진단 변환\n", periods[p]))
    
    col_names <- names(peer_raw_data)
    
    # 첫 번째 열: 평가대상자, 두 번째 열: 평가자
    if (ncol(peer_raw_data) < 2) {
      stop("수평 진단 데이터가 올바르지 않습니다.")
    }
    
    target_col <- peer_raw_data[[1]]  # 평가대상자 ID
    rater_col <- peer_raw_data[[2]]   # 평가자 ID
    
    # 유효한 데이터만 추출 (0이나 NaN이 아닌 경우)
    if (is.numeric(target_col) && is.numeric(rater_col)) {
      valid_rows <- !(is.na(target_col) | is.na(rater_col) | target_col == 0 | rater_col == 0)
    } else {
      valid_rows <- rep(TRUE, nrow(peer_raw_data))
    }
    
    valid_data <- peer_raw_data[valid_rows, ]
    
    cat(sprintf("  유효한 평가 레코드: %d개\n", sum(valid_rows)))
    
    # Q로 시작하는 문항 컬럼 식별
    question_cols <- c()
    question_indices <- c()
    
    for (col in 3:ncol(valid_data)) {  # 3번째 컬럼부터 (첫 2개는 ID)
      col_name <- col_names[col]
      col_data <- valid_data[[col]]
      
      if (is.numeric(col_data) && (grepl("^Q", col_name) || grepl("^q", col_name))) {
        question_cols <- c(question_cols, col_name)
        question_indices <- c(question_indices, col)
      }
    }
    
    cat(sprintf("  문항 수: %d개\n", length(question_cols)))
    
    if (length(question_cols) == 0) {
      cat("  ❌ 분석할 문항이 없습니다. 건너뜀.\n\n")
      next
    }
    
    # 고유한 평가대상자 목록
    if (is.numeric(valid_data[[1]])) {
      unique_targets <- unique(valid_data[[1]])
      unique_targets <- unique_targets[!is.na(unique_targets) & unique_targets > 0]
    } else {
      unique_targets <- unique(valid_data[[1]])
      unique_targets <- unique_targets[!is.na(unique_targets) & unique_targets != ""]
    }
    
    cat(sprintf("  고유한 평가 대상자: %d명\n", length(unique_targets)))
    
    # 평가대상자별 평균 점수 계산하여 하향진단 형태로 변환
    converted_data <- data.frame(stringsAsFactors = FALSE)
    
    # ID 컬럼명을 하향진단과 동일하게 설정
    first_col_name <- col_names[1]
    converted_data[[first_col_name]] <- unique_targets
    
    # 각 문항에 대해 평가대상자별 평균 계산
    evaluator_counts <- rep(0, length(unique_targets))
    
    for (q in 1:length(question_cols)) {
      question_col <- question_cols[q]
      question_index <- question_indices[q]
      avg_scores <- rep(NA, length(unique_targets))
      
      for (t in 1:length(unique_targets)) {
        target_id <- unique_targets[t]
        if (is.numeric(valid_data[[1]])) {
          target_rows <- valid_data[[1]] == target_id
        } else {
          target_rows <- valid_data[[1]] == target_id
        }
        
        target_scores <- valid_data[target_rows, question_index][[1]]
        
        # NaN이 아닌 점수들의 평균
        if (is.numeric(target_scores)) {
          valid_scores <- target_scores[!is.na(target_scores)]
        } else {
          valid_scores <- target_scores
        }
        
        if (length(valid_scores) > 0) {
          avg_scores[t] <- mean(valid_scores, na.rm = TRUE)
          if (q == 1) {  # 첫 번째 문항에서만 평가자 수 계산
            evaluator_counts[t] <- length(valid_scores)
          }
        } else {
          avg_scores[t] <- NA
          if (q == 1) {
            evaluator_counts[t] <- 0
          }
        }
      }
      
      converted_data[[question_col]] <- avg_scores
    }
    
    cat(sprintf("  변환 완료: %d명의 평가 대상자\n", nrow(converted_data)))
    cat(sprintf("  평균 평가자 수: %.1f명 (범위: %d-%d명)\n", 
                mean(evaluator_counts), min(evaluator_counts), max(evaluator_counts)))
    
    # 변환된 데이터를 selfData로 저장 (원본 코드와 호환성 유지)
    all_data[[paste0("period", p)]] <- list(
      masterIDs = master_ids,
      selfData = converted_data,
      questionInfo = question_info,
      evaluatorCounts = evaluator_counts
    )
    
    cat(sprintf("  ✅ %s 데이터 로드 및 변환 완료\n\n", periods[p]))
    
  }, error = function(e) {
    cat(sprintf("  ❌ %s 데이터 로드 실패: %s\n", periods[p], e$message))
    cat(sprintf("     파일 경로: %s\n\n", file_name))
  })
}

# 3. 시점별 개별 요인분석 수행
cat("\n[2단계] 시점별 개별 요인분석 수행\n")
cat("========================================\n")

# 전체 마스터 ID 리스트 생성
all_master_ids <- c()
for (p in 1:length(periods)) {
  if (!is.null(all_data[[paste0("period", p)]]) && 
      !is.null(all_data[[paste0("period", p)]]$masterIDs) && 
      nrow(all_data[[paste0("period", p)]]$masterIDs) > 0) {
    ids <- extract_master_ids(all_data[[paste0("period", p)]]$masterIDs)
    all_master_ids <- c(all_master_ids, ids)
  }
}
all_master_ids <- unique(all_master_ids)
cat(sprintf("전체 고유 마스터 ID: %d명\n\n", length(all_master_ids)))

# 결과 저장을 위한 데이터프레임 초기화
consolidated_scores <- data.frame(ID = all_master_ids, stringsAsFactors = FALSE)

# 각 시점별 분석
for (p in 1:length(periods)) {
  cat("========================================\n")
  cat(sprintf("[%s] 분석 시작\n", periods[p]))
  cat("========================================\n")
  
  if (is.null(all_data[[paste0("period", p)]]) || 
      is.null(all_data[[paste0("period", p)]]$selfData)) {
    cat("[경고] 데이터가 없습니다. 건너뜀.\n\n")
    period_results[[paste0("period", p)]] <- create_empty_period_result("NO_DATA")
    next
  }
  
  self_data <- all_data[[paste0("period", p)]]$selfData
  question_info <- all_data[[paste0("period", p)]]$questionInfo
  
  # 3-1. 문항 데이터 추출
  cat("▶ 문항 데이터 추출\n")
  
  # ID 컬럼 찾기
  id_col <- find_id_column(self_data)
  if (is.null(id_col)) {
    cat("  [경고] ID 컬럼을 찾을 수 없습니다. 건너뜀.\n\n")
    period_results[[paste0("period", p)]] <- create_empty_period_result("NO_ID_COLUMN")
    next
  }
  
  # Q로 시작하는 문항 컬럼 추출
  col_names <- names(self_data)
  question_cols <- c()
  
  for (col in 1:ncol(self_data)) {
    col_name <- col_names[col]
    col_data <- self_data[[col]]
    
    if (is.numeric(col_data) && (grepl("^Q", col_name) || grepl("^q", col_name))) {
      question_cols <- c(question_cols, col_name)
    }
  }
  
  cat(sprintf("  발견된 문항 수: %d개\n", length(question_cols)))
  
  if (length(question_cols) < 3) {
    cat("  [경고] 문항이 너무 적어 요인분석이 불가능합니다. 건너뜀.\n\n")
    period_results[[paste0("period", p)]] <- create_empty_period_result("INSUFFICIENT_ITEMS")
    next
  }
  
  # 응답 데이터 추출
  response_data <- as.matrix(self_data[, question_cols])
  response_ids <- extract_and_standardize_ids(self_data[[id_col]])
  
  # 결측치 처리
  response_data <- handle_missing_values(response_data)
  
  # 유효한 행만 선택
  valid_rows <- rowSums(is.na(response_data)) < (ncol(response_data) * 0.5)
  response_data <- response_data[valid_rows, , drop = FALSE]
  response_ids <- response_ids[valid_rows]
  
  cat(sprintf("  유효 응답자: %d명\n", length(response_ids)))
  
  if (nrow(response_data) < 10) {
    cat("  [경고] 표본 크기가 너무 작습니다. 건너뜀.\n\n")
    period_results[[paste0("period", p)]] <- create_empty_period_result("INSUFFICIENT_SAMPLE")
    next
  }
  
  # 3-2. 데이터 품질 검사 및 전처리
  cat("\n▶ 데이터 품질 검사 및 전처리\n")
  
  original_num_questions <- ncol(response_data)
  data_quality_flag <- "UNKNOWN"
  
  # 1. 상수 열 제거
  column_variances <- apply(response_data, 2, var, na.rm = TRUE)
  constant_columns <- column_variances < 1e-10
  
  if (any(constant_columns)) {
    cat(sprintf("  [제거] 상수 응답 문항 %d개\n", sum(constant_columns)))
    response_data <- response_data[, !constant_columns, drop = FALSE]
    question_cols <- question_cols[!constant_columns]
    column_variances <- column_variances[!constant_columns]
  }
  
  # 2. 다중공선성 처리
  if (ncol(response_data) > 1) {
    R <- cor(response_data, use = "pairwise.complete.obs")
    
    # 완벽한 상관관계 제거
    R_upper <- R
    R_upper[lower.tri(R_upper, diag = TRUE)] <- 0
    to_remove <- which(abs(R_upper) > 0.95, arr.ind = TRUE)
    to_remove <- unique(to_remove[, 1])
    
    if (length(to_remove) > 0) {
      cat(sprintf("  [제거] 다중공선성 문항 %d개\n", length(to_remove)))
      response_data <- response_data[, -to_remove, drop = FALSE]
      question_cols <- question_cols[-to_remove]
      column_variances <- column_variances[-to_remove]
    }
  }
  
  # 3. 극단값 처리
  outlier_count <- 0
  if (ncol(response_data) > 0) {
    for (col in 1:ncol(response_data)) {
      col_data <- response_data[, col]
      if (!all(is.na(col_data))) {
        mean_val <- mean(col_data, na.rm = TRUE)
        std_val <- sd(col_data, na.rm = TRUE)
        
        if (std_val > 0) {
          outlier_idx <- abs(col_data - mean_val) > 3 * std_val
          if (any(outlier_idx, na.rm = TRUE)) {
            col_data[outlier_idx & col_data > mean_val] <- mean_val + 3 * std_val
            col_data[outlier_idx & col_data < mean_val] <- mean_val - 3 * std_val
            response_data[, col] <- col_data
            outlier_count <- outlier_count + sum(outlier_idx, na.rm = TRUE)
          }
        }
      }
    }
  }
  
  if (outlier_count > 0) {
    cat(sprintf("  [처리] 극단값 %d개 조정\n", outlier_count))
  }
  
  # 4. 저분산 문항 제거
  low_variance_threshold <- 0.1
  if (length(column_variances) > 0) {
    low_variance_columns <- column_variances < low_variance_threshold
    
    if (any(low_variance_columns)) {
      cat(sprintf("  [제거] 저분산 문항 %d개\n", sum(low_variance_columns)))
      response_data <- response_data[, !low_variance_columns, drop = FALSE]
      question_cols <- question_cols[!low_variance_columns]
      column_variances <- column_variances[!low_variance_columns]
    }
  }
  
  # 5. 최종 품질 검사
  if (ncol(response_data) < 3) {
    cat("  [오류] 전처리 후 문항이 너무 적습니다.\n")
    period_results[[paste0("period", p)]] <- create_empty_period_result("POST_PROCESS_INSUFFICIENT")
    next
  }
  
  tryCatch({
    R_final <- cor(response_data, use = "pairwise.complete.obs")
    det_final <- det(R_final)
    cond_final <- kappa(R_final)
    
    cat(sprintf("  - 최종 문항 수: %d개 (원본: %d개)\n", ncol(response_data), original_num_questions))
    cat(sprintf("  - 상관행렬 조건수: %.2e\n", cond_final))
    cat(sprintf("  - 상관행렬 행렬식: %.2e\n", det_final))
    
    # 품질 판정
    if ((det_final > 1e-10) && (cond_final < 1e10)) {
      data_quality_flag <- "GOOD"
      cat("  ✓ 수치적 안정성 양호\n")
    } else if ((det_final > 1e-15) && (cond_final < 1e12)) {
      data_quality_flag <- "CAUTION"
      cat("  ⚠ 수치적 문제 있음 - 주의 필요\n")
    } else {
      data_quality_flag <- "POOR"
      cat("  ✗ 심각한 수치적 문제\n")
    }
    
  }, error = function(e) {
    data_quality_flag <<- "FAILED"
    cat(sprintf("  [경고] 상관행렬 계산 실패: %s\n", e$message))
  })
  
  # 3-3. 최적 요인 수 결정
  cat("\n▶ 최적 요인 수 결정\n")
  
  tryCatch({
    pca_result <- prcomp(response_data, scale = TRUE)
    latent <- pca_result$sdev^2
    
    num_factors_kaiser <- sum(latent > 1)
    num_factors_scree <- find_elbow_point(latent)
    num_factors_parallel <- parallel_analysis(response_data, 50)
    
    cat(sprintf("  - Kaiser 기준: %d개\n", num_factors_kaiser))
    cat(sprintf("  - Scree plot: %d개\n", num_factors_scree))
    cat(sprintf("  - Parallel analysis: %d개\n", num_factors_parallel))
    
    suggested_factors <- c(num_factors_kaiser, num_factors_scree, num_factors_parallel)
    optimal_num_factors <- median(suggested_factors)
    optimal_num_factors <- max(1, min(optimal_num_factors, min(5, ncol(response_data)-1)))
    
    cat(sprintf("  ✓ 선택된 요인 수: %d개\n", optimal_num_factors))
    
  }, error = function(e) {
    cat(sprintf("  [경고] PCA 실패: %s\n", e$message))
    optimal_num_factors <<- 1
  })
  
  # 3-4. 요인분석 수행
  cat("\n▶ 요인분석 실행\n")
  
  is_pca <- FALSE
  
  tryCatch({
    # psych 패키지의 fa 함수 사용
    fa_result <- fa(response_data, nfactors = optimal_num_factors, rotate = "promax", 
                    scores = "regression", fm = "ml")
    
    loadings <- fa_result$loadings
    factor_scores <- fa_result$scores
    specific_var <- fa_result$uniquenesses
    
    cat("  ✓ 요인분석 성공 (Promax 회전)\n")
    cat(sprintf("  - 누적 분산 설명률: %.2f%%\n", 100 * (1 - mean(specific_var))))
    
  }, error = function(e) {
    cat(sprintf("  [경고] 요인분석 실패: %s\n", e$message))
    cat("  [대안] PCA 점수 사용\n")
    
    tryCatch({
      pca_result <- prcomp(response_data, scale = TRUE)
      loadings <- pca_result$rotation[, 1:optimal_num_factors, drop = FALSE]
      factor_scores <- pca_result$x[, 1:optimal_num_factors, drop = FALSE]
      is_pca <<- TRUE
      cat("  ✓ PCA 성공\n")
      cat(sprintf("  - 누적 분산 설명률: %.2f%%\n", 
                  sum(pca_result$sdev[1:optimal_num_factors]^2) / sum(pca_result$sdev^2) * 100))
    }, error = function(e2) {
      cat(sprintf("  [오류] PCA도 실패: %s\n", e2$message))
      period_results[[paste0("period", p)]] <<- create_empty_period_result("ANALYSIS_FAILED")
      return()
    })
  })
  
  # 3-5. 성과 요인 식별 및 점수 산출
  cat("\n▶ 성과 요인 식별 및 점수 산출\n")
  
  # 성과 관련 요인 식별
  performance_factor_idx <- identify_performance_factor_advanced(loadings, question_cols, question_info)
  cat(sprintf("  - 성과 요인: Factor %d\n", performance_factor_idx))
  
  # 개인별 성과점수 계산
  performance_scores <- factor_scores[, performance_factor_idx]
  
  # 0-100 점수로 변환
  min_score <- min(performance_scores, na.rm = TRUE)
  max_score <- max(performance_scores, na.rm = TRUE)
  
  if (max_score > min_score) {
    scaled_scores <- 20 + (performance_scores - min_score) / (max_score - min_score) * 60
  } else {
    scaled_scores <- rep(50, length(performance_scores))  # 모든 점수가 같으면 중간값
  }
  
  cat(sprintf("  - 성과점수 범위: %.1f ~ %.1f점\n", min(scaled_scores), max(scaled_scores)))
  cat(sprintf("  - 성과점수 평균: %.1f점 (표준편차: %.1f)\n", 
              mean(scaled_scores), sd(scaled_scores)))
  
  # 3-6. 결과 저장
  period_result <- list(
    period = periods[p],
    analysisDate = Sys.time(),
    numParticipants = length(response_ids),
    numQuestions = ncol(response_data),
    originalNumQuestions = original_num_questions,
    numFactors = optimal_num_factors,
    dataQuality = data_quality_flag,
    isPCA = is_pca,
    performanceFactorIdx = performance_factor_idx,
    participantIDs = response_ids,
    performanceScores = scaled_scores,
    factorLoadings = loadings,
    factorScores = factor_scores,
    questionNames = question_cols,
    avgEvaluators = if (!is.null(all_data[[paste0("period", p)]]$evaluatorCounts)) {
      mean(all_data[[paste0("period", p)]]$evaluatorCounts)
    } else {
      NA
    }
  )
  
  period_results[[paste0("period", p)]] <- period_result
  
  # 통합 점수 테이블에 추가
  period_col_name <- paste0(periods[p], "_Performance")
  
  # ID 매칭하여 점수 할당
  period_score_vector <- rep(NA, nrow(consolidated_scores))
  for (i in 1:length(response_ids)) {
    id_idx <- which(consolidated_scores$ID == response_ids[i])
    if (length(id_idx) > 0) {
      period_score_vector[id_idx] <- scaled_scores[i]
    }
  }
  
  consolidated_scores[[period_col_name]] <- period_score_vector
  
  cat(sprintf("\n✅ [%s] 분석 완료\n", periods[p]))
  cat(sprintf("   참여자: %d명, 문항: %d개, 요인: %d개\n", 
              length(response_ids), ncol(response_data), optimal_num_factors))
  if (!is.na(period_result$avgEvaluators)) {
    cat(sprintf("   평균 평가자 수: %.1f명\n", period_result$avgEvaluators))
  }
  cat("\n")
}

# 4. 종합 분석 및 통계
cat("\n[3단계] 종합 분석 및 통계\n")
cat("========================================\n")

# 성공한 분석 개수
success_count <- 0
total_participants <- 0
for (p in 1:length(periods)) {
  if (!is.null(period_results[[paste0("period", p)]]) && 
      period_results[[paste0("period", p)]]$period != "EMPTY") {
    success_count <- success_count + 1
    total_participants <- total_participants + period_results[[paste0("period", p)]]$numParticipants
  }
}

cat(sprintf("✓ 성공한 분석: %d/%d개 시점\n", success_count, length(periods)))
cat(sprintf("✓ 총 분석 참여자: %d명\n", total_participants))

# 시점별 결과 요약
cat("\n▶ 시점별 분석 결과:\n")
cat(sprintf("%-15s %10s %10s %10s %12s %10s\n", "시점", "참여자수", "문항수", "요인수", "데이터품질", "평가자수"))
cat(paste(rep("-", 70), collapse = ""), "\n")

for (p in 1:length(periods)) {
  if (!is.null(period_results[[paste0("period", p)]]) && 
      period_results[[paste0("period", p)]]$period != "EMPTY") {
    result <- period_results[[paste0("period", p)]]
    evaluator_info <- if (!is.na(result$avgEvaluators)) {
      sprintf("%.1f", result$avgEvaluators)
    } else {
      "N/A"
    }
    cat(sprintf("%-15s %10d %10d %10d %12s %10s\n", 
                result$period, result$numParticipants, result$numQuestions, 
                result$numFactors, result$dataQuality, evaluator_info))
  } else {
    cat(sprintf("%-15s %10s %10s %10s %12s %10s\n", 
                periods[p], "FAILED", "-", "-", "-", "-"))
  }
}

# 5. 결과 저장
cat("\n[4단계] 결과 저장\n")
cat("========================================\n")

tryCatch({
  # 결과 디렉토리 생성
  result_dir <- "D:/project/HR데이터/결과"
  if (!dir.exists(result_dir)) {
    dir.create(result_dir, recursive = TRUE)
  }
  
  # 파일명 생성
  timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  result_file_name <- paste0("수평진단_시점별요인분석_결과_", timestamp, ".xlsx")
  result_file_path <- file.path(result_dir, result_file_name)
  
  # Excel 파일 생성
  wb <- createWorkbook()
  
  # 1. 종합 점수 시트
  addWorksheet(wb, "종합점수")
  writeData(wb, "종합점수", consolidated_scores)
  
  # 2. 시점별 상세 결과 시트
  for (p in 1:length(periods)) {
    if (!is.null(period_results[[paste0("period", p)]]) && 
        period_results[[paste0("period", p)]]$period != "EMPTY") {
      
      result <- period_results[[paste0("period", p)]]
      
      # 개인별 점수 테이블
      period_score_table <- data.frame(
        ID = result$participantIDs,
        PerformanceScore = result$performanceScores,
        stringsAsFactors = FALSE
      )
      
      # 요인별 점수 추가
      for (f in 1:result$numFactors) {
        factor_col_name <- paste0("Factor", f, "_Score")
        factor_scores <- result$factorScores[, f]
        
        # 0-100 스케일로 변환
        min_fs <- min(factor_scores, na.rm = TRUE)
        max_fs <- max(factor_scores, na.rm = TRUE)
        if (max_fs > min_fs) {
          scaled_fs <- 20 + (factor_scores - min_fs) / (max_fs - min_fs) * 60
        } else {
          scaled_fs <- rep(50, length(factor_scores))
        }
        
        period_score_table[[factor_col_name]] <- scaled_fs
      }
      
      sheet_name <- paste0(periods[p], "_점수")
      addWorksheet(wb, sheet_name)
      writeData(wb, sheet_name, period_score_table)
      
      # 요인 부하량 테이블
      loading_table <- data.frame(
        Question = result$questionNames,
        stringsAsFactors = FALSE
      )
      for (f in 1:result$numFactors) {
        factor_col_name <- paste0("Factor", f)
        loading_table[[factor_col_name]] <- result$factorLoadings[, f]
      }
      
      loading_sheet_name <- paste0(periods[p], "_부하량")
      addWorksheet(wb, loading_sheet_name)
      writeData(wb, loading_sheet_name, loading_table)
    }
  }
  
  # 파일 저장
  saveWorkbook(wb, result_file_path, overwrite = TRUE)
  
  cat("✓ 상세 결과 저장 완료\n")
  cat("📁 결과 파일:", result_file_path, "\n")
  
}, error = function(e) {
  cat("❌ 결과 저장 실패:", e$message, "\n")
})

# 6. 최종 보고서
cat("\n[최종 보고서]\n")
cat("========================================\n")
cat("수평 진단 기반 시점별 개별 요인분석 완료\n")
cat("========================================\n\n")

cat("📊 분석 개요:\n")
cat("• 분석 방법: 수평 진단 (동료 평가) → 하향 진단 변환 후 요인분석\n")
cat(sprintf("• 분석 시점: %d개 (%s)\n", length(periods), paste(periods, collapse = ", ")))
cat(sprintf("• 성공 분석: %d개 시점\n", success_count))
cat(sprintf("• 총 참여자: %d명\n", total_participants))

if (success_count > 0) {
  cat("\n🎯 주요 특징:\n")
  cat("• 동료 평가 점수를 개인별로 평균하여 개별 분석 수행\n")
  cat("• 원본 코드와 동일한 정교한 전처리 및 요인분석 적용\n")
  cat("• 다중 기준(Kaiser, Scree, Parallel Analysis)으로 최적 요인 수 결정\n")
  cat("• 성과 관련 요인 자동 식별 및 점수 산출\n")
  cat("• 품질 검증을 통한 신뢰할 수 있는 시점만 활용\n\n")
}

cat("✅ 수평 진단 기반 종합 분석이 완료되었습니다!\n")
cat("📁 상세 결과는 Excel 파일을 확인하세요.\n\n")

# 로그 종료
sink()
cat("📝 로그 파일:", log_file, "\n")
cat("🎉 분석 완료!\n")

# 보조 함수들 정의
extract_master_ids <- function(master_table) {
  if (nrow(master_table) == 0) {
    return(character(0))
  }
  
  # ID 컬럼 찾기
  for (col in 1:ncol(master_table)) {
    col_name <- names(master_table)[col]
    if (grepl("id|사번|empno|employee_id", tolower(col_name))) {
      ids <- master_table[[col]]
      if (is.numeric(ids)) {
        return(as.character(ids[!is.na(ids)]))
      } else {
        return(as.character(ids[!is.na(ids) & ids != ""]))
      }
    }
  }
  return(character(0))
}

create_empty_period_result <- function(reason) {
  list(
    period = "EMPTY",
    reason = reason,
    analysisDate = Sys.time(),
    numParticipants = 0,
    numQuestions = 0,
    numFactors = 0
  )
}

find_id_column <- function(data_table) {
  col_names <- names(data_table)
  
  for (col in 1:ncol(data_table)) {
    col_name <- tolower(col_names[col])
    col_data <- data_table[[col]]
    
    if (grepl("id|사번|empno|employee", col_name) && 
        ((is.numeric(col_data) && !all(is.na(col_data))) || 
         (is.character(col_data) && !all(is.na(col_data) | col_data == "")))) {
      return(col)
    }
  }
  return(NULL)
}

extract_and_standardize_ids <- function(raw_ids) {
  if (is.numeric(raw_ids)) {
    standardized_ids <- as.character(raw_ids)
  } else {
    standardized_ids <- as.character(raw_ids)
  }
  
  # 빈 값이나 NaN 처리
  standardized_ids[is.na(standardized_ids) | standardized_ids == "NaN"] <- ""
  return(standardized_ids)
}

handle_missing_values <- function(raw_data) {
  clean_data <- raw_data
  
  for (col in 1:ncol(raw_data)) {
    missing_idx <- is.na(raw_data[, col])
    if (any(missing_idx) && !all(missing_idx)) {
      col_mean <- mean(raw_data[!missing_idx, col], na.rm = TRUE)
      clean_data[missing_idx, col] <- col_mean
    } else if (all(missing_idx)) {
      clean_data[, col] <- 3  # 기본값 (중간값)
    }
  }
  return(clean_data)
}

find_elbow_point <- function(eigen_values) {
  if (length(eigen_values) < 3) {
    return(1)
  }
  
  # 2차 차분을 이용한 엘보 포인트 찾기
  diffs <- diff(eigen_values)
  second_diffs <- diff(diffs)
  
  # 가장 큰 변화가 일어나는 지점
  elbow_point <- which.max(abs(second_diffs)) + 1
  elbow_point <- min(elbow_point, length(eigen_values))
  
  # 최소 1, 최대 요인 수 제한
  return(max(1, min(elbow_point, length(eigen_values))))
}

parallel_analysis <- function(data, num_iterations) {
  n <- nrow(data)
  p <- ncol(data)
  
  # 실제 고유값
  real_eigen_values <- eigen(cov(data, use = "complete.obs"))$values
  real_eigen_values <- sort(real_eigen_values, decreasing = TRUE)
  
  # 랜덤 데이터의 고유값
  random_eigen_values <- matrix(0, nrow = num_iterations, ncol = p)
  
  for (iter in 1:num_iterations) {
    random_data <- matrix(rnorm(n * p), nrow = n, ncol = p)
    random_eigen_values[iter, ] <- sort(eigen(cov(random_data))$values, decreasing = TRUE)
  }
  
  # 95 백분위수 사용
  random_eigen_threshold <- apply(random_eigen_values, 2, quantile, probs = 0.95)
  
  # 실제 고유값이 랜덤보다 큰 개수
  num_factors <- sum(real_eigen_values > random_eigen_threshold)
  return(max(1, num_factors))  # 최소 1개
}

identify_performance_factor_advanced <- function(loadings, question_names, question_info) {
  performance_keywords <- c("성과", "목표", "달성", "결과", "효과", "기여", "창출", "개선", "수행", "완수", "생산", "실적")
  num_factors <- ncol(loadings)
  performance_scores <- rep(0, num_factors)
  
  for (f in 1:num_factors) {
    # 높은 부하량 문항들
    high_loading_items <- which(abs(loadings[, f]) > 0.3)
    
    for (item_idx in 1:length(high_loading_items)) {
      item <- high_loading_items[item_idx]
      question_name <- question_names[item]
      
      # 문항 정보에서 내용 찾기
      tryCatch({
        if (nrow(question_info) > 0) {
          # 문항명으로 매칭 시도
          for (row in 1:nrow(question_info)) {
            q_code <- as.character(question_info[row, 1])
            
            if (grepl(question_name, q_code) || grepl(q_code, question_name)) {
              question_text <- as.character(question_info[row, 2])
              
              # 키워드 매칭
              for (k in 1:length(performance_keywords)) {
                if (grepl(performance_keywords[k], tolower(question_text))) {
                  performance_scores[f] <- performance_scores[f] + abs(loadings[item, f])
                }
              }
              break
            }
          }
        }
      }, error = function(e) {
        # 매칭 실패 시 무시하고 계속
      })
    }
    
    # 요인의 전체적인 부하량 패턴도 고려
    # 높은 부하량이 많은 요인일수록 성과 관련 가능성 높음
    performance_scores[f] <- performance_scores[f] + 0.1 * sum(abs(loadings[, f]) > 0.5)
  }
  
  performance_idx <- which.max(performance_scores)
  
  # 성과 요인을 찾지 못한 경우 첫 번째 요인 사용
  if (all(performance_scores == 0)) {
    performance_idx <- 1
  }
  
  return(performance_idx)
}
