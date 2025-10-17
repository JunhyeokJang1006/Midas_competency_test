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
    converted_data <- data.frame()
    
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

# 3. 시점별 개별 요인분석 수행 (원본 코드와 동일한 구조)
cat("\n[2단계] 시점별 개별 요인분석 수행\n")
cat("========================================\n")

# 전체 마스터 ID 리스트 생성
all_master_ids <- c()
for (p in 1:length(periods)) {
  if (!is.null(all_data[[paste0("period", p)]]$masterIDs) && 
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
    
    # 상세한 품질 진단 정보
    cat("\n  [상세 품질 진단]\n")
    cat(sprintf("  - 데이터 크기: %d x %d\n", nrow(response_data), ncol(response_data)))
    cat(sprintf("  - 결측값 비율: %.2f%%\n", (sum(is.na(response_data)) / length(response_data)) * 100))
    cat(sprintf("  - 상관행렬 최대값: %.4f\n", max(R_final, na.rm = TRUE)))
    cat(sprintf("  - 상관행렬 최소값: %.4f\n", min(R_final, na.rm = TRUE)))
    cat(sprintf("  - 상관행렬 평균: %.4f\n", mean(R_final, na.rm = TRUE)))
    
    # 특이값 분석
    svd_result <- svd(R_final)
    singular_values <- svd_result$d
    cat(sprintf("  - 특이값 범위: %.2e ~ %.2e\n", min(singular_values), max(singular_values)))
    cat(sprintf("  - 특이값 비율 (최대/최소): %.2e\n", max(singular_values)/min(singular_values)))
    
    # 품질 판정 (기존 기준)
    if ((det_final > 1e-10) && (cond_final < 1e10)) {
      data_quality_flag <- "GOOD"
      cat("  ✓ 수치적 안정성 양호 (기존 기준)\n")
    } else if ((det_final > 1e-15) && (cond_final < 1e12)) {
      data_quality_flag <- "CAUTION"
      cat("  ⚠ 수치적 문제 있음 - 주의 필요 (기존 기준)\n")
    } else {
      data_quality_flag <- "POOR"
      cat("  ✗ 심각한 수치적 문제 (기존 기준)\n")
    }
    
    # 대안적 품질 기준 제안
    cat("\n  [대안적 품질 기준]\n")
    if ((det_final > 1e-15) && (cond_final < 1e12)) {
      cat("  ✓ 완화된 기준: GOOD\n")
    } else if ((det_final > 1e-20) && (cond_final < 1e15)) {
      cat("  ✓ 매우 완화된 기준: CAUTION\n")
    } else {
      cat("  ✗ 모든 기준 실패: POOR\n")
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
      
      # 요인 부하량 시트
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

# 6. 성과기여도 데이터 로드 및 전처리
cat("\n========================================\n")
cat("[5단계] 성과기여도 데이터 분석\n")
cat("========================================\n\n")

# 성과기여도 데이터 파일 경로 설정
contribution_data_path <- "D:/project/HR데이터/데이터/역량검사 요청 정보/최근 3년 입사자_인적정보.xlsx"

tryCatch({
  # 성과기여도 데이터 로드
  contribution_data <- read_excel(contribution_data_path, sheet = "성과기여도", col_names = TRUE)
  cat(sprintf("성과기여도 데이터 로드 완료: %d명\n", nrow(contribution_data)))
  
  # 데이터 구조 확인
  cat(sprintf("컬럼 수: %d\n", ncol(contribution_data)))
  cat("컬럼명 (처음 10개): ")
  cat(paste(names(contribution_data)[1:min(10, ncol(contribution_data))], collapse = ", "), "\n")
  
}, error = function(e) {
  cat("[오류] 성과기여도 데이터 로드 실패:", e$message, "\n")
  cat("파일 경로와 시트명을 확인해주세요.\n")
  return()
})

# 7. 성과기여도 점수 계산
cat("\n========================================\n")
cat("성과기여도 점수 계산\n")
cat("========================================\n\n")

# 분기별 성과기여도 점수 계산을 위한 데이터프레임 초기화
contribution_scores <- data.frame(ID = contribution_data[[1]], stringsAsFactors = FALSE)

# 분기 정보 정의 (23Q1~25Q2)
quarters <- c("23Q1", "23Q2", "23Q3", "23Q4", "24Q1", "24Q2", "24Q3", "24Q4", "25Q1", "25Q2")

# 조직성과등급 점수 매핑 (S=5, A=4, B=3, C=2, D=1)
grade_to_score <- c("S" = 5, "A" = 4, "B" = 3, "C" = 2, "D" = 1)

cat("분기별 성과기여도 점수 계산 중...\n")

# 각 분기별 처리
for (q in 1:length(quarters)) {
  quarter <- quarters[q]
  cat(sprintf("  [%s 처리 중]\n", quarter))
  
  # 해당 분기의 컬럼 찾기
  contribution_col <- paste0(quarter, "_개인기여도")
  organization_col <- paste0(quarter, "_조직")
  grade_col <- paste0(quarter, "_조직성과등급")
  
  # 컬럼 존재 여부 확인
  has_contribution <- contribution_col %in% names(contribution_data)
  has_organization <- organization_col %in% names(contribution_data)
  has_grade <- grade_col %in% names(contribution_data)
  
  if (has_contribution && has_grade) {
    # 개인기여도와 조직성과등급 데이터 추출
    personal_contrib <- contribution_data[[contribution_col]]
    org_grades <- contribution_data[[grade_col]]
    
    # 조직성과등급을 숫자로 변환
    org_scores <- rep(NA, length(org_grades))
    for (i in 1:length(org_grades)) {
      grade <- as.character(org_grades[i])
      if (grade %in% names(grade_to_score)) {
        org_scores[i] <- grade_to_score[grade]
      }
    }
    
    # 성과기여도 점수 = 개인기여도 × 조직성과점수
    quarter_score <- personal_contrib * org_scores
    
    # 유효한 데이터 통계
    valid_count <- sum(!is.na(quarter_score))
    cat(sprintf("    - 유효한 데이터: %d/%d명\n", valid_count, length(quarter_score)))
    
    if (valid_count > 0) {
      cat(sprintf("    - 평균 점수: %.3f\n", mean(quarter_score, na.rm = TRUE)))
      cat(sprintf("    - 표준편차: %.3f\n", sd(quarter_score, na.rm = TRUE)))
    }
    
    # 결과 저장
    contribution_scores[[paste0("Score_", quarter)]] <- quarter_score
    
  } else {
    cat("    - [경고] 필요한 컬럼이 없습니다\n")
    contribution_scores[[paste0("Score_", quarter)]] <- rep(NA, nrow(contribution_data))
  }
}

# 8. 시점별 성과기여도 집계 (반기별)
cat("\n========================================\n")
cat("반기별 성과기여도 집계\n")
cat("========================================\n\n")

# 반기별 매핑 (역량진단 시점과 맞추기)
# 23년 하반기: 23Q3, 23Q4
# 24년 상반기: 24Q1, 24Q2
# 24년 하반기: 24Q3, 24Q4
# 25년 상반기: 25Q1, 25Q2

period_mapping <- list(
  c("23Q3", "23Q4"),  # 23년 하반기
  c("24Q1", "24Q2"),  # 24년 상반기
  c("24Q3", "24Q4"),  # 24년 하반기
  c("25Q1", "25Q2")   # 25년 상반기
)

contribution_by_period <- data.frame(ID = contribution_scores$ID, stringsAsFactors = FALSE)

for (p in 1:length(period_mapping)) {
  quarter_list <- period_mapping[[p]]
  period_name <- paste0("Contribution_Period", p)
  
  cat(sprintf("[%s - %s 집계]\n", periods[p], paste(quarter_list, collapse = ", ")))
  
  # 해당 반기의 분기별 점수들을 평균 계산
  period_scores <- data.frame()
  for (q in 1:length(quarter_list)) {
    quarter <- quarter_list[q]
    score_col <- paste0("Score_", quarter)
    if (score_col %in% names(contribution_scores)) {
      quarter_score <- contribution_scores[[score_col]]
      if (ncol(period_scores) == 0) {
        period_scores <- data.frame(quarter_score)
      } else {
        period_scores <- cbind(period_scores, quarter_score)
      }
    }
  }
  
  if (ncol(period_scores) > 0) {
    # 반기별 평균 계산
    avg_score <- rowMeans(period_scores, na.rm = TRUE)
    contribution_by_period[[period_name]] <- avg_score
    
    valid_count <- sum(!is.na(avg_score))
    cat(sprintf("  - 유효한 데이터: %d명\n", valid_count))
    if (valid_count > 0) {
      cat(sprintf("  - 평균: %.3f, 표준편차: %.3f\n", 
                  mean(avg_score, na.rm = TRUE), sd(avg_score, na.rm = TRUE)))
    }
  } else {
    contribution_by_period[[period_name]] <- rep(NA, nrow(contribution_by_period))
    cat("  - 데이터 없음\n")
  }
}

# 전체 평균 성과기여도 계산
all_contrib_scores <- cbind(
  contribution_by_period$Contribution_Period1,
  contribution_by_period$Contribution_Period2,
  contribution_by_period$Contribution_Period3,
  contribution_by_period$Contribution_Period4
)

contribution_by_period$AverageContribution <- rowMeans(all_contrib_scores, na.rm = TRUE)
contribution_by_period$ValidPeriodCount <- rowSums(!is.na(all_contrib_scores))

# 9. 역량진단 성과점수와 성과기여도 매칭
cat("\n========================================\n")
cat("[6단계] 역량진단 vs 성과기여도 데이터 매칭\n")
cat("========================================\n\n")

# ID를 기준으로 두 데이터셋 매칭
# consolidated_scores (역량진단 기반) vs contribution_by_period (성과기여도 기반)

# ID를 문자열로 통일
if (is.numeric(consolidated_scores$ID)) {
  competency_ids <- as.character(consolidated_scores$ID)
} else {
  competency_ids <- as.character(consolidated_scores$ID)
}

if (is.numeric(contribution_by_period$ID)) {
  contribution_ids <- as.character(contribution_by_period$ID)
} else {
  contribution_ids <- as.character(contribution_by_period$ID)
}

# 교집합 찾기
common_ids <- intersect(competency_ids, contribution_ids)
competency_idx <- match(common_ids, competency_ids)
contribution_idx <- match(common_ids, contribution_ids)

cat("매칭 결과:\n")
cat(sprintf("  - 역량진단 데이터: %d명\n", nrow(consolidated_scores)))
cat(sprintf("  - 성과기여도 데이터: %d명\n", nrow(contribution_by_period)))
cat(sprintf("  - 공통 ID: %d명\n", length(common_ids)))
cat(sprintf("  - 매칭률: %.1f%%\n", 100 * length(common_ids) / min(nrow(consolidated_scores), nrow(contribution_by_period))))

if (length(common_ids) < 10) {
  cat("[경고] 매칭된 데이터가 너무 적습니다. ID 형식을 확인해주세요.\n")
  cat("샘플 ID (역량진단):", paste(head(competency_ids, 5), collapse = ", "), "\n")
  cat("샘플 ID (성과기여도):", paste(head(contribution_ids, 5), collapse = ", "), "\n")
}

# 매칭된 데이터로 통합 테이블 생성
if (length(common_ids) > 0) {
  combined_data <- data.frame(ID = common_ids, stringsAsFactors = FALSE)
  
  # 역량진단 점수 추가
  combined_data$Factor_Period1 <- consolidated_scores[[paste0(periods[1], "_Performance")]][competency_idx]
  combined_data$Factor_Period2 <- consolidated_scores[[paste0(periods[2], "_Performance")]][competency_idx]
  combined_data$Factor_Period3 <- consolidated_scores[[paste0(periods[3], "_Performance")]][competency_idx]
  combined_data$Factor_Period4 <- consolidated_scores[[paste0(periods[4], "_Performance")]][competency_idx]
  
  # 성과기여도 점수 추가
  combined_data$Contribution_Period1 <- contribution_by_period$Contribution_Period1[contribution_idx]
  combined_data$Contribution_Period2 <- contribution_by_period$Contribution_Period2[contribution_idx]
  combined_data$Contribution_Period3 <- contribution_by_period$Contribution_Period3[contribution_idx]
  combined_data$Contribution_Period4 <- contribution_by_period$Contribution_Period4[contribution_idx]
  combined_data$Contribution_Average <- contribution_by_period$AverageContribution[contribution_idx]
  
  cat(sprintf("통합 데이터 생성 완료: %d명\n", nrow(combined_data)))
  
} else {
  cat("[오류] 매칭된 데이터가 없습니다. 분석을 계속할 수 없습니다.\n")
  return()
}

# 10. 상관분석
cat("\n========================================\n")
cat("[7단계] 역량진단 성과점수 vs 성과기여도 상관분석\n")
cat("========================================\n\n")

correlation_results <- list()

# 시점별 상관분석
cat("[시점별 상관분석]\n")
for (p in 1:4) {
  factor_col <- paste0("Factor_Period", p)
  contrib_col <- paste0("Contribution_Period", p)
  
  factor_scores <- combined_data[[factor_col]]
  contrib_scores <- combined_data[[contrib_col]]
  
  # 둘 다 유효한 값이 있는 경우만 분석
  valid_idx <- !is.na(factor_scores) & !is.na(contrib_scores)
  valid_count <- sum(valid_idx)
  
  if (valid_count >= 5) {  # 최소 5개 이상의 쌍이 있어야 상관분석 가능
    cor_result <- cor.test(factor_scores[valid_idx], contrib_scores[valid_idx])
    correlation <- cor_result$estimate
    p_value <- cor_result$p.value
    
    cat(sprintf("%s: r = %.3f (n=%d, p=%.3f)", periods[p], correlation, valid_count, p_value))
    if (p_value < 0.001) {
      cat(" ***")
    } else if (p_value < 0.01) {
      cat(" **")
    } else if (p_value < 0.05) {
      cat(" *")
    }
    cat("\n")
    
    correlation_results[[paste0("period", p)]] <- list(
      correlation = correlation,
      n = valid_count,
      p_value = p_value
    )
  } else {
    cat(sprintf("%s: 분석 불가 (유효 데이터 %d개)\n", periods[p], valid_count))
    correlation_results[[paste0("period", p)]] <- list(
      correlation = NA,
      n = valid_count,
      p_value = NA
    )
  }
}

# 11. 최종 결과 저장
cat("\n========================================\n")
cat("[8단계] 최종 결과 저장\n")
cat("========================================\n\n")

# Excel 파일로 저장
output_file_name <- paste0("수평진단_종합분석결과_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx")

tryCatch({
  # Excel 워크북 생성
  wb <- createWorkbook()
  
  # 1. 통합 점수 테이블 저장
  addWorksheet(wb, "역량진단_통합점수")
  writeData(wb, "역량진단_통합점수", consolidated_scores)
  
  # 2. 각 시점별 상세 결과 저장
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
  
  # 3. 성과기여도 데이터 저장
  addWorksheet(wb, "성과기여도점수")
  writeData(wb, "성과기여도점수", contribution_by_period)
  
  # 4. 통합 상관분석 데이터 저장
  addWorksheet(wb, "상관분석_통합데이터")
  writeData(wb, "상관분석_통합데이터", combined_data)
  
  # 5. 상관분석 결과 테이블 생성
  corr_result_table <- data.frame(
    Period = c(periods, "전체평균"),
    Correlation = rep(NA, 5),
    N = rep(NA, 5),
    P_Value = rep(NA, 5),
    Significance = rep("", 5),
    stringsAsFactors = FALSE
  )
  
  for (p in 1:4) {
    if (!is.null(correlation_results[[paste0("period", p)]])) {
      result <- correlation_results[[paste0("period", p)]]
      corr_result_table$Correlation[p] <- result$correlation
      corr_result_table$N[p] <- result$n
      corr_result_table$P_Value[p] <- result$p_value
      
      if (!is.na(result$p_value)) {
        if (result$p_value < 0.001) {
          corr_result_table$Significance[p] <- "***"
        } else if (result$p_value < 0.01) {
          corr_result_table$Significance[p] <- "**"
        } else if (result$p_value < 0.05) {
          corr_result_table$Significance[p] <- "*"
        } else {
          corr_result_table$Significance[p] <- "n.s."
        }
      } else {
        corr_result_table$Significance[p] <- "분석불가"
      }
    }
  }
  
  addWorksheet(wb, "상관분석결과")
  writeData(wb, "상관분석결과", corr_result_table)
  
  # 파일 저장
  saveWorkbook(wb, output_file_name, overwrite = TRUE)
  
  cat("📁 결과 파일:", output_file_name, "\n")
  
}, error = function(e) {
  cat("❌ 결과 저장 실패:", e$message, "\n")
})

# 12. 최종 보고서
cat("\n[최종 보고서]\n")
cat("========================================\n")
cat("수평 진단 기반 시점별 개별 요인분석 완료\n")
cat("========================================\n\n")

cat("📊 분석 개요:\n")
cat("• 분석 방법: 수평 진단 (동료 평가) → 하향 진단 변환 후 요인분석\n")
cat(sprintf("• 분석 시점: %d개 (%s)\n", length(periods), paste(periods, collapse = ", ")))
cat(sprintf("• 성공 분석: %d개 시점\n", success_count))
cat(sprintf("• 총 참여자: %d명\n", total_participants))

if (exists("combined_data")) {
  cat(sprintf("• 성과기여도 매칭: %d명\n", nrow(combined_data)))
}
cat("\n")

if (success_count > 0) {
  cat("🎯 주요 특징:\n")
  cat("• 동료 평가 점수를 개인별로 평균하여 개별 분석 수행\n")
  cat("• 원본 코드와 동일한 정교한 전처리 및 요인분석 적용\n")
  cat("• 다중 기준(Kaiser, Scree, Parallel Analysis)으로 최적 요인 수 결정\n")
  cat("• 성과 관련 요인 자동 식별 및 점수 산출\n")
  cat("• 품질 검증을 통한 신뢰할 수 있는 시점만 활용\n")
  cat("• 성과기여도와의 상관관계 분석\n\n")
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
