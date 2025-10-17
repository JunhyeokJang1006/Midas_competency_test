# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a MATLAB-based HR analytics project focused on competency assessment and factor analysis. The project analyzes HR competency assessment data across multiple time periods (2023 H2 - 2025 H1) and performs statistical analysis including factor analysis, correlation analysis, and performance evaluation.

## Key Files and Architecture

### Main Analysis Scripts
- **factor_analysis_by_period.m** - Primary analysis script for individual factor analysis by time period. Loads competency assessment data from 4 periods, performs factor analysis independently for each period, calculates performance scores, and conducts correlation analysis with performance contribution measures.

### Analysis Variants
- **factor_analysis_by_period_*.m** (v2, gpt, grok, peer versions) - Alternative implementations and variations of the main period-based analysis

### Data Structure
The project works with Excel files containing:
- **기준인원 검토** sheet - Master participant IDs
- **하향 진단** sheet - Downward assessment data  
- **문항 정보_타인진단** sheet - Question information for peer assessment
- **팀장하향평가_보정**, **24상진단_하향_보정**, **24상진단_수평_보정** sheets - Performance evaluation data

### Output Files
- **competency_correlation_workspace_*.mat** - Workspace files with correlation analysis results
- **competency_performance_correlation_results_*.xlsx** - Excel reports with correlation results
- **역량검사진단_분석결과_*.mat/.xlsx** - Competency assessment analysis results
- Various analysis result files with timestamps

## Common Commands

### Running Main Analysis
```matlab
cd('D:\project\HR데이터\matlab')
factor_analysis_by_period  % Run the main period-based analysis
```

### Data Loading Paths
- Raw data location: `D:\project\HR데이터\데이터\역량진단 데이터\역량진단_응답데이터`
- Working directory: `D:\project\HR데이터\matlab`
- Log directory: `D:\project\matlab_runlog\`

### Key Analysis Periods
- 23년_하반기 (2023 H2)
- 24년_상반기 (2024 H1) 
- 24년_하반기 (2024 H2)
- 25년_상반기 (2025 H1)

## Important Notes

- Scripts automatically exclude interview-only questions (면담용 문항) from analysis
- Special handling for 25년 상반기 data excludes Q40-46 questions
- All scripts use Korean language for output and logging
- Results are automatically saved with timestamps
- Scripts include comprehensive error handling and progress reporting

## Analysis Features

- Individual factor analysis for each time period
- Automatic identification and exclusion of non-scoring questions
- Correlation analysis between competency scores and performance measures
- Comprehensive statistical reporting with Excel output