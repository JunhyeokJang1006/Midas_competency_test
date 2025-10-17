# Repository Guidelines

## Project Structure & Module Organization
- Root scripts (`factor_analysis_by_period*`, `machine_learning_algorithm*`, `regression_*`) drive competency scoring; retain adjacent `.mat` caches for reproducibility.
- `문항기반/` handles question correlations; `_weighted/` and `역량진단 데이터 확인/` cover weighting trials plus certification loaders and histograms.
- `성장단계/` and the 자가불소* folders cover CSR-linked growth analysis and talent-type modelling, yielding `MIDAS_Growth_*` workbooks plus PNG dashboards.
- `claude_code/` mirrors everything in R; source data stays under `D:/project/HR데이터/데이터`, outputs in `matlab_results`, logs in `matlab_runlog`.

## Build, Test, and Development Commands
- Full run: `matlab -batch "run('D:/project/HR데이터/matlab/factor_analysis_by_period.m')"` builds `allData`, exports Excel, and records `matlab_runlog/runlog.txt`.
- Question analysis: `matlab -batch "run('D:/project/HR데이터/matlab/문항기반/corr_item_vs_comp_score.m')"`; swap to `_weighted` scripts when validating new coefficients.
- Cost-sensitive scoring: `matlab -batch "run('D:/project/HR데이터/matlab/자가불소_현주CP님 피드백/competency_statistical_analysis_order_logistic_revised.m')"` once cost-weight `.mat` files are writable.
- R port: `Rscript D:/project/HR데이터/matlab/claude_code/factor_analysis_by_period_peer_cc.R` after a one-time `source('install_packages.R')`.

## Coding Style & Naming Conventions
- Begin scripts with `clear; clc; close all;` plus `%%` headers and document any `rng` seeds before random logic.
- Indent four spaces, align wrapped statements, and keep descriptive snake_case filenames aligned with the main function name.
- Keep `[단계]`-style `fprintf` banners and wrap I/O in `try/catch` so diary logs stay readable.
- Group path constants near the top, add bilingual comments for new terms, and use `writetable(...,'Encoding','UTF-8')` to keep Hangul intact.

## Testing Guidelines
- Smoke tests: run `문항기반/test_improved_standardization.m`, `자가불소_개선/test_logitboost.m`, and `CSR/test_simple_save.m` via `matlab -batch`.
- Augment these scripts with `assert` checks or small summary tables so diary logs capture pass/fail context.
- Use `create_test_data.m` or the latest `competency_correlation_workspace_*.mat` as fixtures, and note expected filenames, sheet counts, and KPI thresholds in the header.

## Commit & Pull Request Guidelines
- Adopt Conventional Commit prefixes (`feat:`, `fix:`, `refactor:`, `docs:`) with ≤72-character subjects once git tracking is enabled.
- Summaries must cite entry script(s), required data folders, and whether new `.mat`/`.xlsx` files replace earlier deliverables.
- Attach key artefacts (e.g., `표준화영향분석_시각화.png`), note new MATLAB/R dependencies, and tag the module lead with before/after KPIs.

## Agent Configuration
- Approval mode stays `never`; sandbox operates as `danger-full-access` for local commands.
- Model runs on `gpt-5-codex` from OpenAI with high reasoning effort and automatic reasoning summaries.
