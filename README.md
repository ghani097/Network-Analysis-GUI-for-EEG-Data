# PLI Network Analysis GUI

A graphical interface for analyzing Phase Lag Index (PLI) data across brain networks using mixed-effects models.

> **Latest Update (April 2026):** Major statistical pipeline overhaul — fixed collinear baseline covariate, added Fisher-z transform, Satterthwaite df, paired within-group tests, FDR correction, Cohen's d effect sizes, and pseudo-replication guard. See [Changelog](#changelog) below.

## Features

- Load PLI data from Excel files
- **Fisher-z (arctanh) transform** of PLI before modelling (back-transformed for display)
- **Linear mixed-effects models** with participant random intercepts, or cluster-robust OLS fallback
- **Satterthwaite-approximated residual degrees of freedom** (not large-sample z reference)
- Pairwise between-group contrasts (independent *t*-test, any number of groups)
- **Paired** within-group session contrasts (`ttest_rel`, aligned on Participant)
- **Benjamini–Hochberg FDR correction** across contrast families
- **Cohen's *d*** reported for every contrast
- Explicit **degrees of freedom** on all effects and contrasts
- **Pseudo-replication guard** (asserts 1 row per Participant × Session × Network × Band)
- Automatic visualization with significance markers and adaptive y-axis scaling
- Export results to Excel (with `p_adj`, `Significance_FDR`, `df`, `Cohens_d`, `ModelType` columns) and CSV
- Methodology documentation and analysis pipeline diagram included
- **APA-formatted report generation** (`scripts/generate_apa_report.py`)
- **Brain figure generation** (`scripts/generate_brain_figure.py`)
- **LMM diagnostic script** (`scripts/diagnose_lmm.py`)
- Robust handling of non-string Session/Group values in Excel input

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/PLI-Network-Analysis-GUI.git
   cd PLI-Network-Analysis-GUI
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Windows
Double-click `run_gui.bat` or run:
```bash
python scripts/network_analysis_gui.py
```

### Command Line
```bash
python scripts/network_analysis.py --input data.xlsx --output results
python scripts/network_analysis.py --no-baseline  # Cluster-robust OLS instead of LMM
```

### LMM Diagnostic
Run the standalone diagnostic to verify statistical rigour:
```bash
python scripts/diagnose_lmm.py                  # default data path
python scripts/diagnose_lmm.py path\to\your_data.xlsx
```
Outputs `analysis_output/lmm_diagnostic.csv` (per-effect table with z-based vs. Satterthwaite p-values) and `analysis_output/lmm_review_findings.md` (ranked findings + draft rebuttal paragraph).

## Input Data Format

The input Excel file should contain the following columns:
- `Participant` — Participant ID (numeric or string)
- `Group` — Group label (e.g., "A", "B", "C" or "Chiro", "Control")
- `Session` — Session name (e.g., "Pre", "Post", "Post4W" or 1, 2, 3)
- `Network` — Brain network (e.g., "DMN", "SN", "CEN")
- `FrequencyTag` — Frequency band (e.g., "Alpha", "Beta", "Delta", "Theta")
- `MeanPLI` — Mean Phase Lag Index value (one value per participant × session × network × band)

Sample test files are included in `Data/test/`. Updated PLI data is available in `Data/PLI_UPDATED.xlsx`.

## Output

The analysis generates:
- Individual plots for each network/frequency combination
- Combined results figure (`combined_results.png`)
- **Statistics Excel file** (`analysis_statistics.xlsx`) with three sheets:
  - `Model_Effects` — Fixed effects (no Intercept), with `df`, `p_adj`, `Significance_FDR`, `ModelType`
  - `Contrasts` — Between-group and within-group contrasts with `df`, `Cohens_d`, `p_adj`, `Significance_FDR`
  - `Group_Means` — Descriptive statistics with `N` = unique participants (not row counts)
- Summary CSV (`summary.csv`)

## Statistical Methods

### Model specification
For each frequency band × network combination, a linear mixed-effects model is fitted:

```
z(PLI) ~ C(Group) * C(Session) + (1 | Participant)
```

- **Dependent variable**: Fisher-z-transformed PLI
- **Fixed effects**: Group, Session, Group × Session interaction
- **Random effect**: Participant random intercept (absorbs between-subject baseline differences)
- **Fallback**: If the random-effects variance is singular, the model falls back to OLS with cluster-robust standard errors on Participant (warns loudly; model type is recorded)

### Inference
- Fixed-effect p-values use a *t* distribution with Satterthwaite-approximated residual df, not the statsmodels default large-sample z reference
- The Intercept (testing PLI = 0) is excluded from the reported effects table
- Between-group contrasts: independent *t*-test with pooled-SD Cohen's *d*
- Within-group contrasts: **paired** *t*-test (aligned on Participant) with Cohen's *d_z*
- **FDR correction**: Benjamini–Hochberg, applied separately within between-group and within-group contrast families

### Previous baseline-covariate issue (fixed)
An earlier version included each participant's pre-session PLI as a fixed-effect covariate (`PLI_Pre_Value`) alongside the `(1 | Participant)` random intercept. Because the baseline value is constant within each participant, it was perfectly collinear with the random intercept, causing singular fits for most models and triggering a silent fallback to naive (unclustered) OLS — inflating degrees of freedom, deflating standard errors, and producing implausibly large test statistics. This has been resolved by relying solely on the random intercept to absorb per-participant baseline differences.

## Documentation

- `docs/METHOD.md` / `docs/METHOD.docx` / `docs/METHOD.pdf` — Detailed methodology write-up suitable for publications
- `scripts/diagnose_lmm.py` — Standalone LMM diagnostic script
- `scripts/generate_apa_report.py` — Generate APA-formatted results report
- `scripts/generate_brain_figure.py` — Generate brain topography figures
- `docs/pipeline_diagram.svg` — Visual overview of the analysis pipeline
- `scripts/render_diagram.py` — Script to regenerate the pipeline diagram from `docs/pipeline_diagram.puml`

## Requirements

- Python 3.8+
- pandas
- numpy
- matplotlib
- scipy
- statsmodels
- PyQt5
- openpyxl

## Changelog

### April 2026 — Statistical Pipeline Overhaul

**Fixes (`scripts/network_analysis.py`):**
1. Removed collinear `PLI_Pre_Value` covariate that caused singular LMM fits and silent OLS fallback
2. OLS fallback now uses cluster-robust SEs on Participant (+ logs WARNING)
3. Intercept excluded from `Model_Effects` output
4. p-values recomputed against *t* distribution with Satterthwaite-approximated residual df
5. Within-group contrasts changed from `ttest_ind` to `ttest_rel` (paired, aligned on Participant)
6. Added Benjamini–Hochberg FDR correction (`p_adj`, `Significance_FDR` columns)
7. Added Fisher-z transform (`MeanPLI_z` for models, `MeanPLI_raw` for plots/means)
8. `compute_means` now reports `Participant.nunique()` for N (not row count)
9. Added Cohen's *d* to all contrasts
10. Added explicit `df` column to all effects and contrasts
11. Added `ModelType` column (mixed vs. ols_cluster)
12. Added pseudo-replication guard in `load_data`

**New files:**
- `scripts/diagnose_lmm.py` — standalone diagnostic script
- `analysis_output/lmm_diagnostic.csv` — per-effect diagnostic table
- `analysis_output/lmm_review_findings.md` — ranked findings + draft rebuttal paragraph

**Updated:**
- `docs/METHOD.md` — fully rewritten to document new statistical approach

## License

MIT License
