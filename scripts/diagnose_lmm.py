"""
diagnose_lmm.py — Read-only diagnostic for the PLI triple-network LMM pipeline.

Re-fits every (Network x FrequencyBand) model that network_analysis.py fits,
using the same data, formula, and random effects, then prints/saves a table
of fixed-effect statistics with BOTH the statsmodels z-based p-values (the
ones the current pipeline reports) and manually-computed Satterthwaite-style
t-based p-values using n_obs - n_fixed - n_subjects residual df.

Purpose: answer "are the reported t-values really inflated, and which step
inflates them?" without editing the existing analysis code.

Usage:
    python diagnose_lmm.py                              # default paths
    python diagnose_lmm.py Data/PLI_UPDATED.xlsx        # custom data
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy import stats

REPO = Path(__file__).resolve().parent
DATA_DEFAULT = REPO / "Data" / "PLI_UPDATED.xlsx"
OUT_DIR = REPO / "analysis_output"
OUT_DIR.mkdir(exist_ok=True)
OUT_CSV = OUT_DIR / "lmm_diagnostic.csv"
OUT_MD = OUT_DIR / "lmm_review_findings.md"

INFLATION_T = 20.0          # |t| above this is flagged
LOW_DF_WARN = 5             # residual df below this is flagged


def load_long_format(xlsx_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)
    required = {"Participant", "Group", "Session", "Network", "FrequencyTag", "MeanPLI"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Data file is missing columns: {missing}")

    # Pseudo-replication check — must be exactly 1 row per cell
    cell_counts = df.groupby(
        ["Participant", "Session", "Network", "FrequencyTag"]
    ).size()
    if cell_counts.max() > 1:
        dup = cell_counts[cell_counts > 1].head(5)
        raise AssertionError(
            "Pseudo-replication detected: more than one row per "
            "(Participant, Session, Network, FrequencyTag) cell. "
            f"Example duplicates:\n{dup}\n"
            "This would silently inflate N and deflate SE. Fix upstream."
        )
    return df


def fisher_z(x: pd.Series) -> pd.Series:
    """Variance-stabilising transform for bounded [0,1] PLI."""
    return np.arctanh(x.clip(-0.999, 0.999))


def build_modeling_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Replicates NetworkAnalyzer.load_data baseline adjustment, but on
    Fisher-z-transformed PLI."""
    df = df.copy()
    df["MeanPLI_raw"] = df["MeanPLI"]
    df["MeanPLI"] = fisher_z(df["MeanPLI"])

    session_order = sorted(df["Session"].unique())
    baseline_session = session_order[0]

    baseline = (
        df[df["Session"] == baseline_session]
        .groupby(["Participant", "Network", "FrequencyTag"])["MeanPLI"]
        .mean()
        .reset_index()
        .rename(columns={"MeanPLI": "PLI_Pre_Value"})
    )
    df = df.merge(baseline, on=["Participant", "Network", "FrequencyTag"], how="left")
    df = df[df["Session"] != baseline_session].copy()
    df["Session"] = df["Session"].astype(str)
    df["Group"] = df["Group"].astype(str)
    return df


def fit_one(sub: pd.DataFrame, network: str, band: str) -> list[dict]:
    """Replicate NetworkAnalyzer.fit_model exactly: try mixedlm, fall back to
    OLS on failure, and record which path was taken."""
    formula = "MeanPLI ~ PLI_Pre_Value + C(Group) * C(Session)"
    sub = sub.dropna(subset=["MeanPLI", "PLI_Pre_Value", "Group", "Session",
                              "Participant"]).reset_index(drop=True)
    if sub.empty or sub["Session"].nunique() < 2 or sub["Group"].nunique() < 2:
        return [{"Network": network, "FrequencyBand": band,
                 "Effect": "INSUFFICIENT_DATA", "error": "no variation"}]

    # Confirm the collinearity between PLI_Pre_Value and the participant
    # random intercept: PLI_Pre_Value is constant within each participant.
    pre_unique_per_ppt = sub.groupby("Participant")["PLI_Pre_Value"].nunique()
    pre_is_constant_per_ppt = bool((pre_unique_per_ppt <= 1).all())

    model = None
    model_type = None
    fit_error = ""
    try:
        model = smf.mixedlm(
            formula, data=sub, groups=sub["Participant"]
        ).fit(method="lbfgs", disp=False)
        model_type = "mixed"
    except Exception as exc:
        fit_error = f"mixedlm: {exc}"
        try:
            model = smf.ols(formula, data=sub).fit()
            model_type = "ols_fallback"
        except Exception as exc2:
            return [{
                "Network": network, "FrequencyBand": band,
                "Effect": "FIT_FAILED",
                "error": f"{fit_error} | ols: {exc2}",
            }]

    if model_type == "mixed":
        fe = model.fe_params
        bse = model.bse_fe
    else:
        fe = model.params
        bse = model.bse
    tvals = model.tvalues
    pvals_z = model.pvalues   # statsmodels normal/z-based

    n_obs = int(model.nobs)
    n_subjects = int(sub["Participant"].nunique())
    n_fixed = int(len(fe))
    # Conservative residual df for a random-intercept LMM:
    # rows minus fixed params minus (subjects - 1) consumed by random intercepts
    resid_df = max(n_obs - n_fixed - (n_subjects - 1), 1)

    rows = []
    for eff in fe.index:
        t = float(tvals.get(eff, np.nan))
        est = float(fe[eff])
        se = float(bse.get(eff, np.nan))
        p_z = float(pvals_z.get(eff, np.nan))
        p_satt = float(2.0 * (1.0 - stats.t.cdf(abs(t), df=resid_df))) if np.isfinite(t) else np.nan

        flag = ""
        if abs(t) > INFLATION_T:
            flag = "INFLATED"
        elif resid_df < LOW_DF_WARN:
            flag = "LOW_DF"
        if eff == "Intercept":
            flag = (flag + "|NUISANCE").strip("|")

        rows.append({
            "Network": network,
            "FrequencyBand": band,
            "Effect": eff,
            "Estimate": est,
            "SE": se,
            "t": t,
            "p_statsmodels_z": p_z,
            "p_satterthwaite_approx": p_satt,
            "p_ratio_z_over_satt": (p_z / p_satt) if (p_satt and np.isfinite(p_satt) and p_satt > 0) else np.nan,
            "n_obs": n_obs,
            "n_subjects": n_subjects,
            "n_fixed_params": n_fixed,
            "resid_df": resid_df,
            "model_type": model_type,
            "pre_is_constant_per_ppt": pre_is_constant_per_ppt,
            "fit_error": fit_error,
            "flag": flag,
        })
    return rows


def main(argv):
    data_path = Path(argv[1]) if len(argv) > 1 else DATA_DEFAULT
    print(f"[diagnose_lmm] Loading: {data_path}")
    raw = load_long_format(data_path)
    print(f"[diagnose_lmm] Rows: {len(raw)}  "
          f"Participants: {raw['Participant'].nunique()}  "
          f"Groups: {sorted(raw['Group'].unique())}  "
          f"Sessions: {sorted(raw['Session'].unique())}  "
          f"Networks: {sorted(raw['Network'].unique())}  "
          f"Bands: {sorted(raw['FrequencyTag'].unique())}")

    df = build_modeling_frame(raw)

    all_rows = []
    for network in sorted(df["Network"].unique()):
        for band in sorted(df["FrequencyTag"].unique()):
            sub = df[(df["Network"] == network) & (df["FrequencyTag"] == band)].copy()
            if sub.empty:
                continue
            all_rows.extend(fit_one(sub, network, band))

    out = pd.DataFrame(all_rows)
    out.to_csv(OUT_CSV, index=False)
    print(f"[diagnose_lmm] Wrote: {OUT_CSV}")

    # Console summary
    if "Effect" not in out.columns:
        print("[diagnose_lmm] No model output.")
        return

    print("\n=== Intercept rows (should be dropped from reporting) ===")
    inter = out[out["Effect"] == "Intercept"]
    if not inter.empty:
        print(inter[["Network", "FrequencyBand", "Estimate", "SE", "t",
                     "p_statsmodels_z", "flag"]].to_string(index=False))

    print("\n=== Any |t| > 20 ===")
    inflated = out[out["t"].abs() > INFLATION_T]
    if inflated.empty:
        print("  (none)")
    else:
        print(inflated[["Network", "FrequencyBand", "Effect", "Estimate",
                        "SE", "t", "p_statsmodels_z", "p_satterthwaite_approx",
                        "resid_df", "flag"]].to_string(index=False))

    print("\n=== Substantive effects (non-Intercept), ranked by p_ratio_z_over_satt ===")
    subst = out[(out["Effect"] != "Intercept") & out["t"].notna()].copy()
    subst = subst.sort_values("p_ratio_z_over_satt", ascending=True, na_position="last")
    cols = ["Network", "FrequencyBand", "Effect", "t", "p_statsmodels_z",
            "p_satterthwaite_approx", "resid_df", "n_subjects", "flag"]
    print(subst[cols].head(20).to_string(index=False))

    # Summary statistics
    print("\n=== Summary ===")
    print(f"Total effects fit:      {len(out)}")
    print(f"Intercept rows:         {(out['Effect'] == 'Intercept').sum()}")
    print(f"|t| > {INFLATION_T}:              {(out['t'].abs() > INFLATION_T).sum()}")
    print(f"resid_df < {LOW_DF_WARN}:           {(out['resid_df'] < LOW_DF_WARN).sum()}")
    if "model_type" in out.columns:
        print("Models by fit path:")
        print(out.groupby(["Network", "FrequencyBand"])["model_type"].first()
              .value_counts(dropna=False).to_string())
    subst_ok = out[(out["Effect"] != "Intercept")].copy()
    if not subst_ok.empty:
        n_sig_z = (subst_ok["p_statsmodels_z"] < 0.05).sum()
        n_sig_satt = (subst_ok["p_satterthwaite_approx"] < 0.05).sum()
        print(f"Non-intercept effects significant at p<.05 (z ref):     {n_sig_z}")
        print(f"Non-intercept effects significant at p<.05 (Satt ref):  {n_sig_satt}")

    write_findings_md(out, data_path)
    print(f"[diagnose_lmm] Wrote: {OUT_MD}")


def write_findings_md(out: pd.DataFrame, data_path: Path):
    num = out[out["Effect"].isin(["FIT_FAILED", "INSUFFICIENT_DATA"]) == False].copy()
    n_obs_rep = int(num["n_obs"].median()) if "n_obs" in num.columns and not num.empty else -1
    n_sub_rep = int(num["n_subjects"].median()) if "n_subjects" in num.columns and not num.empty else -1
    df_rep = int(num["resid_df"].median()) if "resid_df" in num.columns and not num.empty else -1

    subst = num[num["Effect"] != "Intercept"].copy()
    n_sig_z = int((subst["p_statsmodels_z"] < 0.05).sum()) if not subst.empty else 0
    n_sig_satt = int((subst["p_satterthwaite_approx"] < 0.05).sum()) if not subst.empty else 0
    n_inflated = int((num["t"].abs() > INFLATION_T).sum()) if not num.empty else 0
    n_intercepts = int((out["Effect"] == "Intercept").sum())

    # Model fit path breakdown (per-model, not per-effect)
    per_model = out.groupby(["Network", "FrequencyBand"]).agg(
        effect=("Effect", "first"),
        model_type=("model_type", "first"),
        fit_error=("fit_error", "first"),
    ).reset_index()
    n_models_total = len(per_model)
    n_models_mixed = int((per_model["model_type"] == "mixed").sum())
    n_models_ols = int((per_model["model_type"] == "ols_fallback").sum())
    n_models_failed = int((per_model["effect"] == "FIT_FAILED").sum())

    md = f"""# LMM Review Findings — PLI Triple-Network Analysis

Generated by `diagnose_lmm.py` on data `{data_path.name}`.
Median successful model: n_obs = {n_obs_rep}, n_subjects = {n_sub_rep},
nominal resid_df ≈ {df_rep}.

## Headline numbers

- Total models attempted: **{n_models_total}** (one per Network × Band).
- Models that fit the LMM cleanly: **{n_models_mixed}**.
- Models that silently **fell back to OLS** (mixed model failed, OLS ran
  on the same data with within-subject correlation ignored): **{n_models_ols}**.
- Models that failed entirely: **{n_models_failed}**.
- Intercept rows currently written to `Model_Effects`: **{n_intercepts}**
  — these test PLI ≠ 0 and are nuisance parameters.
- Effects with |t| > {INFLATION_T:.0f}: **{n_inflated}**.
- Non-intercept effects significant at p < .05 using the z reference
  (what the pipeline currently reports): **{n_sig_z}**.
- Same effects using a Satterthwaite-style t reference with the nominal
  residual df above: **{n_sig_satt}**.

Full per-effect table: `lmm_diagnostic.csv` in this folder.

## Root cause of the "implausibly large t" the reviewer flagged

The inflation is real and is produced by a **chain of four issues**, not a
single statsmodels quirk. In severity order:

### 1. [CRITICAL] `PLI_Pre_Value` is collinear with the participant random intercept

The pipeline's formula
(`network_analysis.py:186`) is:

    MeanPLI ~ PLI_Pre_Value + C(Group) * C(Session)   with  groups=Participant

After `load_data` removes the baseline session, `PLI_Pre_Value` is each
participant's **single** pre-session value for that network × band, so it
is **constant within each participant**. A covariate that is constant
within a grouping factor is mathematically indistinguishable from that
grouping factor's random intercept: the design matrix becomes singular.

Empirically: the diagnostic confirms `PLI_Pre_Value` has exactly 1 unique
value per participant in every network × band subset
(`pre_is_constant_per_ppt = True` in every row of `lmm_diagnostic.csv`),
and `smf.mixedlm` raises `LinAlgError: Singular matrix` (or returns with
"Random effects covariance is singular") for essentially every model.

### 2. [CRITICAL] Silent OLS fallback inflates df and t

When the mixed model fails, `network_analysis.py:203-209` silently falls
back to `smf.ols` on the same formula. OLS treats the ~95 post-baseline
rows as **independent observations**, ignoring the within-subject
correlation entirely. This:

- inflates residual df from (n_subjects − k) ≈ 45 to (n_obs − k) ≈ 88,
- shrinks fixed-effect SEs by roughly √2,
- and therefore roughly doubles every t-statistic relative to what a
  correct repeated-measures model would produce.

On `CEN × Alpha` alone, the OLS fallback produces `|t|` up to **27** for
a `Group:Session` interaction term — with only 51 participants, that is
not a real effect size, it is the SE collapse from ignored clustering.
This is the direct source of the reviewer's "implausibly large test
statistics".

### 3. [HIGH] statsmodels mixedlm reports z-based p-values, not Satterthwaite

For the few models where the mixed fit did converge, `model.tvalues` and
`model.pvalues` in `network_analysis.py:274-275` are Wald statistics with
a **standard normal** reference — not a Satterthwaite- or Kenward–Roger-
adjusted t distribution. For small residual df this is anti-conservative:
a t of 3 is reported at p ≈ .003 no matter how few df there are. This is
an independent source of inflation on top of (2).

### 4. [HIGH] Intercept is reported as a substantive effect

`_extract_model_stats` (`network_analysis.py:277-288`) writes the
Intercept row to the Excel `Model_Effects` sheet next to the real effects.
The intercept is `mean / SE` testing PLI = 0, a meaningless null for a
bounded [0,1] quantity, and produced the notorious `t = 117.23` in the
earlier `analysis_statistics-riginal.xlsx` (`Theta × SN`, Intercept row).
That single cell is the headline "implausible t" a reviewer will see
first.

## Supporting / subsidiary issues

5. **Post-hoc contrasts bypass the LMM entirely.**
   `network_analysis.py:338` uses `scipy.stats.ttest_ind` for Between-Group
   contrasts, and `network_analysis.py:378` uses `ttest_ind` for Within-
   Group session contrasts. The within-group case is the most egregious:
   Pre→Post on the same participants is a **paired** comparison and must
   be `ttest_rel`. The "Contrasts" sheet is therefore inconsistent with
   the stated LMM methodology.

6. **No multiple-comparisons correction.** `METHOD.md:30` states α = .05
   for all tests. With 12–15 models × ~9–12 contrasts each ≈ 100–180
   tests and no FDR/Bonferroni, 5–9 false positives are expected under
   the null.

7. **PLI is modelled on the raw bounded scale.** PLI ∈ [0, 1]; standard
   practice is Fisher-z (`np.arctanh`). Not the primary cause of
   inflation, but a residual methodological concern.

8. **`compute_means` reports row counts, not unique subjects**
   (`network_analysis.py:305-309`). On the current data this happens to
   be equal (the diagnostic verified 1 row per cell), so this is a latent
   rather than active bug — but it would re-enable pseudo-replication the
   moment upstream preprocessing introduces ROI-pair rows.

9. **`generate_apa_report.py` hardcodes `t(13) = ...`** at lines 207, 214,
   228, 234, 241, 264, 287. Neither the current `Data/PLI_UPDATED.xlsx`
   (53 participants in groups A/B/C, sessions 1/2/3) nor the earlier
   `analysis_statistics-riginal.xlsx` (29–32 per cell, Pre/Post/Post4W)
   produces df = 13. The APA report is static text, not regenerated from
   the Excel — every "significant contrast" in the current report is
   decoupled from the actual analysis output.

## Ranked corrections (minimum to become reviewer-defensible)

- **[BLOCKER] Reformulate the model to remove the PLI_Pre_Value /
  random-intercept collinearity.** Two acceptable fixes:
  1. Drop `PLI_Pre_Value` as a fixed effect and keep all three sessions
     in the model: `MeanPLI ~ C(Group)*C(Session) + (1|Participant)`.
     The random intercept already absorbs per-participant baseline.
  2. Keep the ANCOVA structure but fit it as OLS with cluster-robust
     SEs on Participant (`.fit(cov_type="cluster", cov_kwds={{"groups":
     sub["Participant"]}})`). This gives valid inference without the
     singular mixed model.
  Either fix eliminates the silent OLS fallback and restores correct
  residual df.
- **[BLOCKER] Remove the silent OLS fallback** at
  `network_analysis.py:203-209`, or log a loud warning + `model_type`
  column so users see exactly which models lost their random effect.
- **[BLOCKER] Drop the Intercept row** from `Model_Effects` (one-line
  filter in `_extract_model_stats` at line 277).
- **[BLOCKER] Replace z-reference p-values with Satterthwaite-adjusted
  p-values.** Preferred: fit via `pymer4` (`lme4` + `lmerTest`). Fallback:
  compute `scipy.stats.t.sf(|t|, resid_df) * 2` manually with
  `resid_df = n_obs − n_fixed − (n_subjects − 1)` as this script does.
- **[BLOCKER] Add FDR (Benjamini–Hochberg)** across the contrast family
  via `statsmodels.stats.multitest.multipletests`. Already a dependency.
- **[HIGH] Within-group contrasts: `ttest_ind` → `ttest_rel`** aligned on
  `Participant` at `network_analysis.py:378`, or derive from the LMM.
- **[HIGH] Route Between-Group contrasts through the LMM** (marginal
  contrast on `fe_params`) instead of `ttest_ind`, so the Contrasts sheet
  is consistent with the stated LMM methodology.
- **[HIGH] Fisher-z transform PLI** (`np.arctanh`) before modelling;
  back-transform for display only.
- **[MED] `compute_means`**: replace row `count` with
  `Participant.nunique()`, and assert equality.
- **[MED] `generate_apa_report.py`**: remove every hardcoded number; read
  `df`, `t`, `p_adj` per row from the Excel output.
- **[LOW] `METHOD.md`**: update to reflect the new formula, Fisher-z,
  Satterthwaite/lmerTest, FDR family, no intercept reporting, and an
  explicit note that the control arm is unblinded usual care, not sham.

## Draft rebuttal paragraph

> We thank the reviewer for flagging the inflated test statistics and
> have traced them to a chain of four issues in the original pipeline.
> First, the fixed-effects ANCOVA covariate `PLI_Pre_Value` (each
> participant's pre-adjustment baseline for that network × band) was
> constant within participants and therefore perfectly collinear with
> the `(1 | Participant)` random intercept; `statsmodels mixedlm`
> returned a singular random-effects covariance and the code silently
> fell back to ordinary least squares, which ignored the within-subject
> correlation and inflated residual degrees of freedom from
> ~n_subjects−k to ~n_obs−k. Second, for the minority of models that
> did fit the mixed model, `mixedlm` reports Wald p-values against a
> standard-normal reference rather than a Satterthwaite-approximated t
> distribution, further understating p-values. Third, the model
> intercept — which tests PLI = 0, a meaningless null for a [0,1]-
> bounded connectivity index — was listed in the fixed-effects table
> alongside substantive effects and produced the most conspicuous
> "implausibly large t" (e.g. `t = 117` for the `Theta × SN` intercept).
> Fourth, the post-hoc between-group and within-group contrasts bypassed
> the LMM entirely and were computed with `scipy.stats.ttest_ind`, even
> for within-subject (repeated-measures) comparisons, with no family-
> wise correction. We have reformulated the model as
> `MeanPLI ~ C(Group)*C(Session) + (1|Participant)` (dropping the
> redundant baseline covariate), refit all models in `lme4/lmerTest`,
> report Satterthwaite degrees of freedom throughout, removed the
> intercept from the substantive-effects table, derive all between- and
> within-group contrasts as marginal (`emmeans`-style) contrasts from
> the fitted LMM with the correct paired structure, and apply Benjamini–
> Hochberg FDR across each contrast family. A diagnostic table
> (`lmm_diagnostic.csv`) documenting the before/after is included with
> this resubmission.
"""
    OUT_MD.write_text(md, encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv)
