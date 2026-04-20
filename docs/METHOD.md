### Method

#### **Data Preprocessing and Preparation**

The analysis was conducted using a custom Python script leveraging several open-source libraries, including pandas (v1.5.0), numpy (v1.21.0), statsmodels (v0.13.0), and scipy (v1.9.0). The raw data was loaded from an Excel file containing Phase Lag Index (PLI) values, with each row corresponding to a specific participant, session, brain network, and frequency band.

The dataset was initially filtered to include the selected groups, sessions, networks, and frequency bands. A pseudo-replication check was performed to ensure exactly one observation per (Participant, Session, Network, FrequencyTag) cell.

PLI values were Fisher-z transformed (arctanh) prior to modelling. PLI is bounded on [0, 1] and its sampling distribution is skewed near the bounds; the Fisher-z transform stabilises variance and is the standard preprocessing step for bounded connectivity measures before linear modelling. Group means, standard deviations, and standard errors reported in tables and figures are back-transformed to the original PLI scale for interpretability.

#### **Statistical Analysis**

To investigate the effects of group and session on network-level PLI, a separate model was fitted for each combination of brain network and frequency band.

The model was specified as:

*   **Dependent Variable**: Fisher-z-transformed PLI (`MeanPLI_z`)
*   **Fixed Effects**: `Group`, `Session`, and the `Group × Session` interaction term.
*   **Random Effect**: A random intercept for each `Participant`, which absorbs between-subject baseline differences and accounts for the repeated-measures structure of the design.

When a random-intercept model is selected (the default), the model is fitted via restricted maximum likelihood (REML) using the `lbfgs` optimiser in `statsmodels.formula.api.mixedlm`. If the random-effects covariance is singular (i.e., the participant random-effect variance converges to zero), the model falls back to ordinary least squares (OLS) with cluster-robust standard errors clustered on Participant, so that within-subject correlation is still honoured. The model type (`mixed` or `ols_cluster`) is recorded in the output for transparency.

*Note on the previous baseline-covariate specification.* An earlier version of the pipeline included each participant's pre-session PLI as a fixed-effect covariate (`PLI_Pre_Value`) alongside the `(1 | Participant)` random intercept. Because the baseline value is constant within each participant, it was perfectly collinear with the random intercept, causing a singular design matrix for most models and triggering a silent fallback to OLS without cluster correction—thereby inflating degrees of freedom, deflating standard errors, and producing implausibly large test statistics. The current specification avoids this by relying solely on the random intercept to absorb per-participant baseline differences.

Fixed-effect p-values from the mixed model are computed against a *t* distribution with approximate Satterthwaite-style residual degrees of freedom (`n_obs − n_fixed − (n_subjects − 1)`), rather than the large-sample normal (Wald z) reference that `statsmodels` uses by default. This correction is important for studies where the number of participants is not large enough for the normal approximation to be adequate. The Intercept (which tests PLI = 0, a meaningless null for a bounded connectivity measure) is excluded from the reported fixed-effects table.

#### **Post-Hoc Contrasts**

Following the modelling, post-hoc contrasts were performed to examine specific group and session differences.

1.  **Between-Group Contrasts**: For each session, PLI values were compared across groups using an independent-samples *t*-test (`scipy.stats.ttest_ind`), with degrees of freedom `n₁ + n₂ − 2`. Cohen's *d* (pooled SD) is reported alongside each test.
2.  **Within-Group Contrasts**: To assess changes over time within each group, PLI values between pairs of sessions were compared using a **paired-samples** *t*-test (`scipy.stats.ttest_rel`), aligning observations on Participant, with `df = n_pairs − 1`. Cohen's *d_z* (mean difference / SD of differences) is reported.

#### **Multiple-Comparisons Correction**

Benjamini–Hochberg false-discovery-rate (FDR) correction is applied separately within two contrast families: (a) all between-group contrasts and (b) all within-group contrasts. FDR-adjusted p-values (`p_adj`) and the corresponding significance flags (`Significance_FDR`) are reported in the Excel output alongside the uncorrected values. Fixed-effect p-values in the Model_Effects sheet are likewise FDR-corrected across all model effects. A threshold of q < 0.05 is used.

#### **Data Visualization and Reporting**

The results of the analysis were visualized by generating line plots for each network and frequency band combination. These plots displayed the mean PLI values for each group across all sessions, with error bars representing the standard error of the mean (SE). Significance markers (`*` for p < 0.05, `**` for p < 0.01, `***` for p < 0.001) derived from the model were overlaid on the plots to indicate significant session effects or group-by-session interactions.

A comprehensive report of all statistical results, including the fixed effects estimates, degrees of freedom, *t*-values, uncorrected and FDR-corrected p-values, Cohen's *d*, and the model type (mixed vs. cluster-robust OLS), was exported to an Excel file with three sheets: `Model_Effects`, `Contrasts`, and `Group_Means`. The `N` column in `Group_Means` reports the number of unique participants (not rows) per cell. A combined figure showing all plots in a grid layout was also generated.
