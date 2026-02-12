### Method

#### **Data Preprocessing and Preparation**

The analysis was conducted using a custom Python script leveraging several open-source libraries, including pandas (v1.5.0), numpy (v1.21.0), statsmodels (v0.13.0), and scipy (v1.9.0). The raw data was loaded from an Excel file containing Phase Lag Index (PLI) values, with each row corresponding to a specific participant, session, brain network, and frequency band.

The dataset was initially filtered to include only the selected groups (e.g., "Chiro", "Control"), sessions (e.g., "Pre", "Post", "Post4W"), networks (e.g., "DMN", "SN", "CEN"), and frequency bands (e.g., "Alpha", "Beta").

A baseline adjustment was performed to account for individual differences in initial PLI values. The mean PLI value from the first session (designated as 'Pre') was calculated for each participant, network, and frequency band. This baseline value was then included as a covariate in the statistical model. Consequently, the 'Pre' session data was excluded from the main analysis of session effects.

#### **Statistical Analysis**

To investigate the effects of group and session on network-level PLI, a separate linear mixed-effects model was fitted for each combination of brain network and frequency band using the `statsmodels` library.

The model was specified as:

*   **Dependent Variable**: `MeanPLI`
*   **Fixed Effects**: `Group`, `Session`, and the `Group * Session` interaction term. When baseline adjustment was applied, the baseline PLI value (`PLI_Pre_Value`) was also included as a fixed-effect covariate.
*   **Random Effect**: A random intercept was included for each `Participant` to account for repeated measures and variability among individuals.

The model was fitted using the `lbfgs` optimization algorithm. The significance of the fixed effects was evaluated to determine the main effects of group and session, as well as their interaction.

#### **Post-Hoc Contrasts**

Following the mixed-effects modeling, post-hoc contrasts were performed to examine specific group and session differences.

1.  **Between-Group Contrasts**: For each session, the mean PLI values of the two groups were compared using an independent samples t-test (`scipy.stats.ttest_ind`). This identified at which specific time points the groups differed significantly.
2.  **Within-Group Contrasts**: To assess changes over time within each group, the mean PLI values between pairs of sessions (e.g., 'Pre' vs. 'Post', 'Post' vs. 'Post4W') were compared using independent samples t-tests.

For all t-tests, a p-value less than 0.05 was considered statistically significant.

#### **Data Visualization and Reporting**

The results of the analysis were visualized by generating line plots for each network and frequency band combination. These plots displayed the mean PLI values for each group across all sessions, with error bars representing the standard error of the mean (SE). Significance markers (`*` for p < 0.05, `**` for p < 0.01, `***` for p < 0.001) derived from the mixed-effects model were overlaid on the plots to indicate significant session effects or group-by-session interactions.

A comprehensive report of all statistical results, including the fixed effects estimates from the mixed-effects models, p-values, and the results of all post-hoc contrasts, was exported to an Excel file for detailed review. Additionally, a combined figure showing all plots in a grid layout was generated to provide a comprehensive overview of the findings.
