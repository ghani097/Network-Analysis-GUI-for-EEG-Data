"""
Network Level Analysis - Python Implementation
==============================================

Simple Python implementation of network-level PLI analysis.

Usage:
    python network_analysis.py
    python network_analysis.py --no-baseline
    python network_analysis.py --input data.xlsx --output results
"""

import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
from scipy import stats

try:
    import statsmodels.formula.api as smf
    from statsmodels.stats.anova import anova_lm
    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False


class NetworkAnalysis:
    """Network-level PLI analysis with mixed-effects models."""

    def __init__(self, input_file, output_dir="analysis_output",
                 adjust_baseline=True, groups=None, sessions=None,
                 networks=None, frequency_bands=None, callback=None):
        """
        Initialize analysis.

        Args:
            input_file: Path to Excel file with PLI data
            output_dir: Directory for output files
            adjust_baseline: If True, use Pre session as covariate
            groups: List of groups to include (default: all)
            sessions: List of sessions to include (default: all)
            networks: List of networks to analyze (default: all)
            frequency_bands: List of frequency bands (default: all)
            callback: Function to call with progress updates (msg, percent)
        """
        self.input_file = Path(input_file).resolve()
        self.output_dir = Path(output_dir).resolve()
        self.adjust_baseline = adjust_baseline
        self.callback = callback or (lambda msg, pct: print(msg))

        # All of these can be None → auto-detected from data in load_data()
        self.groups = groups  # resolved in load_data()
        self.sessions = sessions  # resolved in load_data()
        self.networks = networks  # resolved in load_data()
        self.frequency_bands = frequency_bands  # resolved in load_data()

        # Session order - set after sessions are resolved
        self.session_order = list(self.sessions) if self.sessions else []

        self.data = None
        self.results = {}
        self.all_stats = []  # Store all statistics for Excel export
        self.all_plots = {}  # Store plot data for combined figure
        self.contrast_results = {}  # Store contrast results for significance display

    def _update(self, msg, pct=None):
        """Send progress update."""
        if self.callback:
            self.callback(msg, pct)

    def _get_session_order(self, sessions):
        """Get sessions in correct order based on user-provided session order."""
        return [s for s in self.session_order if s in sessions]

    def load_data(self):
        """Load and preprocess data."""
        self._update(f"Loading: {self.input_file.name}", 5)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_file}")

        df = pd.read_excel(str(self.input_file))
        self._update(f"Loaded {len(df)} rows", 10)

        # Auto-detect groups if not specified
        data_groups = sorted(df['Group'].dropna().unique().tolist())
        if self.groups is None or not set(self.groups).intersection(data_groups):
            if self.groups:
                self._update(
                    f"Warning: groups {self.groups} not found in data. "
                    f"Auto-detecting: {data_groups}", None
                )
            self.groups = data_groups

        # Auto-detect sessions / networks / bands from the data when not specified,
        # or when the user-supplied values have no overlap with the actual data.
        data_sessions = sorted(df['Session'].dropna().unique().tolist())
        if self.sessions is None or not set(self.sessions).intersection(data_sessions):
            if self.sessions:
                self._update(
                    f"Warning: sessions {self.sessions} not found in data. "
                    f"Auto-detecting: {data_sessions}", None
                )
            self.sessions = data_sessions
            self.session_order = list(data_sessions)

        data_networks = sorted(df['Network'].dropna().unique().tolist())
        if self.networks is None or not set(self.networks).intersection(data_networks):
            if self.networks:
                self._update(
                    f"Warning: networks {self.networks} not found in data. "
                    f"Auto-detecting: {data_networks}", None
                )
            self.networks = data_networks

        data_bands = sorted(df['FrequencyTag'].dropna().unique().tolist())
        if self.frequency_bands is None or not set(self.frequency_bands).intersection(data_bands):
            if self.frequency_bands:
                self._update(
                    f"Warning: bands {self.frequency_bands} not found in data. "
                    f"Auto-detecting: {data_bands}", None
                )
            self.frequency_bands = data_bands

        # Filter data
        df = df[df['Group'].isin(self.groups)]
        df = df[df['Network'].isin(self.networks)]
        df = df[df['FrequencyTag'].isin(self.frequency_bands)]

        if self.adjust_baseline:
            # Extract baseline (first session) values - dynamically use first session
            baseline_session = self.session_order[0] if self.session_order else 'Pre'
            pre_data = df[df['Session'] == baseline_session].copy()
            baseline = pre_data.groupby(
                ['Participant', 'Network', 'FrequencyTag']
            )['MeanPLI'].mean().reset_index()
            baseline.columns = ['Participant', 'Network', 'FrequencyTag', 'PLI_Pre_Value']

            # Merge baseline and remove baseline session
            df = df.merge(baseline, on=['Participant', 'Network', 'FrequencyTag'], how='left')
            df = df[df['Session'] != baseline_session]

            valid_sessions = [s for s in self.sessions if s != baseline_session]
            df = df[df['Session'].isin(valid_sessions)]
            self._update(f"Baseline adjustment: ON (using {baseline_session} as baseline)", 15)
        else:
            df = df[df['Session'].isin(self.sessions)]
            self._update("Baseline adjustment: OFF", 15)

        # Set session as ordered categorical
        available_sessions = self._get_session_order(df['Session'].unique())
        df['Session'] = pd.Categorical(df['Session'], categories=available_sessions, ordered=True)

        self._update(f"Filtered: {len(df)} rows", 20)
        self.data = df
        return df

    def fit_model(self, network, freq_band):
        """Fit mixed-effects model for one network/frequency combination."""
        subset = self.data[
            (self.data['Network'] == network) &
            (self.data['FrequencyTag'] == freq_band)
        ].copy()

        if len(subset) == 0:
            return None

        model_name = f"{freq_band} x {network}"
        model_result = {'model': None, 'data': subset, 'name': model_name, 'anova': None, 'model_significance': {}}

        if HAS_STATSMODELS:
            try:
                if self.adjust_baseline and 'PLI_Pre_Value' in subset.columns:
                    formula = "MeanPLI ~ PLI_Pre_Value + C(Group) * C(Session)"
                else:
                    formula = "MeanPLI ~ C(Group) * C(Session)"

                model = smf.mixedlm(
                    formula, data=subset, groups=subset['Participant']
                ).fit(method='lbfgs', disp=False)

                model_result['model'] = model

                # Extract model statistics
                self._extract_model_stats(model, model_name, network, freq_band)

                # Extract significance for plotting (session effects and interactions)
                model_result['model_significance'] = self._extract_plot_significance(model, subset)

            except Exception as e:
                self._update(f"  Model error: {e}", None)

        return model_result

    def _extract_plot_significance(self, model, subset):
        """Extract significance for each session from model coefficients."""
        significance = {}
        if model is None:
            return significance

        pvalues = model.pvalues
        sessions = subset['Session'].unique()

        # Check each parameter for session-related significance
        for param in pvalues.index:
            p = pvalues[param]
            stars = self._get_sig_stars(p)

            if stars and stars != 'ns':
                # Check for interaction terms (Group:Session) - these indicate group differences at specific sessions
                if ':C(Session)' in param or 'C(Session)' in param and ':' in param:
                    # Extract session name from interaction term like "C(Group)[T.B]:C(Session)[T.post2]"
                    for session in sessions:
                        session_str = str(session)
                        if f'[T.{session_str}]' in param:
                            # Interaction term - group difference at this session
                            if session_str not in significance:
                                significance[session_str] = {'p': p, 'stars': stars, 'type': 'interaction'}
                            elif p < significance[session_str]['p']:
                                significance[session_str] = {'p': p, 'stars': stars, 'type': 'interaction'}

                # Also check main session effects
                elif 'C(Session)' in param and ':' not in param:
                    # Main effect of session like "C(Session)[T.post2]"
                    for session in sessions:
                        session_str = str(session)
                        if f'[T.{session_str}]' in param:
                            # Only add if no interaction significance already exists for this session
                            if session_str not in significance:
                                significance[session_str] = {'p': p, 'stars': stars, 'type': 'session_effect'}

        return significance

    def _extract_model_stats(self, model, model_name, network, freq_band):
        """Extract and store model statistics for Excel export."""
        if model is None:
            return

        # Get fixed effects
        fe = model.fe_params
        bse = model.bse_fe
        tvalues = model.tvalues
        pvalues = model.pvalues

        for param in fe.index:
            self.all_stats.append({
                'Model': model_name,
                'Network': network,
                'FrequencyBand': freq_band,
                'Effect': param,
                'Estimate': fe[param],
                'Std.Error': bse[param] if param in bse.index else np.nan,
                't-value': tvalues[param] if param in tvalues.index else np.nan,
                'p-value': pvalues[param] if param in pvalues.index else np.nan,
                'Significance': self._get_sig_stars(pvalues[param]) if param in pvalues.index else ''
            })

    def _get_sig_stars(self, p):
        """Get significance stars for p-value."""
        if pd.isna(p):
            return ''
        if p < 0.001:
            return '***'
        elif p < 0.01:
            return '**'
        elif p < 0.05:
            return '*'
        else:
            return 'ns'

    def compute_means(self, subset):
        """Compute group means by session."""
        means = subset.groupby(['Group', 'Session'], observed=True).agg({
            'MeanPLI': ['mean', 'std', 'count']
        }).reset_index()
        means.columns = ['Group', 'Session', 'Mean', 'SD', 'N']
        means['SE'] = means['SD'] / np.sqrt(means['N'])

        # Sort by session order
        session_order = self._get_session_order(means['Session'].unique())
        means['Session'] = pd.Categorical(means['Session'], categories=session_order, ordered=True)
        means = means.sort_values(['Group', 'Session'])

        return means

    def run_contrasts(self, subset, model_name, network, freq_band):
        """Run between-group and within-session contrasts."""
        results = {'group': [], 'session': [], 'significance': {}}
        output_lines = []
        contrast_stats = []

        # Get sessions in correct order
        sessions = self._get_session_order(subset['Session'].unique())

        # Between-group contrasts by session
        for session in sessions:
            s_data = subset[subset['Session'] == session]
            groups = sorted(s_data['Group'].unique())

            if len(groups) == 2:
                g1 = s_data[s_data['Group'] == groups[0]]['MeanPLI']
                g2 = s_data[s_data['Group'] == groups[1]]['MeanPLI']

                t, p = stats.ttest_ind(g1, g2)
                diff = g1.mean() - g2.mean()
                sig = self._get_sig_stars(p)

                # Store significance for plotting (only if significant)
                if sig and sig != 'ns':
                    results['significance'][session] = {'p': p, 'stars': sig}

                output_lines.append(f"  {session}: {groups[0]} vs {groups[1]} = {diff:+.4f}, p={p:.4f} {sig}")
                results['group'].append({
                    'session': session, 'contrast': f"{groups[0]} - {groups[1]}",
                    'diff': diff, 't': t, 'p': p
                })

                # Store for Excel
                contrast_stats.append({
                    'Model': model_name,
                    'Network': network,
                    'FrequencyBand': freq_band,
                    'ContrastType': 'Between-Group',
                    'Session': session,
                    'Group': '',
                    'Contrast': f"{groups[0]} vs {groups[1]}",
                    'Difference': diff,
                    't-value': t,
                    'p-value': p,
                    'Significance': sig
                })

        # Within-group session contrasts
        for group in sorted(subset['Group'].unique()):
            g_data = subset[subset['Group'] == group]

            for i, s1 in enumerate(sessions):
                for s2 in sessions[i+1:]:
                    d1 = g_data[g_data['Session'] == s1]['MeanPLI']
                    d2 = g_data[g_data['Session'] == s2]['MeanPLI']

                    if len(d1) > 0 and len(d2) > 0:
                        t, p = stats.ttest_ind(d1, d2)
                        diff = d1.mean() - d2.mean()
                        sig = self._get_sig_stars(p)

                        output_lines.append(f"  {group}: {s1} vs {s2} = {diff:+.4f}, p={p:.4f} {sig}")
                        results['session'].append({
                            'group': group, 'contrast': f"{s1} - {s2}",
                            'diff': diff, 't': t, 'p': p
                        })

                        # Store for Excel
                        contrast_stats.append({
                            'Model': model_name,
                            'Network': network,
                            'FrequencyBand': freq_band,
                            'ContrastType': 'Within-Group',
                            'Session': '',
                            'Group': group,
                            'Contrast': f"{s1} vs {s2}",
                            'Difference': diff,
                            't-value': t,
                            'p-value': p,
                            'Significance': sig
                        })

        return results, output_lines, contrast_stats

    def _get_group_colors(self, groups):
        """Get color mapping for groups dynamically."""
        # Color palette that works well for distinguishing groups
        color_palette = [
            '#E69F00',  # Orange
            '#56B4E9',  # Sky blue
            '#009E73',  # Green
            '#F0E442',  # Yellow
            '#0072B2',  # Blue
            '#D55E00',  # Red-orange
            '#CC79A7',  # Pink
            '#999999',  # Gray
        ]
        colors = {}
        for i, group in enumerate(sorted(groups)):
            colors[group] = color_palette[i % len(color_palette)]
        return colors

    def _get_group_markers(self, groups):
        """Get marker mapping for groups dynamically."""
        marker_palette = ['o', 's', '^', 'D', 'v', 'p', 'h', '*']
        markers = {}
        for i, group in enumerate(sorted(groups)):
            markers[group] = marker_palette[i % len(marker_palette)]
        return markers

    def create_plot(self, subset, model_name, significance=None, save=True):
        """Create visualization with correct session order and significance markers."""
        means = self.compute_means(subset)
        significance = significance or {}

        # Get sessions in correct order
        sessions = self._get_session_order(means['Session'].unique())

        fig, ax = plt.subplots(figsize=(6, 4))

        # Get dynamic colors and markers for groups
        unique_groups = means['Group'].unique()
        colors = self._get_group_colors(unique_groups)
        markers = self._get_group_markers(unique_groups)

        # Track max y values for significance marker placement
        max_y_per_session = {}

        for group in sorted(means['Group'].unique()):
            gdata = means[means['Group'] == group]
            # Ensure correct order
            gdata = gdata.set_index('Session').reindex(sessions).reset_index()

            color = colors[group]
            marker = markers[group]

            ax.errorbar(
                range(len(sessions)), gdata['Mean'], yerr=gdata['SE'],
                marker=marker, markersize=8, capsize=5,
                label=group, color=color, linewidth=2
            )

            # Track max y for significance markers
            for i, session in enumerate(sessions):
                y_val = gdata[gdata['Session'] == session]['Mean'].values
                se_val = gdata[gdata['Session'] == session]['SE'].values
                if len(y_val) > 0 and len(se_val) > 0:
                    max_y = y_val[0] + se_val[0]
                    if session not in max_y_per_session or max_y > max_y_per_session[session]:
                        max_y_per_session[session] = max_y

        # Add significance markers
        y_range = ax.get_ylim()[1] - ax.get_ylim()[0]
        for i, session in enumerate(sessions):
            if session in significance:
                stars = significance[session]['stars']
                if session in max_y_per_session:
                    y_pos = max_y_per_session[session] + y_range * 0.05
                else:
                    y_pos = ax.get_ylim()[1] - y_range * 0.1
                ax.text(i, y_pos, stars, ha='center', va='bottom', fontsize=12, fontweight='bold')

        ax.set_xticks(range(len(sessions)))
        ax.set_xticklabels(sessions)
        ax.set_xlabel('Session', fontsize=11)
        ax.set_ylabel('Mean PLI', fontsize=11)
        ax.set_title(model_name, fontsize=12, fontweight='bold')
        ax.legend(loc='best')
        ax.grid(True, alpha=0.3)

        plt.tight_layout()

        # Store for combined figure
        self.all_plots[model_name] = {
            'means': means,
            'sessions': sessions,
            'significance': significance
        }

        if save:
            # Save individual plot
            filename = model_name.replace(' ', '_').replace('x', '_') + '.png'
            filepath = self.output_dir / filename
            fig.savefig(str(filepath), dpi=150, bbox_inches='tight')
            plt.close(fig)
            return filepath
        else:
            return fig

    def create_combined_figure(self):
        """Create a combined grid figure with all plots."""
        if not self.all_plots:
            return None

        self._update("\nCreating combined figure...", 92)

        # Determine grid layout: rows = frequency bands, columns = networks
        networks = self.networks
        freq_bands = self.frequency_bands

        n_rows = len(freq_bands)
        n_cols = len(networks)

        fig, axes = plt.subplots(n_rows, n_cols, figsize=(4*n_cols, 3.5*n_rows))

        # Handle single row/column case
        if n_rows == 1:
            axes = axes.reshape(1, -1)
        if n_cols == 1:
            axes = axes.reshape(-1, 1)

        for i, freq in enumerate(freq_bands):
            for j, network in enumerate(networks):
                ax = axes[i, j]
                model_name = f"{freq} x {network}"

                if model_name in self.all_plots:
                    plot_data = self.all_plots[model_name]
                    means = plot_data['means']
                    sessions = plot_data['sessions']
                    significance = plot_data.get('significance', {})

                    # Get dynamic colors and markers for groups
                    unique_groups = means['Group'].unique()
                    colors = self._get_group_colors(unique_groups)
                    markers = self._get_group_markers(unique_groups)

                    # Track max y values for significance marker placement
                    max_y_per_session = {}

                    for group in sorted(means['Group'].unique()):
                        gdata = means[means['Group'] == group]
                        gdata = gdata.set_index('Session').reindex(sessions).reset_index()

                        color = colors[group]
                        marker = markers[group]

                        ax.errorbar(
                            range(len(sessions)), gdata['Mean'], yerr=gdata['SE'],
                            marker=marker, markersize=6, capsize=4,
                            label=group, color=color, linewidth=1.5
                        )

                        # Track max y for significance markers
                        for idx, session in enumerate(sessions):
                            y_val = gdata[gdata['Session'] == session]['Mean'].values
                            se_val = gdata[gdata['Session'] == session]['SE'].values
                            if len(y_val) > 0 and len(se_val) > 0:
                                max_y = y_val[0] + se_val[0]
                                if session not in max_y_per_session or max_y > max_y_per_session[session]:
                                    max_y_per_session[session] = max_y

                    # Add significance markers
                    y_range = ax.get_ylim()[1] - ax.get_ylim()[0]
                    for idx, session in enumerate(sessions):
                        if session in significance:
                            stars = significance[session]['stars']
                            if session in max_y_per_session:
                                y_pos = max_y_per_session[session] + y_range * 0.05
                            else:
                                y_pos = ax.get_ylim()[1] - y_range * 0.1
                            ax.text(idx, y_pos, stars, ha='center', va='bottom', fontsize=10, fontweight='bold')

                    ax.set_xticks(range(len(sessions)))
                    ax.set_xticklabels(sessions, fontsize=9)
                    ax.set_title(model_name, fontsize=10, fontweight='bold')
                    ax.grid(True, alpha=0.3)

                    # Only show y-label on leftmost column
                    if j == 0:
                        ax.set_ylabel('Mean PLI', fontsize=9)

                    # Only show x-label on bottom row
                    if i == n_rows - 1:
                        ax.set_xlabel('Session', fontsize=9)

                    # Only show legend on first plot
                    if i == 0 and j == 0:
                        ax.legend(loc='best', fontsize=8)
                else:
                    ax.text(0.5, 0.5, 'No Data', ha='center', va='center', transform=ax.transAxes)
                    ax.set_title(model_name, fontsize=10)

        plt.tight_layout()

        # Save combined figure
        filepath = self.output_dir / 'combined_results.png'
        fig.savefig(str(filepath), dpi=200, bbox_inches='tight')
        plt.close(fig)

        self._update(f"Combined figure saved: {filepath.name}", 95)
        return filepath

    def save_statistics_excel(self, all_contrasts):
        """Save all statistics to Excel file."""
        self._update("\nSaving statistics to Excel...", 90)

        excel_path = self.output_dir / 'analysis_statistics.xlsx'

        with pd.ExcelWriter(str(excel_path), engine='openpyxl') as writer:
            # Sheet 1: Model Fixed Effects
            if self.all_stats:
                stats_df = pd.DataFrame(self.all_stats)
                stats_df.to_excel(writer, sheet_name='Model_Effects', index=False)

            # Sheet 2: Contrasts
            if all_contrasts:
                contrasts_df = pd.DataFrame(all_contrasts)
                contrasts_df.to_excel(writer, sheet_name='Contrasts', index=False)

            # Sheet 3: Group Means Summary
            means_rows = []
            for model_name, result in self.results.items():
                if result and 'data' in result:
                    means = self.compute_means(result['data'])
                    for _, row in means.iterrows():
                        means_rows.append({
                            'Model': model_name,
                            'Group': row['Group'],
                            'Session': row['Session'],
                            'Mean': row['Mean'],
                            'SD': row['SD'],
                            'SE': row['SE'],
                            'N': row['N']
                        })

            if means_rows:
                means_df = pd.DataFrame(means_rows)
                means_df.to_excel(writer, sheet_name='Group_Means', index=False)

        self._update(f"Statistics saved: {excel_path.name}", 92)
        return excel_path

    def run(self):
        """Run complete analysis."""
        self._update("=" * 50, 0)
        self._update("NETWORK LEVEL ANALYSIS", 0)
        self._update("=" * 50, 0)

        # Create output directory
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._update(f"Output: {self.output_dir}", 5)

        # Load data
        self.load_data()

        # Run analysis for each combination
        all_results = []
        all_contrasts = []
        total = len(self.networks) * len(self.frequency_bands)
        current = 0

        for network in self.networks:
            for freq in self.frequency_bands:
                current += 1
                pct = 20 + int((current / total) * 60)

                model_name = f"{freq} x {network}"
                self._update(f"\n{'=' * 40}", pct)
                self._update(f"Model: {model_name} ({current}/{total})", pct)
                self._update('=' * 40, pct)

                result = self.fit_model(network, freq)

                if result is None:
                    self._update("  No data for this combination", pct)
                    continue

                subset = result['data']

                # Compute means
                means = self.compute_means(subset)
                self._update("\nGroup Means:", pct)
                for _, row in means.iterrows():
                    self._update(f"  {row['Group']:8} {row['Session']:8}: {row['Mean']:.4f} +/- {row['SE']:.4f}", pct)

                # Run contrasts
                contrasts, contrast_lines, contrast_stats = self.run_contrasts(
                    subset, model_name, network, freq
                )
                all_contrasts.extend(contrast_stats)

                self._update("\nContrasts:", pct)
                for line in contrast_lines:
                    self._update(line, pct)

                # Create plot with significance markers from model
                # Use model-based significance (from mixed model coefficients) for consistency with stats table
                model_significance = result.get('model_significance', {})
                # Merge with contrast significance (t-test based), preferring model significance
                contrast_significance = contrasts.get('significance', {})
                significance = {**contrast_significance, **model_significance}  # Model takes precedence
                plot_path = self.create_plot(subset, model_name, significance=significance)
                self._update(f"\nPlot saved: {plot_path.name}", pct)

                all_results.append({
                    'model': model_name,
                    'means': means,
                    'contrasts': contrasts
                })

                self.results[model_name] = result

        # Save statistics to Excel
        self.save_statistics_excel(all_contrasts)

        # Create combined figure
        self.create_combined_figure()

        # Save summary CSV
        self._save_summary(all_results)

        self._update("\n" + "=" * 50, 100)
        self._update("ANALYSIS COMPLETE", 100)
        self._update(f"Results saved to: {self.output_dir}", 100)
        self._update("=" * 50, 100)

        return all_results

    def _save_summary(self, results):
        """Save results summary to CSV."""
        rows = []
        for r in results:
            for _, row in r['means'].iterrows():
                rows.append({
                    'Model': r['model'],
                    'Group': row['Group'],
                    'Session': row['Session'],
                    'Mean': row['Mean'],
                    'SE': row['SE'],
                    'N': row['N']
                })

        if rows:
            df = pd.DataFrame(rows)
            csv_path = self.output_dir / 'summary.csv'
            df.to_csv(str(csv_path), index=False)
            self._update(f"Summary CSV saved: {csv_path.name}", 97)


def main():
    """Command-line entry point."""
    import argparse

    parser = argparse.ArgumentParser(description='Network Level PLI Analysis')
    parser.add_argument('--input', '-i', default='PLI-UK-Both-Groups-UP3.xlsx',
                        help='Input Excel file')
    parser.add_argument('--output', '-o', default='analysis_output',
                        help='Output directory')
    parser.add_argument('--no-baseline', action='store_true',
                        help='Disable baseline adjustment')
    parser.add_argument('--groups', nargs='+', default=None,
                        help='Groups to include (default: auto-detect from data)')
    parser.add_argument('--sessions', nargs='+', default=None,
                        help='Sessions to include (default: auto-detect from data)')
    parser.add_argument('--networks', nargs='+', default=None,
                        help='Networks to analyze (default: auto-detect from data)')
    parser.add_argument('--bands', nargs='+', default=None,
                        help='Frequency bands (default: auto-detect from data)')

    args = parser.parse_args()

    analysis = NetworkAnalysis(
        input_file=args.input,
        output_dir=args.output,
        adjust_baseline=not args.no_baseline,
        groups=args.groups,
        sessions=args.sessions,
        networks=args.networks,
        frequency_bands=args.bands
    )

    analysis.run()


if __name__ == '__main__':
    main()
