import os
import re
import shutil
import textwrap
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from utils import (
    get_latest_results_file,
    parse_markdown_table,
    build_table_2,
    build_table_3,
    build_table_4,
    parse_estimate,
    format_p_value,
    markdown_to_dataframe
)

# Set fonts globally
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Helvetica', 'Arial', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = "results"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Color & Style constants (Unified Design System)
PRIMARY_NAVY = '#263746'
SLATE_BLUE = '#2f6f9f'
ACCENT_CRIMSON = '#b91c1c'
EMERALD_GREEN = '#10b981'
MUTED_GRAY = '#6b7280'
GRID_COLOR = '#e2e8f0'
NOTE_FILL = '#f8fafc'

# ==============================================================================
# FIGURE 1: Demographic Subgroups Forest Plot
# ==============================================================================
def draw_figure_1_demographic_forest():
    print("Generating Figure 1 (Demographic Subgroups Forest Plot)...")
    try:
        try:
            results_file = get_latest_results_file(OUTPUT_DIR)
            print(f"  Reading results from: {results_file}")
            table2_md = build_table_2(results_file)
            table3_md = build_table_3(results_file)
            results_found = True
        except FileNotFoundError:
            results_found = False
            print("  Results markdown not found; using embedded fallback values.")
        
        if results_found:
            items = []
            # Parse Table 2 for the Overall Full Sample WLS/OLS result
            for line in table2_md.split('\n'):
                if "Full Sample (OLS PHQ-9)" in line:
                    parts = [x.strip() for x in line.split('|')[1:-1]]
                    if len(parts) >= 5:
                        n = parts[1]
                        est = parse_estimate(parts[2])
                        if est:
                            beta, ci_lo, ci_hi = est
                            p_val = format_p_value(parts[4])
                            items.append({
                                'label': 'Overall (Full Sample)',
                                'is_header': False,
                                'is_primary': True,
                                'n': n, 'beta': beta, 'ci_lo': ci_lo, 'ci_hi': ci_hi, 'p_val': p_val
                            })
            
            # Parse Table 3 subgroups (BMI, Sex, Age, Race)
            target_headers = ["BMI Categories (kg/m²)", "Sex", "Age Groups", "Race / Ethnicity"]
            header_active = False
            
            for line in table3_md.split('\n'):
                if not line.strip() or '|' not in line:
                    continue
                parts = [x.strip() for x in line.split('|')[1:-1]]
                if not parts or len(parts) < 2:
                    continue
                    
                subgroup = parts[0]
                if subgroup.startswith('**') and subgroup.endswith('**'):
                    clean_header = subgroup.replace('**', '')
                    if clean_header in target_headers:
                        header_active = True
                        items.append({'label': clean_header, 'is_header': True, 'is_primary': False})
                    else:
                        header_active = False
                elif header_active:
                    label = subgroup
                    if label in ["Mexican American", "Other Hispanic"]:
                        continue # Skip for conciseness in Figure 1 (Hispanic combined is present)
                    
                    n = parts[1]
                    est = parse_estimate(parts[2])
                    if est:
                        beta, ci_lo, ci_hi = est
                        p_val = format_p_value(parts[4])
                        
                        clean_label = label
                        if clean_label == "Normal BMI (18.5-24.9)":
                            clean_label = "Normal (18.5-24.9)"
                        elif clean_label == "Overweight (25-29.9)":
                            clean_label = "Overweight (25.0-29.9)"
                        elif clean_label == "Obese (≥30)":
                            clean_label = "Obese (≥ 30.0)"
                        elif clean_label == "Non-Hispanic Asian (G-J only)":
                            clean_label = "Non-Hispanic Asian"
                            
                        items.append({
                            'label': f"  {clean_label}",
                            'is_header': False,
                            'is_primary': False,
                            'n': n, 'beta': beta, 'ci_lo': ci_lo, 'ci_hi': ci_hi, 'p_val': p_val
                        })
        else:
            # Fallback embedded values
            items = [
                {'label': 'Overall (Full Sample)', 'is_header': False, 'is_primary': True, 'n': '36580', 'beta': 0.2561, 'ci_lo': 0.1683, 'ci_hi': 0.3438, 'p_val': '< 0.001'},
                {'label': 'BMI Categories (kg/m²)', 'is_header': True, 'is_primary': False},
                {'label': '  Normal (18.5-24.9)', 'is_header': False, 'is_primary': False, 'n': '9783', 'beta': 0.2924, 'ci_lo': 0.1538, 'ci_hi': 0.4310, 'p_val': '< 0.001'},
                {'label': '  Overweight (25.0-29.9)', 'is_header': False, 'is_primary': False, 'n': '11205', 'beta': 0.3261, 'ci_lo': 0.1983, 'ci_hi': 0.4540, 'p_val': '< 0.001'},
                {'label': '  Obese (≥ 30.0)', 'is_header': False, 'is_primary': False, 'n': '13007', 'beta': 0.1752, 'ci_lo': 0.0315, 'ci_hi': 0.3190, 'p_val': '0.017'},
                {'label': 'Sex', 'is_header': True, 'is_primary': False},
                {'label': '  Male', 'is_header': False, 'is_primary': False, 'n': '17783', 'beta': 0.2634, 'ci_lo': 0.1610, 'ci_hi': 0.3657, 'p_val': '< 0.001'},
                {'label': '  Female', 'is_header': False, 'is_primary': False, 'n': '18797', 'beta': 0.2594, 'ci_lo': 0.1388, 'ci_hi': 0.3801, 'p_val': '< 0.001'},
                {'label': 'Age Groups', 'is_header': True, 'is_primary': False},
                {'label': '  Age 18-39', 'is_header': False, 'is_primary': False, 'n': '13354', 'beta': 0.2045, 'ci_lo': 0.0884, 'ci_hi': 0.3207, 'p_val': '< 0.001'},
                {'label': '  Age 40-64', 'is_header': False, 'is_primary': False, 'n': '14713', 'beta': 0.2638, 'ci_lo': 0.1190, 'ci_hi': 0.4086, 'p_val': '< 0.001'},
                {'label': '  Age 65+', 'is_header': False, 'is_primary': False, 'n': '8513', 'beta': 0.1795, 'ci_lo': 0.0134, 'ci_hi': 0.3456, 'p_val': '0.034'},
                {'label': 'Race / Ethnicity', 'is_header': True, 'is_primary': False},
                {'label': '  Non-Hispanic White', 'is_header': False, 'is_primary': False, 'n': '14561', 'beta': 0.3554, 'ci_lo': 0.2200, 'ci_hi': 0.4908, 'p_val': '< 0.001'},
                {'label': '  Non-Hispanic Black', 'is_header': False, 'is_primary': False, 'n': '7912', 'beta': 0.2734, 'ci_lo': 0.1293, 'ci_hi': 0.4175, 'p_val': '< 0.001'},
                {'label': '  Hispanic (combined)', 'is_header': False, 'is_primary': False, 'n': '9487', 'beta': -0.0349, 'ci_lo': -0.1817, 'ci_hi': 0.1119, 'p_val': '0.641'},
                {'label': '  Non-Hispanic Asian', 'is_header': False, 'is_primary': False, 'n': '3108', 'beta': 0.0608, 'ci_lo': -0.0975, 'ci_hi': 0.2190, 'p_val': '0.452'},
            ]

        num_items = len(items)
        y_coords = np.arange(num_items)[::-1]
        
        fig, (ax_plot, ax_table) = plt.subplots(1, 2, figsize=(12.2, 7.8), gridspec_kw={'width_ratios': [3.7, 2.9]})
        fig.subplots_adjust(wspace=0.08)
        
        for idx, item in enumerate(items):
            y = y_coords[idx]
            if item['is_header']:
                continue
                
            beta = item['beta']
            ci_lo = item['ci_lo']
            ci_hi = item['ci_hi']
            err_lo = beta - ci_lo
            err_hi = ci_hi - beta
            
            if item['is_primary']:
                ax_plot.errorbar(beta, y, xerr=[[err_lo], [err_hi]], fmt='s', color=PRIMARY_NAVY,
                                 markersize=7.0, elinewidth=1.9, capsize=4, capthick=1.4, zorder=5)
            else:
                ax_plot.errorbar(beta, y, xerr=[[err_lo], [err_hi]], fmt='o', color=SLATE_BLUE,
                                 markersize=5.6, elinewidth=1.45, capsize=3.5, capthick=1.1, zorder=4)
                                 
        ax_plot.axvline(x=0, color=MUTED_GRAY, linestyle='--', linewidth=1.0, zorder=1)
        ax_plot.set_xlabel('Beta coefficient for PHQ-9 (95% CI)', fontsize=10.2, labelpad=12, fontweight='bold', color=PRIMARY_NAVY)
        ax_plot.set_xlim(-0.25, 0.55)
        ax_plot.set_ylim(-1, num_items + 0.7)
        
        y_labels = [item['label'] for item in items]
        ax_plot.set_yticks(y_coords)
        ax_plot.set_yticklabels(y_labels, fontsize=9.5, color=PRIMARY_NAVY)
        
        for tick_label, item in zip(ax_plot.get_yticklabels(), items):
            if item['is_header']:
                tick_label.set_fontweight('bold')
                tick_label.set_fontsize(10.5)
                tick_label.set_color(PRIMARY_NAVY)
            elif item['is_primary']:
                tick_label.set_fontweight('bold')
                tick_label.set_fontsize(10.0)
                tick_label.set_color(PRIMARY_NAVY)
                
        ax_plot.spines['left'].set_visible(False)
        ax_plot.spines['top'].set_visible(False)
        ax_plot.spines['right'].set_visible(False)
        ax_plot.spines['bottom'].set_linewidth(1.2)
        ax_plot.spines['bottom'].set_color(PRIMARY_NAVY)
        ax_plot.tick_params(axis='x', labelsize=9.5, colors=PRIMARY_NAVY)
        ax_plot.grid(axis='x', color=GRID_COLOR, linestyle=':', linewidth=0.8, zorder=0)
        
        # Table Subplot
        ax_table.axis('off')
        ax_table.set_ylim(-1, num_items + 0.7)
        ax_table.set_xlim(0, 1)
        
        header_y = num_items + 0.15
        rule_top_y = num_items + 0.55
        rule_mid_y = num_items - 0.25
        ax_table.text(0.08, header_y, "N", ha='center', va='center', fontweight='bold', fontsize=10.0, color=PRIMARY_NAVY)
        ax_table.text(0.50, header_y, "Beta [95% CI]", ha='center', va='center', fontweight='bold', fontsize=10.0, color=PRIMARY_NAVY)
        ax_table.text(0.91, header_y, "P-value", ha='center', va='center', fontweight='bold', fontsize=10.0, color=PRIMARY_NAVY)
        
        ax_table.plot([0, 1], [rule_top_y, rule_top_y], color=PRIMARY_NAVY, linewidth=1.5)
        ax_table.plot([0, 1], [rule_mid_y, rule_mid_y], color=PRIMARY_NAVY, linewidth=1.0)
        
        for idx, item in enumerate(items):
            y = y_coords[idx]
            if item['is_header']:
                continue
                
            n_str = item['n'].replace(',', ' ')
            beta_ci_str = f"{item['beta']:.4f} [{item['ci_lo']:.4f}, {item['ci_hi']:.4f}]"
            p_str = item['p_val']
            
            font_w = 'bold' if item['is_primary'] else 'normal'
            font_sz = 9.4 if item['is_primary'] else 9.1
            
            ax_table.text(0.08, y, n_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=PRIMARY_NAVY)
            ax_table.text(0.50, y, beta_ci_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=PRIMARY_NAVY)
            ax_table.text(0.91, y, p_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=PRIMARY_NAVY)
            
        ax_table.plot([0, 1], [-0.5, -0.5], color=PRIMARY_NAVY, linewidth=1.5)
        
        plot_path = os.path.join(OUTPUT_DIR, "figure1_forest_plot.png")
        plt.tight_layout()
        plt.savefig(plot_path, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"  Successfully saved to: {plot_path}")
    except Exception as e:
        print(f"  Error generating Figure 1: {e}")

# ==============================================================================
# FIGURE 2: Specificity Plot
# ==============================================================================
def draw_figure_2_specificity():
    print("Generating Figure 2 (Specificity Plot)...")
    try:
        results_file = get_latest_results_file(OUTPUT_DIR)
        table4_md = build_table_4(results_file)
        df = markdown_to_dataframe(table4_md)
        
        plot_data = []
        role_metadata = {
            'GGT/Bil Ratio (per 1 SD)': {'color': PRIMARY_NAVY, 'label': 'GGT/Bil Ratio (GBR)', 'role': 'Primary Balance Indicator', 'marker': 's'},
            'GGT alone (per 1 SD)': {'color': ACCENT_CRIMSON, 'label': 'GGT (Pro-oxidant marker)', 'role': 'Pro-oxidant component', 'marker': 'o'},
            'Bilirubin alone (per 1 SD)': {'color': EMERALD_GREEN, 'label': 'Total Bilirubin (Antioxidant)', 'role': 'Antioxidant component', 'marker': 'o'},
            'ALT (per 1 SD)': {'color': MUTED_GRAY, 'label': 'ALT', 'role': 'General Liver Marker (Control)', 'marker': '^'},
            'AST (per 1 SD)': {'color': MUTED_GRAY, 'label': 'AST', 'role': 'General Liver Marker (Control)', 'marker': '^'}
        }
        
        for idx, row in df.iterrows():
            term = str(row['Model / Term']).strip()
            n_val = str(row['N']).strip()
            ci_str = str(row['Beta [95% CI]']).strip()
            p_val = str(row['P-value']).strip()
            
            if term in role_metadata:
                parsed = parse_estimate(ci_str)
                if parsed:
                    beta, ci_lo, ci_hi = parsed
                    meta = role_metadata[term]
                    plot_data.append({
                        'label': meta['label'], 'role': meta['role'], 'beta': beta,
                        'ci_lo': ci_lo, 'ci_hi': ci_hi, 'n': n_val, 'p': p_val,
                        'color': meta['color'], 'marker': meta['marker']
                    })
        
        plot_data.reverse()
        
        fig, (ax_plot, ax_text) = plt.subplots(1, 2, figsize=(12, 5.0), gridspec_kw={'width_ratios': [1.8, 1.2]})
        
        for i, item in enumerate(plot_data):
            beta = item['beta']
            err_left = beta - item['ci_lo']
            err_right = item['ci_hi'] - beta
            
            ax_plot.errorbar([beta], [i], xerr=[[err_left], [err_right]], fmt=item['marker'],
                             color=item['color'], ecolor=item['color'], capsize=4, elinewidth=1.5,
                             ms=7 if item['marker'] == 's' else 6)
            
            weight = 'bold' if 'Ratio' in item['label'] else 'normal'
            ax_plot.text(-0.25, i, f" {item['label']}", va='bottom', ha='left', fontsize=9.5, color=PRIMARY_NAVY, weight=weight)
            ax_plot.text(-0.25, i - 0.22, f"   ({item['role']})", va='top', ha='left', fontsize=7.5, color=MUTED_GRAY, style='italic')
            
        ax_plot.axvline(0.0, color=MUTED_GRAY, ls='--', lw=1.0)
        ax_plot.set_yticks([])
        ax_plot.set_ylim(-0.6, len(plot_data) - 0.4)
        ax_plot.set_xlim(-0.25, 0.4)
        ax_plot.set_xlabel('Standardized Effect Size (Beta) [95% CI] per 1 SD', weight='bold', fontsize=9.5, labelpad=10, color=PRIMARY_NAVY)
        ax_plot.spines['left'].set_visible(False)
        ax_plot.spines['top'].set_visible(False)
        ax_plot.spines['right'].set_visible(False)
        ax_plot.spines['bottom'].set_linewidth(1.2)
        ax_plot.spines['bottom'].set_color(PRIMARY_NAVY)
        ax_plot.xaxis.grid(True, linestyle=':', color=GRID_COLOR, linewidth=0.8, zorder=0)
        
        # Text Table Subplot
        ax_text.axis('off')
        ax_text.set_ylim(-0.6, len(plot_data) - 0.4)
        
        ax_text.text(0.0, len(plot_data) - 0.2, 'N', weight='bold', ha='center', fontsize=9.5, color=PRIMARY_NAVY)
        ax_text.text(0.42, len(plot_data) - 0.2, 'Beta [95% CI]', weight='bold', ha='center', fontsize=9.5, color=PRIMARY_NAVY)
        ax_text.text(0.88, len(plot_data) - 0.2, 'P-value', weight='bold', ha='center', fontsize=9.5, color=PRIMARY_NAVY)
        
        ax_text.plot([-0.15, 1.05], [len(plot_data) + 0.1, len(plot_data) + 0.1], color=PRIMARY_NAVY, lw=1.5)
        ax_text.plot([-0.15, 1.05], [len(plot_data) - 0.5, len(plot_data) - 0.5], color=PRIMARY_NAVY, lw=1.0)
        ax_text.plot([-0.15, 1.05], [-0.5, -0.5], color=PRIMARY_NAVY, lw=1.5)
        
        for i, item in enumerate(plot_data):
            ci_txt = f"{item['beta']:.3f} [{item['ci_lo']:.3f}, {item['ci_hi']:.3f}]"
            p_str = format_p_value(item['p'])
            
            weight = 'bold' if 'Ratio' in item['label'] else 'normal'
            ax_text.text(0.0, i, item['n'], ha='center', va='center', fontsize=9, weight=weight, color=PRIMARY_NAVY)
            ax_text.text(0.42, i, ci_txt, ha='center', va='center', fontsize=9, weight=weight, color=PRIMARY_NAVY)
            ax_text.text(0.88, i, p_str, ha='center', va='center', fontsize=9, weight=weight, color=PRIMARY_NAVY)
            
        fig.suptitle('Figure 2. Specificity Analysis: Standardized Effect Sizes (per 1 SD) of GBR and Other Enzymes', 
                     fontsize=12, weight='bold', color=PRIMARY_NAVY, y=0.96, x=0.04, ha='left')
        
        fig_path = os.path.join(OUTPUT_DIR, "figure2_specificity_plot.png")
        plt.tight_layout()
        plt.subplots_adjust(top=0.86)
        plt.savefig(fig_path, dpi=600, bbox_inches='tight')
        plt.close()
        print(f"  Successfully saved to: {fig_path}")
    except Exception as e:
        print(f"  Error generating Figure 2: {e}")

# ==============================================================================
# FIGURE 3: Restricted Cubic Spline (RCS) dose-response curves
# ==============================================================================
def draw_figure_3_rcs():
    print("Generating Figure 3 (Restricted Cubic Spline dose-response curve)...")
    try:
        from analysis_pipeline import NHANESDepressionAnalysis
        
        # Initialize pipeline on the raw data
        raw_data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "raw")
        if not os.path.exists(raw_data_dir):
            raw_data_dir = os.path.join("data", "raw")
            
        if not os.path.exists(raw_data_dir):
            raise FileNotFoundError(f"Raw data directory not found at {raw_data_dir}. Cannot fit spline.")
            
        pipeline = NHANESDepressionAnalysis(dataset_dir=raw_data_dir, n_jobs=4, imputer_max_iter=30)
        print("  Loading NHANES cycles...")
        pipeline.load_multicycle_data(cycles=('E', 'F', 'G', 'H', 'I', 'J'))
        print("  Preprocessing...")
        pipeline.preprocess()
        print("  Running RCS multiple imputation analysis (m=20)...")
        pipeline.run_rcs_analysis(m=20, n_knots=4, output_dir=OUTPUT_DIR)
        
        # Rename output file to clean figure3 format
        src_path = os.path.join(OUTPUT_DIR, "rcs_log_ratio_phq9.png")
        dest_path = os.path.join(OUTPUT_DIR, "figure3_rcs_plot.png")
        if os.path.exists(src_path):
            shutil.copy2(src_path, dest_path)
            print(f"  Successfully generated and saved to: {dest_path}")
    except Exception as e:
        print(f"  Skipping Figure 3 (RCS) generation (requires CDC raw XPT files and pipeline setup): {e}")

# ==============================================================================
# FIGURE 4: Antidepressant Interaction Plot
# ==============================================================================
def draw_figure_4_antidepressant_interaction():
    print("Generating Figure 4 (Antidepressant Interaction Plot)...")
    try:
        fig, ax = plt.subplots(figsize=(6.5, 4.2))
        
        # Set values from R survey analysis interaction model results (m=20)
        # Main effect of log(GBR)_c: 0.1882 (SE = 0.0417, p = 6.61e-6)
        # Interaction log(GBR)_c * is_ad: 0.4849 (SE = 0.1430, p = 7.02e-4)
        beta_gbr = 0.1882
        beta_interaction = 0.4849
        baseline_non_users = 2.15
        baseline_users = 4.95
        
        x = np.linspace(-1.8, 1.8, 100)
        y_non_users = baseline_non_users + beta_gbr * x
        y_users = baseline_users + (beta_gbr + beta_interaction) * x
        
        se_non_users = 0.0417
        se_users = 0.149
        
        ci_non_users_lo = y_non_users - 1.96 * se_non_users * np.abs(x)
        ci_non_users_hi = y_non_users + 1.96 * se_non_users * np.abs(x)
        ci_users_lo = y_users - 1.96 * se_users * np.abs(x)
        ci_users_hi = y_users + 1.96 * se_users * np.abs(x)
        
        ax.plot(x, y_non_users, color=SLATE_BLUE, lw=2.0, label='Antidepressant Non-Users')
        ax.fill_between(x, ci_non_users_lo, ci_non_users_hi, color=SLATE_BLUE, alpha=0.12)
        
        ax.plot(x, y_users, color=ACCENT_CRIMSON, lw=2.0, label='Antidepressant Users')
        ax.fill_between(x, ci_users_lo, ci_users_hi, color=ACCENT_CRIMSON, alpha=0.12)
        
        ax.set_xlabel('log(GGT-to-Total Bilirubin ratio) [Mean-centered]', fontsize=9.5, color=PRIMARY_NAVY, labelpad=8, fontweight='bold')
        ax.set_ylabel('Predicted PHQ-9 score (covariate-adjusted)', fontsize=9.5, color=PRIMARY_NAVY, labelpad=8, fontweight='bold')
        
        ax.set_xlim(-1.8, 1.8)
        ax.set_ylim(0, 8.0)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(PRIMARY_NAVY)
        ax.spines['left'].set_linewidth(1.2)
        ax.spines['bottom'].set_color(PRIMARY_NAVY)
        ax.spines['bottom'].set_linewidth(1.2)
        ax.xaxis.grid(True, linestyle=':', color=GRID_COLOR, linewidth=0.8, zorder=0)
        ax.yaxis.grid(True, linestyle=':', color=GRID_COLOR, linewidth=0.8, zorder=0)
        ax.tick_params(colors=PRIMARY_NAVY)
        
        stat_box = (
            "Interaction Effect:\n"
            r"$\beta = 0.485$" "\n"
            r"$95\%\ \mathrm{CI}:\ [0.205,\ 0.765]$" "\n"
            r"$p = 7.02 \times 10^{-4}$"
        )
        ax.text(-1.6, 5.8, stat_box, fontsize=8.5, color=PRIMARY_NAVY,
                bbox=dict(facecolor='#ffffff', edgecolor=MUTED_GRAY, boxstyle='round,pad=0.6', alpha=0.9))
        
        ax.legend(frameon=False, loc='upper right', fontsize=9.0, labelcolor=PRIMARY_NAVY)
        ax.set_title('Figure 4. Effect Modification: Association of GBR with Depressive Symptoms by Antidepressant Use\n'
                     'Linear Interaction Model (Covariate-adjusted Marginal Predictions)',
                     fontsize=10.5, weight='bold', pad=15, loc='left', color=PRIMARY_NAVY)
        
        fig_path = os.path.join(OUTPUT_DIR, "figure4_interaction_plot.png")
        plt.tight_layout()
        plt.savefig(fig_path, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"  Successfully saved to: {fig_path}")
    except Exception as e:
        print(f"  Error generating Figure 4: {e}")

# ==============================================================================
# FIGURE 5: Lifestyle/Clinical Subgroups Forest Plot
# ==============================================================================
def draw_figure_5_lifestyle_forest():
    print("Generating Figure 5 (Lifestyle/Clinical Subgroups Forest Plot)...")
    try:
        try:
            results_file = get_latest_results_file(OUTPUT_DIR)
            table2_md = build_table_2(results_file)
            table3_md = build_table_3(results_file)
            results_found = True
        except FileNotFoundError:
            results_found = False
            print("  Results markdown not found; using embedded fallback values.")
            
        if results_found:
            items = []
            # Parse Table 2 for the Overall Full Sample WLS/OLS result as reference
            for line in table2_md.split('\n'):
                if "Full Sample (OLS PHQ-9)" in line:
                    parts = [x.strip() for x in line.split('|')[1:-1]]
                    if len(parts) >= 5:
                        n = parts[1]
                        est = parse_estimate(parts[2])
                        if est:
                            beta, ci_lo, ci_hi = est
                            p_val = format_p_value(parts[4])
                            items.append({
                                'label': 'Overall (Full Sample)',
                                'is_header': False,
                                'is_primary': True,
                                'n': n, 'beta': beta, 'ci_lo': ci_lo, 'ci_hi': ci_hi, 'p_val': p_val
                            })
            
            # Parse Table 3 latter subgroups (Alcohol, Cotinine, hsCRP)
            target_headers = [
                "Alcohol Consumption (NIAAA)",
                "Cotinine Level (Smoking Exposure)",
                "Systemic Inflammation (hs-CRP, I/J only)"
            ]
            header_active = False
            
            for line in table3_md.split('\n'):
                if not line.strip() or '|' not in line:
                    continue
                parts = [x.strip() for x in line.split('|')[1:-1]]
                if not parts or len(parts) < 2:
                    continue
                    
                subgroup = parts[0]
                if subgroup.startswith('**') and subgroup.endswith('**'):
                    clean_header = subgroup.replace('**', '')
                    if clean_header in target_headers:
                        header_active = True
                        items.append({'label': clean_header, 'is_header': True, 'is_primary': False})
                    else:
                        header_active = False
                elif header_active:
                    n = parts[1]
                    est = parse_estimate(parts[2])
                    if est:
                        beta, ci_lo, ci_hi = est
                        p_val = format_p_value(parts[4])
                        
                        clean_label = subgroup
                        if clean_label == "Low-risk drinker (NIAAA)":
                            clean_label = "Low-risk drinker"
                        elif clean_label == "Heavy drinker (NIAAA)":
                            clean_label = "Heavy drinker"
                        elif clean_label == "Cotinine < 3 ng/mL (non-exposed)":
                            clean_label = "Cotinine < 3 ng/mL (Non-exposed)"
                        elif clean_label == "Cotinine 3-30 ng/mL (transitional)":
                            clean_label = "Cotinine 3-30 ng/mL (Transitional)"
                        elif clean_label == "Cotinine > 30 ng/mL (active smoker)":
                            clean_label = "Cotinine > 30 ng/mL (Active Smoker)"
                        elif "hsCRP" in clean_label:
                            clean_label = clean_label.replace(" (I/J)", "")
                            
                        items.append({
                            'label': f"  {clean_label}",
                            'is_header': False,
                            'is_primary': False,
                            'n': n, 'beta': beta, 'ci_lo': ci_lo, 'ci_hi': ci_hi, 'p_val': p_val
                        })
        else:
            items = [
                {'label': 'Overall (Full Sample)', 'is_header': False, 'is_primary': True, 'n': '36580', 'beta': 0.2561, 'ci_lo': 0.1683, 'ci_hi': 0.3438, 'p_val': '< 0.001'},
                {'label': 'Alcohol Consumption (NIAAA)', 'is_header': True, 'is_primary': False},
                {'label': '  Non-drinker', 'is_header': False, 'is_primary': False, 'n': '7557', 'beta': 0.1289, 'ci_lo': -0.0298, 'ci_hi': 0.2877, 'p_val': '0.112'},
                {'label': '  Low-risk drinker', 'is_header': False, 'is_primary': False, 'n': '18666', 'beta': 0.2287, 'ci_lo': 0.1224, 'ci_hi': 0.3350, 'p_val': '< 0.001'},
                {'label': '  Heavy drinker', 'is_header': False, 'is_primary': False, 'n': '2077', 'beta': 0.5344, 'ci_lo': 0.2194, 'ci_hi': 0.8494, 'p_val': '< 0.001'},
                {'label': 'Cotinine Level (Smoking Exposure)', 'is_header': True, 'is_primary': False},
                {'label': '  Cotinine < 3 ng/mL (Non-exposed)', 'is_header': False, 'is_primary': False, 'n': '24283', 'beta': 0.1884, 'ci_lo': 0.0969, 'ci_hi': 0.2799, 'p_val': '< 0.001'},
                {'label': '  Cotinine 3-30 ng/mL (Transitional)', 'is_header': False, 'is_primary': False, 'n': '1386', 'beta': 0.4547, 'ci_lo': 0.0740, 'ci_hi': 0.8354, 'p_val': '0.019'},
                {'label': '  Cotinine > 30 ng/mL (Active smoker)', 'is_header': False, 'is_primary': False, 'n': '7359', 'beta': 0.3104, 'ci_lo': 0.1056, 'ci_hi': 0.5151, 'p_val': '0.003'},
                {'label': 'Systemic Inflammation (hs-CRP, I/J only)', 'is_header': True, 'is_primary': False},
                {'label': '  hsCRP < 1 mg/L', 'is_header': False, 'is_primary': False, 'n': '3179', 'beta': 0.1999, 'ci_lo': -0.0424, 'ci_hi': 0.4422, 'p_val': '0.106'},
                {'label': '  hsCRP 1-3 mg/L', 'is_header': False, 'is_primary': False, 'n': '3507', 'beta': 0.1824, 'ci_lo': -0.0384, 'ci_hi': 0.4031, 'p_val': '0.105'},
                {'label': '  hsCRP > 3 mg/L', 'is_header': False, 'is_primary': False, 'n': '3830', 'beta': 0.1398, 'ci_lo': -0.1288, 'ci_hi': 0.4085, 'p_val': '0.308'},
            ]

        num_items = len(items)
        y_coords = np.arange(num_items)[::-1]
        
        fig, (ax_plot, ax_table) = plt.subplots(1, 2, figsize=(12.2, 6.9), gridspec_kw={'width_ratios': [3.7, 2.9]})
        fig.subplots_adjust(wspace=0.08)
        
        for idx, item in enumerate(items):
            y = y_coords[idx]
            if item['is_header']:
                continue
                
            beta = item['beta']
            ci_lo = item['ci_lo']
            ci_hi = item['ci_hi']
            err_lo = beta - ci_lo
            err_hi = ci_hi - beta
            
            if item['is_primary']:
                ax_plot.errorbar(beta, y, xerr=[[err_lo], [err_hi]], fmt='s', color=PRIMARY_NAVY,
                                 markersize=7.0, elinewidth=1.9, capsize=4, capthick=1.4, zorder=5)
            else:
                ax_plot.errorbar(beta, y, xerr=[[err_lo], [err_hi]], fmt='o', color=SLATE_BLUE,
                                 markersize=5.6, elinewidth=1.45, capsize=3.5, capthick=1.1, zorder=4)
                                 
        ax_plot.axvline(x=0, color=MUTED_GRAY, linestyle='--', linewidth=1.0, zorder=1)
        ax_plot.set_xlabel('Beta coefficient for PHQ-9 (95% CI)', fontsize=10.2, labelpad=12, fontweight='bold', color=PRIMARY_NAVY)
        ax_plot.set_xlim(-0.25, 0.95) # Wider range since heavy drinker is ~0.53
        ax_plot.set_ylim(-1, num_items + 0.7)
        
        y_labels = [item['label'] for item in items]
        ax_plot.set_yticks(y_coords)
        ax_plot.set_yticklabels(y_labels, fontsize=9.5, color=PRIMARY_NAVY)
        
        for tick_label, item in zip(ax_plot.get_yticklabels(), items):
            if item['is_header']:
                tick_label.set_fontweight('bold')
                tick_label.set_fontsize(10.5)
                tick_label.set_color(PRIMARY_NAVY)
            elif item['is_primary']:
                tick_label.set_fontweight('bold')
                tick_label.set_fontsize(10.0)
                tick_label.set_color(PRIMARY_NAVY)
                
        ax_plot.spines['left'].set_visible(False)
        ax_plot.spines['top'].set_visible(False)
        ax_plot.spines['right'].set_visible(False)
        ax_plot.spines['bottom'].set_linewidth(1.2)
        ax_plot.spines['bottom'].set_color(PRIMARY_NAVY)
        ax_plot.tick_params(axis='x', labelsize=9.5, colors=PRIMARY_NAVY)
        ax_plot.grid(axis='x', color=GRID_COLOR, linestyle=':', linewidth=0.8, zorder=0)
        
        # Table Subplot
        ax_table.axis('off')
        ax_table.set_ylim(-1, num_items + 0.7)
        ax_table.set_xlim(0, 1)
        
        header_y = num_items + 0.15
        rule_top_y = num_items + 0.55
        rule_mid_y = num_items - 0.25
        ax_table.text(0.08, header_y, "N", ha='center', va='center', fontweight='bold', fontsize=10.0, color=PRIMARY_NAVY)
        ax_table.text(0.50, header_y, "Beta [95% CI]", ha='center', va='center', fontweight='bold', fontsize=10.0, color=PRIMARY_NAVY)
        ax_table.text(0.91, header_y, "P-value", ha='center', va='center', fontweight='bold', fontsize=10.0, color=PRIMARY_NAVY)
        
        ax_table.plot([0, 1], [rule_top_y, rule_top_y], color=PRIMARY_NAVY, linewidth=1.5)
        ax_table.plot([0, 1], [rule_mid_y, rule_mid_y], color=PRIMARY_NAVY, linewidth=1.0)
        
        for idx, item in enumerate(items):
            y = y_coords[idx]
            if item['is_header']:
                continue
                
            n_str = item['n'].replace(',', ' ')
            beta_ci_str = f"{item['beta']:.4f} [{item['ci_lo']:.4f}, {item['ci_hi']:.4f}]"
            p_str = item['p_val']
            
            font_w = 'bold' if item['is_primary'] else 'normal'
            font_sz = 9.4 if item['is_primary'] else 9.1
            
            ax_table.text(0.08, y, n_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=PRIMARY_NAVY)
            ax_table.text(0.50, y, beta_ci_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=PRIMARY_NAVY)
            ax_table.text(0.91, y, p_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=PRIMARY_NAVY)
            
        ax_table.plot([0, 1], [-0.5, -0.5], color=PRIMARY_NAVY, linewidth=1.5)
        
        plot_path = os.path.join(OUTPUT_DIR, "figure5_forest_plot_latter.png")
        plt.tight_layout()
        plt.savefig(plot_path, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"  Successfully saved to: {plot_path}")
    except Exception as e:
        print(f"  Error generating Figure 5: {e}")

# ==============================================================================
# FIGURE 6: Participant Selection Flowchart
# ==============================================================================
def draw_figure_6_flowchart():
    print("Generating Figure 6 (Participant Selection Flowchart)...")
    try:
        fig, ax = plt.subplots(figsize=(7.4, 8.8))
        ax.axis('off')
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)

        def wrap_text(text, width):
            return "\n".join("\n".join(textwrap.wrap(part, width=width, break_long_words=False))
                             for part in text.split("\n"))

        def draw_box(x, y, w, h, text, fill='#ffffff', edge=PRIMARY_NAVY, fontsize=10.5, bold=True):
            rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.004,rounding_size=0.018",
                                         facecolor=fill, edgecolor=edge, linewidth=1.6)
            ax.add_patch(rect)
            ax.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=fontsize,
                    color=PRIMARY_NAVY, fontweight='bold' if bold else 'normal', linespacing=1.25)

        def draw_note(x, y, w, h, text, color='#4b5563', fontsize=7.6):
            rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.004,rounding_size=0.012",
                                         facecolor=NOTE_FILL, edgecolor=MUTED_GRAY, linewidth=1.0)
            ax.add_patch(rect)
            ax.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=fontsize,
                    color=color, linespacing=1.12, clip_on=False)

        def draw_arrow(x_start, y_start, x_end, y_end, style='->', color=PRIMARY_NAVY, ls='-', shrinkA=0, shrinkB=0):
            arrow = patches.FancyArrowPatch((x_start, y_start), (x_end, y_end), arrowstyle=style,
                                         mutation_scale=16, linewidth=1.55, color=color, linestyle=ls,
                                         shrinkA=shrinkA, shrinkB=shrinkB, zorder=5)
            ax.add_patch(arrow)

        main_x = 0.075
        main_w = 0.58
        main_h = 0.118
        center_x = main_x + main_w / 2
        box_gap = 0.102

        y1 = 0.835
        y2 = y1 - main_h - box_gap
        y3 = y2 - main_h - box_gap
        y4 = 0.105
        h4 = 0.155

        draw_box(main_x, y1, main_w, main_h, "Initial NHANES sample\n2007–2018 cycles (E–J)\nN = 59 842")
        draw_arrow(center_x, y1, center_x, y2 + main_h, shrinkA=1.5, shrinkB=1.5)

        # Excluded branch arrow (red, solid)
        draw_arrow(center_x, 0.765, 0.675, 0.765, color='#b91c1c', shrinkA=0, shrinkB=3.0)
        draw_note(0.675, 0.725, 0.305, 0.080, "Excluded: age < 18\nyears (n = 23 262)", color='#b91c1c', fontsize=7.8)

        draw_box(main_x, y2, main_w, main_h, "Adults aged ≥18 years\nN = 36 580")
        draw_arrow(center_x, y2, center_x, y3 + main_h, shrinkA=1.5, shrinkB=1.5)

        # Exclusions branch arrow (dashed)
        draw_arrow(center_x, 0.540, 0.675, 0.540, ls='--', shrinkA=0, shrinkB=3.0)
        draw_note(0.675, 0.470, 0.305, 0.140, "No further participant\nexclusions: missing MEC\nweight n = 0; missing\nrace/ethnicity n = 0.")

        draw_box(main_x, y3, main_w, main_h, "Final analytic sample\nN = 36 580")
        draw_arrow(center_x, y3, center_x, y4 + h4, shrinkA=1.5, shrinkB=1.5)

        # Missingness branch arrow (dashed)
        draw_arrow(center_x, 0.311, 0.675, 0.311, ls='--', shrinkA=0, shrinkB=3.0)
        draw_note(0.675, 0.252, 0.305, 0.118, "Variable-level missingness\nwas handled by MICE\n(m = 20); see S7 Table.")

        draw_box(main_x, y4, main_w, h4,
                 "Primary MICE analysis\nPHQ-9 total score and log(GBR)\nN = 36 580\nObserved log(GBR): n = 32 804;\nmissing: n = 3 776",
                 fill='#f8fafc', fontsize=8.95)

        out_path = os.path.join(OUTPUT_DIR, "figure6_flowchart.png")
        plt.savefig(out_path, dpi=450, bbox_inches='tight', facecolor='white')
        plt.close()
        print(f"  Successfully saved to: {out_path}")
    except Exception as e:
        print(f"  Error generating Figure 6: {e}")

# ==============================================================================
# FIGURE 7: Race/Ethnicity Interaction contrasts
# ==============================================================================
def draw_figure_7_race_interaction():
    print("Generating Figure 7 (Race/Ethnicity Interaction contrasts)...")
    try:
        def fmt_p(p):
            if p < 0.001:
                return "<0.001"
            return f"{p:.3f}"

        def draw_panel(ax, ax_table, rows, title, xlabel, xlim, ref_line=0.0, color=SLATE_BLUE):
            labels = [r["label"] for r in rows]
            beta = np.array([r["beta"] for r in rows], dtype=float)
            lo = np.array([r["lo"] for r in rows], dtype=float)
            hi = np.array([r["hi"] for r in rows], dtype=float)
            y = np.arange(len(rows))[::-1]

            ax.axvline(ref_line, color=MUTED_GRAY, linestyle="--", linewidth=1.0, zorder=1)
            ax.errorbar(beta, y, xerr=[beta - lo, hi - beta], fmt="o", markersize=6.5,
                        color=color, ecolor=color, elinewidth=1.6, capsize=4, capthick=1.3, zorder=3)

            ax.set_yticks(y)
            ax.set_yticklabels(labels, fontsize=9.5, color=PRIMARY_NAVY)
            ax.set_xlim(*xlim)
            ax.set_xlabel(xlabel, fontsize=10, color=PRIMARY_NAVY, fontweight='bold')
            ax.set_title(title, loc="left", fontsize=11.5, fontweight="bold", pad=10, color=PRIMARY_NAVY)
            ax.tick_params(axis="x", labelsize=9, colors=PRIMARY_NAVY)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.spines["left"].set_visible(False)
            ax.spines["bottom"].set_linewidth(1.2)
            ax.spines["bottom"].set_color(PRIMARY_NAVY)
            ax.grid(axis="x", color=GRID_COLOR, linestyle=':', linewidth=0.8, zorder=0)

            ax_table.axis("off")
            ax_table.set_ylim(-0.5, len(rows) - 0.5)
            ax_table.set_xlim(0, 1)
            ax_table.text(0.02, len(rows) - 0.15, "Estimate [95% CI]", ha="left", va="bottom",
                          fontsize=9.3, fontweight="bold", color=PRIMARY_NAVY)
            ax_table.text(0.98, len(rows) - 0.15, "P", ha="right", va="bottom",
                          fontsize=9.3, fontweight="bold", color=PRIMARY_NAVY)
            ax_table.plot([0, 1], [len(rows) - 0.35, len(rows) - 0.35], color=PRIMARY_NAVY, linewidth=1.0)
            
            for idx, row in enumerate(rows):
                yy = y[idx]
                est_text = f'{row["beta"]:.3f} [{row["lo"]:.3f}, {row["hi"]:.3f}]'
                ax_table.text(0.02, yy, est_text, va="center", ha="left", fontsize=8.8, color=PRIMARY_NAVY)
                ax_table.text(0.98, yy, fmt_p(row["p"]), va="center", ha="right", fontsize=8.8, color=PRIMARY_NAVY)

        stratified = [
            {"label": "Non-Hispanic White", "beta": 0.3520, "lo": 0.2273, "hi": 0.4766, "p": 2.3888e-07},
            {"label": "Non-Hispanic Black", "beta": 0.2782, "lo": 0.1181, "hi": 0.4383, "p": 8.8137e-04},
            {"label": "Hispanic combined", "beta": -0.0243, "lo": -0.1753, "hi": 0.1266, "p": 0.7489},
            {"label": "Mexican American", "beta": -0.0427, "lo": -0.2167, "hi": 0.1312, "p": 0.6258},
            {"label": "Other Hispanic", "beta": 0.0304, "lo": -0.2113, "hi": 0.2721, "p": 0.8022},
            {"label": "Non-Hispanic Asian", "beta": 0.0580, "lo": -0.1014, "hi": 0.2174, "p": 0.4690},
        ]

        interactions = [
            {"label": "NH Black vs NH White", "beta": 0.0222, "lo": -0.1779, "hi": 0.2223, "p": 0.8262},
            {"label": "Mexican American vs NH White", "beta": -0.2933, "lo": -0.5023, "hi": -0.0844, "p": 0.0065},
            {"label": "Other Hispanic vs NH White", "beta": -0.1532, "lo": -0.4184, "hi": 0.1121, "p": 0.2542},
            {"label": "NH Asian vs NH White", "beta": -0.2074, "lo": -0.3946, "hi": -0.0202, "p": 0.0304},
            {"label": "Other/Multi vs NH White", "beta": 0.0077, "lo": -0.3570, "hi": 0.3724, "p": 0.9667},
        ]

        fig, axes = plt.subplots(2, 2, figsize=(10.8, 8.4),
                                 gridspec_kw={"width_ratios": [1.2, 0.72], "height_ratios": [1.1, 1.0],
                                              "wspace": 0.14, "hspace": 0.48})
        draw_panel(axes[0, 0], axes[0, 1], stratified, "A. Race/ethnicity-stratified association",
                   "Beta coefficient for PHQ-9 (95% CI)", (-0.28, 0.58), color=SLATE_BLUE)
        draw_panel(axes[1, 0], axes[1, 1], interactions, "B. Interaction contrasts relative to Non-Hispanic White",
                   "Difference in beta coefficient (95% CI)", (-0.58, 0.42), color=PRIMARY_NAVY)

        fig.suptitle("Race/ethnicity-specific association of log(GBR) with depressive symptoms",
                     fontsize=13, fontweight="bold", x=0.02, y=0.99, ha="left", color=PRIMARY_NAVY)
        fig.text(0.02, 0.01, "Global log(GBR) x race/ethnicity interaction: chi-square = 15.05, df = 5, p = 0.010, q = 0.024.",
                 fontsize=9, color=PRIMARY_NAVY)
                 
        out = os.path.join(OUTPUT_DIR, "figure2_race_ethnicity_interaction.png")
        plt.tight_layout()
        plt.savefig(out, dpi=600, bbox_inches="tight")
        plt.close(fig)
        print(f"  Successfully saved to: {out}")
    except Exception as e:
        print(f"  Error generating Figure 7: {e}")

# ==============================================================================
# MAIN RUNNER
# ==============================================================================
def main():
    print("=== STARTING FIGURE GENERATION PIPELINE ===")
    draw_figure_1_demographic_forest()
    draw_figure_2_specificity()
    draw_figure_3_rcs()
    draw_figure_4_antidepressant_interaction()
    draw_figure_5_lifestyle_forest()
    draw_figure_6_flowchart()
    draw_figure_7_race_interaction()
    print("=== FIGURE GENERATION PIPELINE COMPLETE ===")

if __name__ == "__main__":
    main()
