import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from generate_manuscript import (
    get_latest_results_file,
    build_table_2,
    build_table_3_subgroups
)

def parse_estimate(val_str):
    # Parses beta, ci_lo, ci_hi from e.g. "0.3013 [0.1541, 0.4484]"
    match = re.search(r"(-?\d+\.\d+)\s*\[\s*(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)\s*\]", val_str)
    if match:
        return float(match.group(1)), float(match.group(2)), float(match.group(3))
    return None

def format_p_value(p_str):
    try:
        p_val = float(p_str)
        if p_val < 0.001:
            return "< 0.001"
        else:
            return f"{p_val:.3f}"
    except:
        return p_str

def fallback_items():
    return [
        {'label': 'Overall (Full Sample)', 'is_header': False, 'is_primary': True, 'n': '36580', 'beta': 0.2564, 'ci_lo': 0.1713, 'ci_hi': 0.3414, 'p_val': '< 0.001'},
        {'label': 'BMI Categories (kg/m²)', 'is_header': True, 'is_primary': False},
        {'label': '  Normal (18.5-24.9)', 'is_header': False, 'is_primary': False, 'n': '9783', 'beta': 0.3013, 'ci_lo': 0.1541, 'ci_hi': 0.4484, 'p_val': '< 0.001'},
        {'label': '  Overweight (25.0-29.9)', 'is_header': False, 'is_primary': False, 'n': '11205', 'beta': 0.3278, 'ci_lo': 0.1961, 'ci_hi': 0.4594, 'p_val': '< 0.001'},
        {'label': '  Obese (≥ 30.0)', 'is_header': False, 'is_primary': False, 'n': '13007', 'beta': 0.1623, 'ci_lo': 0.0325, 'ci_hi': 0.2921, 'p_val': '0.015'},
        {'label': 'Sex', 'is_header': True, 'is_primary': False},
        {'label': '  Male', 'is_header': False, 'is_primary': False, 'n': '17783', 'beta': 0.2578, 'ci_lo': 0.1550, 'ci_hi': 0.3606, 'p_val': '< 0.001'},
        {'label': '  Female', 'is_header': False, 'is_primary': False, 'n': '18797', 'beta': 0.2667, 'ci_lo': 0.1514, 'ci_hi': 0.3820, 'p_val': '< 0.001'},
        {'label': 'Age Groups', 'is_header': True, 'is_primary': False},
        {'label': '  Age 18-39', 'is_header': False, 'is_primary': False, 'n': '13354', 'beta': 0.2066, 'ci_lo': 0.0896, 'ci_hi': 0.3235, 'p_val': '< 0.001'},
        {'label': '  Age 40-64', 'is_header': False, 'is_primary': False, 'n': '14713', 'beta': 0.2670, 'ci_lo': 0.1346, 'ci_hi': 0.3993, 'p_val': '< 0.001'},
        {'label': '  Age 65+', 'is_header': False, 'is_primary': False, 'n': '8513', 'beta': 0.1832, 'ci_lo': 0.0260, 'ci_hi': 0.3404, 'p_val': '0.023'},
        {'label': 'Race / Ethnicity', 'is_header': True, 'is_primary': False},
        {'label': '  Non-Hispanic White', 'is_header': False, 'is_primary': False, 'n': '14561', 'beta': 0.3520, 'ci_lo': 0.2273, 'ci_hi': 0.4766, 'p_val': '< 0.001'},
        {'label': '  Non-Hispanic Black', 'is_header': False, 'is_primary': False, 'n': '7912', 'beta': 0.2782, 'ci_lo': 0.1181, 'ci_hi': 0.4383, 'p_val': '< 0.001'},
        {'label': '  Hispanic (combined)', 'is_header': False, 'is_primary': False, 'n': '9487', 'beta': -0.0243, 'ci_lo': -0.1753, 'ci_hi': 0.1266, 'p_val': '0.749'},
        {'label': '  Non-Hispanic Asian', 'is_header': False, 'is_primary': False, 'n': '3108', 'beta': 0.0580, 'ci_lo': -0.1014, 'ci_hi': 0.2174, 'p_val': '0.469'},
    ]

def main():
    try:
        try:
            results_file = get_latest_results_file()
            print(f"Reading results from: {results_file}")
        except FileNotFoundError:
            results_file = None
            print("Results markdown not found; using embedded forest-plot values.")
        
        table2_md = build_table_2(results_file) if results_file else ""
        table3_md = build_table_3_subgroups(results_file) if results_file else ""
        
        items = [] if results_file else fallback_items()
        
        # 1. Parse Table 2 for the Overall Full Sample WLS/OLS result
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
                            'n': n,
                            'beta': beta,
                            'ci_lo': ci_lo,
                            'ci_hi': ci_hi,
                            'p_val': p_val
                        })
        
        # 2. Parse Table 3 subgroups (BMI, Sex, Age, Race)
        target_headers = [
            "BMI Categories (kg/m²)",
            "Sex",
            "Age Groups",
            "Race / Ethnicity"
        ]
        
        current_header = None
        header_active = False
        
        for line in table3_md.split('\n'):
            if not line.strip() or '|' not in line:
                continue
            parts = [x.strip() for x in line.split('|')[1:-1]]
            if not parts or len(parts) < 2:
                continue
                
            subgroup = parts[0]
            # Check if this is a section header
            if subgroup.startswith('**') and subgroup.endswith('**'):
                clean_header = subgroup.replace('**', '')
                if clean_header in target_headers:
                    current_header = clean_header
                    header_active = True
                    items.append({
                        'label': clean_header,
                        'is_header': True,
                        'is_primary': False
                    })
                else:
                    header_active = False
            elif header_active:
                # Exclude sub-subgroups like Mexican American / Other Hispanic to keep forest plot concise,
                # focusing on the main standard subcategories:
                # Sex: Male, Female
                # BMI: Normal BMI (18.5-24.9), Overweight (25-29.9), Obese (≥30)
                # Age: Age 18-39, Age 40-64, Age 65+
                # Race: Non-Hispanic White, Non-Hispanic Black, Hispanic (combined), Non-Hispanic Asian (G-J only)
                label = subgroup
                if label in ["Mexican American", "Other Hispanic"]:
                    continue # Skip to avoid redundancy since Hispanic (combined) is present
                
                n = parts[1]
                est = parse_estimate(parts[2])
                if est:
                    beta, ci_lo, ci_hi = est
                    p_val = format_p_value(parts[4])
                    
                    # Clean label formatting
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
                        'n': n,
                        'beta': beta,
                        'ci_lo': ci_lo,
                        'ci_hi': ci_hi,
                        'p_val': p_val
                    })
        
        # 3. Plotting using matplotlib
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Helvetica', 'Arial', 'DejaVu Sans', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        
        num_items = len(items)
        y_coords = np.arange(num_items)[::-1]
        
        fig, (ax_plot, ax_table) = plt.subplots(
            1, 2, figsize=(12.2, 7.8),
            gridspec_kw={'width_ratios': [3.7, 2.9]}
        )
        fig.subplots_adjust(wspace=0.08)
        
        # Color & Style constants (Unified Design System)
        primary_color = '#263746'  # Primary Navy
        secondary_color = '#2f6f9f'  # Slate Blue
        ref_line_color = '#6b7280'  # Muted Gray
        grid_color = '#e2e8f0'  # Grid Lines
        
        # Draw plot
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
                ax_plot.errorbar(beta, y, xerr=[[err_lo], [err_hi]], fmt='s', color=primary_color,
                                 markersize=7.0, elinewidth=1.9, capsize=4, capthick=1.4, zorder=5)
            else:
                ax_plot.errorbar(beta, y, xerr=[[err_lo], [err_hi]], fmt='o', color=secondary_color,
                                 markersize=5.6, elinewidth=1.45, capsize=3.5, capthick=1.1, zorder=4)
                                 
        # Reference line
        ax_plot.axvline(x=0, color=ref_line_color, linestyle='--', linewidth=1.0, zorder=1)
        
        # Axes labels and styling
        ax_plot.set_xlabel('Beta coefficient for PHQ-9 (95% CI)', fontsize=10.2, labelpad=12, fontweight='bold', color=primary_color)
        ax_plot.set_xlim(-0.25, 0.55)
        ax_plot.set_ylim(-1, num_items + 0.7)
        
        # Set Y ticks and labels
        y_labels = [item['label'] for item in items]
        ax_plot.set_yticks(y_coords)
        ax_plot.set_yticklabels(y_labels, fontsize=9.5, color=primary_color)
        
        # Bold headers and primary label
        for tick_label, item in zip(ax_plot.get_yticklabels(), items):
            if item['is_header']:
                tick_label.set_fontweight('bold')
                tick_label.set_fontsize(10.5)
                tick_label.set_color(primary_color)
            elif item['is_primary']:
                tick_label.set_fontweight('bold')
                tick_label.set_fontsize(10.0)
                tick_label.set_color(primary_color)
                
        ax_plot.spines['left'].set_visible(False)
        ax_plot.spines['top'].set_visible(False)
        ax_plot.spines['right'].set_visible(False)
        ax_plot.spines['bottom'].set_linewidth(1.2)
        ax_plot.spines['bottom'].set_color(primary_color)
        ax_plot.tick_params(axis='x', labelsize=9.5, colors=primary_color)
        ax_plot.grid(axis='x', color=grid_color, linestyle=':', linewidth=0.8, zorder=0)
        
        # Table subplot
        ax_table.axis('off')
        ax_table.set_ylim(-1, num_items + 0.7)
        ax_table.set_xlim(0, 1)
        
        # Table headers
        header_y = num_items + 0.15
        rule_top_y = num_items + 0.55
        rule_mid_y = num_items - 0.25
        ax_table.text(0.08, header_y, "N", ha='center', va='center', fontweight='bold', fontsize=10.0, color=primary_color)
        ax_table.text(0.50, header_y, "Beta [95% CI]", ha='center', va='center', fontweight='bold', fontsize=10.0, color=primary_color)
        ax_table.text(0.91, header_y, "P-value", ha='center', va='center', fontweight='bold', fontsize=10.0, color=primary_color)
        
        # Booktabs horizontal lines (Unified Design System rule weights)
        ax_table.plot([0, 1], [rule_top_y, rule_top_y], color=primary_color, linewidth=1.5)
        ax_table.plot([0, 1], [rule_mid_y, rule_mid_y], color=primary_color, linewidth=1.0)
        
        for idx, item in enumerate(items):
            y = y_coords[idx]
            if item['is_header']:
                continue
                
            n_str = item['n'].replace(',', ' ')
            beta_ci_str = f"{item['beta']:.4f} [{item['ci_lo']:.4f}, {item['ci_hi']:.4f}]"
            p_str = item['p_val']
            
            font_w = 'bold' if item['is_primary'] else 'normal'
            color_val = primary_color
            font_sz = 9.4 if item['is_primary'] else 9.1
            
            ax_table.text(0.08, y, n_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=color_val)
            ax_table.text(0.50, y, beta_ci_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=color_val)
            ax_table.text(0.91, y, p_str, ha='center', va='center', fontweight=font_w, fontsize=font_sz, color=color_val)
            
        # Bottom booktabs line
        ax_table.plot([0, 1], [-0.5, -0.5], color=primary_color, linewidth=1.5)
        
        os.makedirs("results", exist_ok=True)
        plot_path = "results/figure1_forest_plot.png"
        plt.tight_layout()
        plt.savefig(plot_path, dpi=300, bbox_inches='tight')
        
        # Also copy to results/tp_figures/supp_figure2_demographic_subgroups.png for medrxiv single docx
        tp_fig_dir = "results/tp_figures"
        os.makedirs(tp_fig_dir, exist_ok=True)
        tp_fig_path = os.path.join(tp_fig_dir, "supp_figure2_demographic_subgroups.png")
        import shutil
        shutil.copy2(plot_path, tp_fig_path)
        
        plt.close()
        print(f"Forest plot successfully generated and saved to: {os.path.abspath(plot_path)} and copied to {tp_fig_path}")
        
    except Exception as e:
        print(f"Error generating forest plot: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
