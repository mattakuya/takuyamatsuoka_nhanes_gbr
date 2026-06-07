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
        {'label': 'Alcohol Consumption (NIAAA)', 'is_header': True, 'is_primary': False},
        {'label': '  Non-drinker', 'is_header': False, 'is_primary': False, 'n': '7557', 'beta': 0.1231, 'ci_lo': -0.0372, 'ci_hi': 0.2834, 'p_val': '0.130'},
        {'label': '  Low-risk drinker', 'is_header': False, 'is_primary': False, 'n': '18666', 'beta': 0.2273, 'ci_lo': 0.1200, 'ci_hi': 0.3347, 'p_val': '< 0.001'},
        {'label': '  Heavy drinker', 'is_header': False, 'is_primary': False, 'n': '2077', 'beta': 0.5407, 'ci_lo': 0.2076, 'ci_hi': 0.8738, 'p_val': '0.002'},
        {'label': 'Cotinine Level (Smoking Exposure)', 'is_header': True, 'is_primary': False},
        {'label': '  Cotinine < 3 ng/mL (Non-exposed)', 'is_header': False, 'is_primary': False, 'n': '24283', 'beta': 0.1861, 'ci_lo': 0.0937, 'ci_hi': 0.2784, 'p_val': '< 0.001'},
        {'label': '  Cotinine 3-30 ng/mL (Transitional)', 'is_header': False, 'is_primary': False, 'n': '1386', 'beta': 0.4568, 'ci_lo': 0.0627, 'ci_hi': 0.8509, 'p_val': '0.024'},
        {'label': '  Cotinine > 30 ng/mL (Active smoker)', 'is_header': False, 'is_primary': False, 'n': '7359', 'beta': 0.3144, 'ci_lo': 0.1165, 'ci_hi': 0.5123, 'p_val': '0.002'},
        {'label': 'Systemic Inflammation (hs-CRP, I/J only)', 'is_header': True, 'is_primary': False},
        {'label': '  hsCRP < 1 mg/L', 'is_header': False, 'is_primary': False, 'n': '3179', 'beta': 0.2114, 'ci_lo': -0.0485, 'ci_hi': 0.4712, 'p_val': '0.107'},
        {'label': '  hsCRP 1-3 mg/L', 'is_header': False, 'is_primary': False, 'n': '3507', 'beta': 0.1773, 'ci_lo': -0.0669, 'ci_hi': 0.4215, 'p_val': '0.148'},
        {'label': '  hsCRP > 3 mg/L', 'is_header': False, 'is_primary': False, 'n': '3830', 'beta': 0.1325, 'ci_lo': -0.1230, 'ci_hi': 0.3880, 'p_val': '0.297'},
    ]

def main():
    try:
        try:
            results_file = get_latest_results_file()
            print(f"Reading results for latter subgroups from: {results_file}")
        except FileNotFoundError:
            results_file = None
            print("Results markdown not found; using embedded latter forest-plot values.")
        
        table2_md = build_table_2(results_file) if results_file else ""
        table3_md = build_table_3_subgroups(results_file) if results_file else ""
        
        items = [] if results_file else fallback_items()
        
        # 1. Parse Table 2 for the Overall Full Sample OLS result as reference
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
        
        # 2. Parse Table 3 latter subgroups (Alcohol, Cotinine, hsCRP)
        target_headers = [
            "Alcohol Consumption (NIAAA)",
            "Cotinine Level (Smoking Exposure)",
            "Systemic Inflammation (hs-CRP, I/J only)"
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
                n = parts[1]
                est = parse_estimate(parts[2])
                if est:
                    beta, ci_lo, ci_hi = est
                    p_val = format_p_value(parts[4])
                    
                    # Clean label formatting
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
            1, 2, figsize=(12.2, 6.9),
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
        ax_plot.set_xlim(-0.25, 0.95) # Wider range since heavy drinker is around 0.54
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
        plot_path = "results/figure5_forest_plot_latter.png"
        plt.tight_layout()
        plt.savefig(plot_path, dpi=300, bbox_inches='tight')
        
        # Also copy to results/tp_figures/supp_figure3_lifestyle_clinical_subgroups.png for medrxiv single docx
        tp_fig_dir = "results/tp_figures"
        os.makedirs(tp_fig_dir, exist_ok=True)
        tp_fig_path = os.path.join(tp_fig_dir, "supp_figure3_lifestyle_clinical_subgroups.png")
        import shutil
        shutil.copy2(plot_path, tp_fig_path)
        
        plt.close()
        print(f"Latter forest plot successfully generated and saved to: {os.path.abspath(plot_path)} and copied to {tp_fig_path}")
        
    except Exception as e:
        print(f"Error generating latter forest plot: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
