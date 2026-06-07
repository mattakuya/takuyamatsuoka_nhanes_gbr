import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from generate_manuscript import (
    get_latest_results_file,
    build_table_3
)

def markdown_to_dataframe(md_table_str):
    lines = md_table_str.strip().split('\n')
    lines = [l for l in lines if not re.match(r'^\s*\|?\s*:?-+:?\s*\|', l)]
    
    if not lines:
        return pd.DataFrame()
        
    header_line = lines[0]
    headers = [x.strip() for x in header_line.split('|')[1:-1]]
    
    rows = []
    for line in lines[1:]:
        if not line.strip():
            continue
        row = [x.strip() for x in line.split('|')[1:-1]]
        if len(row) < len(headers):
            row += [''] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[:len(headers)]
        rows.append(row)
        
    return pd.DataFrame(rows, columns=headers)

def parse_ci_column(ci_str):
    match = re.search(r'(-?\d+\.\d+)\s*\[\s*(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)\s*\]', ci_str)
    if match:
        return float(match.group(1)), float(match.group(2)), float(match.group(3))
    return None

def main():
    try:
        results_file = get_latest_results_file()
        table3_md = build_table_3(results_file)
        df = markdown_to_dataframe(table3_md)
        
        # Parse and filter univariate models (we only plot the single standardized ones for clean comparison)
        plot_data = []
        
        # Color coding based on biological role (Unified Design System)
        primary_color = '#263746'    # Primary Navy (GBR)
        accent_color = '#b91c1c'     # Accent Crimson (GGT)
        emerald_color = '#10b981'    # Emerald Green (Bilirubin)
        muted_gray = '#6b7280'       # Muted Gray (ALT, AST, reference lines)
        grid_color = '#e2e8f0'       # Grid Lines
        
        role_metadata = {
            'GGT/Bil Ratio (per 1 SD)': {'color': primary_color, 'label': 'GGT/Bil Ratio (GBR)', 'role': 'Primary Balance Indicator', 'marker': 's'},
            'GGT alone (per 1 SD)': {'color': accent_color, 'label': 'GGT (Pro-oxidant marker)', 'role': 'Pro-oxidant component', 'marker': 'o'},
            'Bilirubin alone (per 1 SD)': {'color': emerald_color, 'label': 'Total Bilirubin (Antioxidant)', 'role': 'Antioxidant component', 'marker': 'o'},
            'ALT (per 1 SD)': {'color': muted_gray, 'label': 'ALT', 'role': 'General Liver Marker (Control)', 'marker': '^'},
            'AST (per 1 SD)': {'color': muted_gray, 'label': 'AST', 'role': 'General Liver Marker (Control)', 'marker': '^'}
        }
        
        for idx, row in df.iterrows():
            term = str(row['Model / Term']).strip()
            n_val = str(row['N']).strip()
            ci_str = str(row['Beta [95% CI]']).strip()
            p_val = str(row['P-value']).strip()
            
            # We only plot the univariate models for clear specificity comparison
            if term in role_metadata:
                parsed = parse_ci_column(ci_str)
                if parsed:
                    beta, ci_lo, ci_hi = parsed
                    meta = role_metadata[term]
                    plot_data.append({
                        'label': meta['label'],
                        'role': meta['role'],
                        'beta': beta,
                        'ci_lo': ci_lo,
                        'ci_hi': ci_hi,
                        'n': n_val,
                        'p': p_val,
                        'color': meta['color'],
                        'marker': meta['marker']
                    })
        
        # Reverse for bottom-up plotting
        plot_data.reverse()
        
        # Setup Figure
        fig, (ax_plot, ax_text) = plt.subplots(
            1, 2, 
            figsize=(12, 5.0), 
            gridspec_kw={'width_ratios': [1.8, 1.2]}
        )
        
        # Set fonts globally (Unified Design System)
        plt.rcParams['font.sans-serif'] = ['Helvetica', 'Arial', 'DejaVu Sans', 'sans-serif']
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.size'] = 10
        
        # 1. Plotting the Forest Plot (Left Subplot)
        for i, item in enumerate(plot_data):
            beta = item['beta']
            err_left = beta - item['ci_lo']
            err_right = item['ci_hi'] - beta
            
            # Plot point and 95% CI
            ax_plot.errorbar(
                [beta], [i], 
                xerr=[[err_left], [err_right]], 
                fmt=item['marker'], color=item['color'], ecolor=item['color'], 
                capsize=4, elinewidth=1.5, ms=7 if item['marker'] == 's' else 6
            )
            
            # Label on the y-axis
            weight = 'bold' if 'Ratio' in item['label'] else 'normal'
            ax_plot.text(
                -0.25, i, f" {item['label']}", 
                va='bottom', ha='left', fontsize=9.5, color=primary_color, weight=weight
            )
            # Sub-label for role description
            ax_plot.text(
                -0.25, i - 0.22, f"   ({item['role']})", 
                va='top', ha='left', fontsize=7.5, color=muted_gray, style='italic'
            )
            
        # Style Left Plot
        ax_plot.axvline(0.0, color=muted_gray, ls='--', lw=1.0) # Zero effect reference line
        ax_plot.set_yticks([]) # Hide default y-ticks
        ax_plot.set_ylim(-0.6, len(plot_data) - 0.4)
        ax_plot.set_xlim(-0.25, 0.4)
        ax_plot.set_xlabel('Standardized Effect Size (Beta) [95% CI] per 1 SD', weight='bold', fontsize=9.5, labelpad=10, color=primary_color)
        ax_plot.spines['left'].set_visible(False)
        ax_plot.spines['top'].set_visible(False)
        ax_plot.spines['right'].set_visible(False)
        ax_plot.spines['bottom'].set_linewidth(1.2)
        ax_plot.spines['bottom'].set_color(primary_color)
        ax_plot.xaxis.grid(True, linestyle=':', color=grid_color, linewidth=0.8, zorder=0)
        
        # 2. Text Table (Right Subplot)
        ax_text.axis('off')
        ax_text.set_ylim(-0.6, len(plot_data) - 0.4)
        
        # Headers for Table
        ax_text.text(0.0, len(plot_data) - 0.2, 'N', weight='bold', ha='center', fontsize=9.5, color=primary_color)
        ax_text.text(0.42, len(plot_data) - 0.2, 'Beta [95% CI]', weight='bold', ha='center', fontsize=9.5, color=primary_color)
        ax_text.text(0.88, len(plot_data) - 0.2, 'P-value', weight='bold', ha='center', fontsize=9.5, color=primary_color)
        
        # Booktabs style lines (Unified rule weights and color)
        ax_text.plot([-0.15, 1.05], [len(plot_data) + 0.1, len(plot_data) + 0.1], color=primary_color, lw=1.5)
        ax_text.plot([-0.15, 1.05], [len(plot_data) - 0.5, len(plot_data) - 0.5], color=primary_color, lw=1.0)
        ax_text.plot([-0.15, 1.05], [-0.5, -0.5], color=primary_color, lw=1.5)
        
        for i, item in enumerate(plot_data):
            ci_txt = f"{item['beta']:.3f} [{item['ci_lo']:.3f}, {item['ci_hi']:.3f}]"
            
            p_str = item['p']
            try:
                p_float = float(p_str)
                if p_float < 0.001:
                    p_str = "<0.001"
                else:
                    p_str = f"{p_float:.3f}"
            except:
                pass
                
            weight = 'bold' if 'Ratio' in item['label'] else 'normal'
            color = primary_color
            
            ax_text.text(0.0, i, item['n'], ha='center', va='center', fontsize=9, weight=weight, color=color)
            ax_text.text(0.42, i, ci_txt, ha='center', va='center', fontsize=9, weight=weight, color=color)
            ax_text.text(0.88, i, p_str, ha='center', va='center', fontsize=9, weight=weight, color=color)
            
        # Title of the figure
        fig.suptitle(
            'Figure 2. Specificity Analysis: Standardized Effect Sizes (per 1 SD) of GBR and Other Enzymes', 
            fontsize=12, weight='bold', color=primary_color, y=0.96, x=0.04, ha='left'
        )
        
        # Save figure
        output_dir = "results"
        os.makedirs(output_dir, exist_ok=True)
        fig_path = os.path.join(output_dir, "figure2_specificity_plot.png")
        
        plt.tight_layout()
        plt.subplots_adjust(top=0.86)
        
        plt.savefig(fig_path, dpi=600, bbox_inches='tight')
        
        # Also copy to results/tp_figures/figure2_specificity.png for medrxiv single docx
        tp_fig_dir = os.path.join(output_dir, "tp_figures")
        os.makedirs(tp_fig_dir, exist_ok=True)
        tp_fig_path = os.path.join(tp_fig_dir, "figure2_specificity.png")
        import shutil
        shutil.copy2(fig_path, tp_fig_path)
        
        plt.close()
        print(f"Specificity plot successfully generated and saved to: {fig_path} and copied to {tp_fig_path}")
        
    except Exception as e:
        print(f"Error generating specificity plot: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
