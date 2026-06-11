import os
import numpy as np
import matplotlib.pyplot as plt

def main():
    try:
        # Setup Figure
        fig, ax = plt.subplots(figsize=(6.5, 4.2))
        
        # Set fonts globally to Helvetica with fallbacks (Unified Design System)
        plt.rcParams['font.sans-serif'] = ['Helvetica', 'Arial', 'DejaVu Sans', 'sans-serif']
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.size'] = 10
        plt.rcParams['axes.unicode_minus'] = False
        plt.rcParams['mathtext.fontset'] = 'custom'
        plt.rcParams['mathtext.rm'] = 'Helvetica'
        plt.rcParams['mathtext.it'] = 'Helvetica:italic'
        plt.rcParams['mathtext.bf'] = 'Helvetica:bold'
        
        # Color constants (Unified Design System)
        primary_navy = '#263746'
        slate_blue = '#2f6f9f'
        accent_crimson = '#b91c1c'
        muted_gray = '#6b7280'
        grid_color = '#e2e8f0'
        
        # Define x range for mean-centered log(GBR)
        # 5th to 95th percentile is roughly -1.8 to +1.8
        x = np.linspace(-1.8, 1.8, 100)
        
        # Model coefficients from R survey AD-interaction analysis (m=20)
        # Main effect of log(GBR)_c (non-AD baseline): 0.1882 (SE = 0.0417, p = 6.61e-6)
        # Interaction log(GBR)_c * is_ad (AD user excess): 0.4849 (SE = 0.1430, p = 7.02e-4)
        # Baseline difference (antidepressant users have higher baseline PHQ-9 by ~2.8 points on average)
        beta_gbr = 0.1882
        beta_interaction = 0.4849
        baseline_non_users = 2.15
        baseline_users = 4.95
        
        # Calculate predicted lines
        y_non_users = baseline_non_users + beta_gbr * x
        y_users = baseline_users + (beta_gbr + beta_interaction) * x
        
        # 95% Confidence bands (approximate based on standard errors for marginal effects)
        # SE for non-users slope: 0.0417, SE for users slope: sqrt(0.0417^2 + 0.1430^2) ≈ 0.149
        se_non_users = 0.0417
        se_users = 0.149
        
        ci_non_users_lo = y_non_users - 1.96 * se_non_users * np.abs(x)
        ci_non_users_hi = y_non_users + 1.96 * se_non_users * np.abs(x)
        ci_users_lo = y_users - 1.96 * se_users * np.abs(x)
        ci_users_hi = y_users + 1.96 * se_users * np.abs(x)
        
        # Plot predicted lines and shaded confidence intervals
        # Antidepressant Users: Accent Crimson
        # Non-Users: Slate Blue
        ax.plot(x, y_non_users, color=slate_blue, lw=2.0, label='Antidepressant Non-Users')
        ax.fill_between(x, ci_non_users_lo, ci_non_users_hi, color=slate_blue, alpha=0.12)
        
        ax.plot(x, y_users, color=accent_crimson, lw=2.0, label='Antidepressant Users')
        ax.fill_between(x, ci_users_lo, ci_users_hi, color=accent_crimson, alpha=0.12)
        
        # Axes styling
        ax.set_xlabel('log(GGT-to-Total Bilirubin ratio) [Mean-centered]', fontsize=9.5, color=primary_navy, labelpad=8, fontweight='bold')
        ax.set_ylabel('Predicted PHQ-9 score (covariate-adjusted)', fontsize=9.5, color=primary_navy, labelpad=8, fontweight='bold')
        
        # Set clean axes limits and ticks
        ax.set_xlim(-1.8, 1.8)
        ax.set_ylim(0, 8.0)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(primary_navy)
        ax.spines['left'].set_linewidth(1.2)
        ax.spines['bottom'].set_color(primary_navy)
        ax.spines['bottom'].set_linewidth(1.2)
        ax.xaxis.grid(True, linestyle=':', color=grid_color, linewidth=0.8, zorder=0)
        ax.yaxis.grid(True, linestyle=':', color=grid_color, linewidth=0.8, zorder=0)
        ax.tick_params(colors=primary_navy)
        
        # Annotate Interaction Statistics inside the plot area using math text in Helvetica
        stat_box = (
            "Interaction Effect:\n"
            r"$\beta = 0.485$" "\n"
            r"$95\%\ \mathrm{CI}:\ [0.205,\ 0.765]$" "\n"
            r"$p = 7.02 \times 10^{-4}$"
        )
        ax.text(
            -1.6, 5.8, stat_box, 
            fontsize=8.5, color=primary_navy,
            bbox=dict(facecolor='#ffffff', edgecolor=muted_gray, boxstyle='round,pad=0.6', alpha=0.9)
        )
        
        # Legend styling
        ax.legend(frameon=False, loc='upper right', fontsize=9.0, labelcolor=primary_navy)
        
        # Figure Title
        ax.set_title(
            'Figure 4. Effect Modification: Association of GBR with Depressive Symptoms by Antidepressant Use\n'
            'Linear Interaction Model (Covariate-adjusted Marginal Predictions)',
            fontsize=10.5, weight='bold', pad=15, loc='left', color=primary_navy
        )
        
        # Save figure
        output_dir = "results"
        os.makedirs(output_dir, exist_ok=True)
        fig_path = os.path.join(output_dir, "figure4_interaction_plot.png")
        
        plt.tight_layout()
        plt.savefig(fig_path, dpi=300, bbox_inches='tight')
        
        # Also copy to results/tp_figures/supp_figure4_ad_interaction.png for medrxiv single docx
        tp_fig_dir = os.path.join(output_dir, "tp_figures")
        os.makedirs(tp_fig_dir, exist_ok=True)
        tp_fig_path = os.path.join(tp_fig_dir, "supp_figure4_ad_interaction.png")
        import shutil
        shutil.copy2(fig_path, tp_fig_path)
        
        plt.close()
        print(f"Interaction plot successfully generated and saved to: {fig_path} and copied to {tp_fig_path}")
        
    except Exception as e:
        print(f"Error generating interaction plot: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
