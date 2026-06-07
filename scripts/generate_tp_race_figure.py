import os
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np


ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "results" / "tp_figures"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def fmt_p(p):
    if p < 0.001:
        return "<0.001"
    return f"{p:.3f}"


def draw_panel(ax, ax_table, rows, title, xlabel, xlim, ref_line=0.0, color="#2f6f9f"):
    labels = [r["label"] for r in rows]
    beta = np.array([r["beta"] for r in rows], dtype=float)
    lo = np.array([r["lo"] for r in rows], dtype=float)
    hi = np.array([r["hi"] for r in rows], dtype=float)
    y = np.arange(len(rows))[::-1]

    # Color & Style constants (Unified Design System)
    primary_color = '#263746'
    muted_gray = '#6b7280'
    grid_color = '#e2e8f0'

    ax.axvline(ref_line, color=muted_gray, linestyle="--", linewidth=1.0, zorder=1)
    ax.errorbar(
        beta,
        y,
        xerr=[beta - lo, hi - beta],
        fmt="o",
        markersize=6.5,
        color=color,
        ecolor=color,
        elinewidth=1.6,
        capsize=4,
        capthick=1.3,
        zorder=3,
    )

    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9.5, color=primary_color)
    ax.set_xlim(*xlim)
    ax.set_xlabel(xlabel, fontsize=10, color=primary_color, fontweight='bold')
    ax.set_title(title, loc="left", fontsize=11.5, fontweight="bold", pad=10, color=primary_color)
    ax.tick_params(axis="x", labelsize=9, colors=primary_color)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_linewidth(1.2)
    ax.spines["bottom"].set_color(primary_color)
    ax.grid(axis="x", color=grid_color, linestyle=':', linewidth=0.8, zorder=0)

    ax_table.axis("off")
    ax_table.set_ylim(-0.5, len(rows) - 0.5)
    ax_table.set_xlim(0, 1)
    ax_table.text(0.02, len(rows) - 0.15, "Estimate [95% CI]", ha="left", va="bottom",
                  fontsize=9.3, fontweight="bold", color=primary_color)
    ax_table.text(0.98, len(rows) - 0.15, "P", ha="right", va="bottom",
                  fontsize=9.3, fontweight="bold", color=primary_color)
    ax_table.plot([0, 1], [len(rows) - 0.35, len(rows) - 0.35],
                  color=primary_color, linewidth=1.0)
    for idx, row in enumerate(rows):
        yy = y[idx]
        est_text = f'{row["beta"]:.3f} [{row["lo"]:.3f}, {row["hi"]:.3f}]'
        ax_table.text(0.02, yy, est_text, va="center", ha="left",
                      fontsize=8.8, color=primary_color)
        ax_table.text(0.98, yy, fmt_p(row["p"]), va="center", ha="right",
                      fontsize=8.8, color=primary_color)


def main():
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Helvetica", "Arial", "DejaVu Sans", "sans-serif"]
    plt.rcParams["axes.unicode_minus"] = False

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

    fig, axes = plt.subplots(
        2,
        2,
        figsize=(10.8, 8.4),
        gridspec_kw={
            "width_ratios": [1.2, 0.72],
            "height_ratios": [1.1, 1.0],
            "wspace": 0.14,
            "hspace": 0.48,
        },
    )
    draw_panel(
        axes[0, 0],
        axes[0, 1],
        stratified,
        "A. Race/ethnicity-stratified association",
        "Beta coefficient for PHQ-9 (95% CI)",
        (-0.28, 0.58),
        color="#2f6f9f",  # Slate Blue
    )
    draw_panel(
        axes[1, 0],
        axes[1, 1],
        interactions,
        "B. Interaction contrasts relative to Non-Hispanic White",
        "Difference in beta coefficient (95% CI)",
        (-0.58, 0.42),
        color="#263746",  # Primary Navy
    )

    fig.suptitle(
        "Race/ethnicity-specific association of log(GBR) with depressive symptoms",
        fontsize=13,
        fontweight="bold",
        x=0.02,
        y=0.99,
        ha="left",
        color="#263746",
    )
    fig.text(
        0.02,
        0.01,
        "Global log(GBR) x race/ethnicity interaction: chi-square = 15.05, df = 5, p = 0.010, q = 0.024.",
        fontsize=9,
        color="#263746",
    )
    plt.tight_layout()
    out = OUT_DIR / "figure2_race_ethnicity_interaction.png"
    fig.savefig(out, dpi=600, bbox_inches="tight")
    
    # Also copy to figure1_race_ethnicity_interaction.png for medrxiv single docx
    out_medrxiv = OUT_DIR / "figure1_race_ethnicity_interaction.png"
    import shutil
    shutil.copy2(out, out_medrxiv)
    
    plt.close(fig)
    print(f"Saved {out} and copied to {out_medrxiv}")


if __name__ == "__main__":
    main()
