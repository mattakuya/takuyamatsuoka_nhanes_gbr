import os
import shutil
import textwrap
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def main():
    # Set fonts
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Helvetica', 'Arial', 'DejaVu Sans', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False

    fig, ax = plt.subplots(figsize=(7.4, 8.8))
    ax.axis('off')
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    # Styling constants (Unified Design System)
    box_edge = '#263746'
    note_edge = '#6b7280'
    text_color = '#263746'
    main_fill = '#ffffff'
    note_fill = '#f8fafc'
    arrow_color = '#263746'

    def wrap_text(text, width):
        return "\n".join(
            "\n".join(textwrap.wrap(part, width=width, break_long_words=False))
            for part in text.split("\n")
        )

    def draw_box(x, y, w, h, text, fill=main_fill, edge=box_edge, fontsize=10.5, bold=True):
        rect = patches.FancyBboxPatch(
            (x, y), w, h,
            boxstyle="round,pad=0.004,rounding_size=0.018",
            facecolor=fill, edgecolor=edge, linewidth=1.6,
        )
        ax.add_patch(rect)
        ax.text(
            x + w / 2, y + h / 2, text,
            ha='center', va='center', fontsize=fontsize, color=text_color,
            fontweight='bold' if bold else 'normal', linespacing=1.25,
        )

    def draw_note(x, y, w, h, text, color='#4b5563', fontsize=7.6):
        rect = patches.FancyBboxPatch(
            (x, y), w, h,
            boxstyle="round,pad=0.004,rounding_size=0.012",
            facecolor=note_fill, edgecolor=note_edge, linewidth=1.0,
        )
        ax.add_patch(rect)
        ax.text(
            x + w / 2, y + h / 2, text,
            ha='center', va='center', fontsize=fontsize, color=color,
            linespacing=1.12, clip_on=False,
        )

    def draw_arrow(x_start, y_start, x_end, y_end, style='->', color=arrow_color, ls='-', shrinkA=0, shrinkB=0):
        arrow = patches.FancyArrowPatch(
            (x_start, y_start), (x_end, y_end),
            arrowstyle=style,
            mutation_scale=16,
            linewidth=1.55,
            color=color,
            linestyle=ls,
            shrinkA=shrinkA,
            shrinkB=shrinkB,
            zorder=5,
        )
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

    draw_box(
        main_x, y1, main_w, main_h,
        "Initial NHANES sample\n2007–2018 cycles (E–J)\nN = 59 842",
    )
    # Vertical arrow with proper shrink
    draw_arrow(center_x, y1, center_x, y2 + main_h, shrinkA=1.5, shrinkB=1.5)

    # Excluded branch arrow (red, solid) - aligned to vertical center of target box (Y = 0.765)
    draw_arrow(center_x, 0.765, 0.675, 0.765, color='#b91c1c', shrinkA=0, shrinkB=3.0)

    draw_note(
        0.675, 0.725, 0.305, 0.080,
        "Excluded: age < 18\nyears (n = 23 262)",
        color='#b91c1c',
        fontsize=7.8,
    )

    draw_box(
        main_x, y2, main_w, main_h,
        "Adults aged ≥18 years\nN = 36 580",
    )
    # Vertical arrow with proper shrink
    draw_arrow(center_x, y2, center_x, y3 + main_h, shrinkA=1.5, shrinkB=1.5)

    # Exclusions description branch arrow (dashed) - aligned to vertical center of target box (Y = 0.540)
    draw_arrow(center_x, 0.540, 0.675, 0.540, ls='--', shrinkA=0, shrinkB=3.0)

    draw_note(
        0.675, 0.470, 0.305, 0.140,
        "No further participant\nexclusions: missing MEC\nweight n = 0; missing\nrace/ethnicity n = 0.",
    )

    draw_box(
        main_x, y3, main_w, main_h,
        "Final analytic sample\nN = 36 580",
    )
    # Vertical arrow with proper shrink
    draw_arrow(center_x, y3, center_x, y4 + h4, shrinkA=1.5, shrinkB=1.5)

    # Missingness description branch arrow (dashed) - aligned to vertical center of target box (Y = 0.311)
    draw_arrow(center_x, 0.311, 0.675, 0.311, ls='--', shrinkA=0, shrinkB=3.0)

    draw_note(
        0.675, 0.252, 0.305, 0.118,
        "Variable-level missingness\nwas handled by MICE\n(m = 20); see S7 Table.",
    )

    draw_box(
        main_x, y4, main_w, h4,
        "Primary MICE analysis\nPHQ-9 total score and log(GBR)\nN = 36 580\nObserved log(GBR): n = 32 804;\nmissing: n = 3 776",
        fill='#f8fafc',
        fontsize=8.95,
    )

    output_dir = "results"
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, "figure6_flowchart.png")
    plt.savefig(out_path, dpi=450, bbox_inches='tight', facecolor='white')
    plt.close()

    copy_targets = [
        os.path.join("results", "publication_assets", "figure6_flowchart.png"),
        os.path.join("results", "tp_figures", "supp_figure1_flowchart.png"),
        os.path.join("submission_packages", "translational_psychiatry", "supplementary_figures", "supp_figure1_flowchart.png"),
    ]
    for target in copy_targets:
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.copy2(out_path, target)
    print(f"Flowchart successfully generated at: {out_path}")

if __name__ == "__main__":
    main()
