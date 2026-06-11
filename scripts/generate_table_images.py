import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from generate_manuscript import (
    get_latest_results_file,
    parse_markdown_table,
    build_table_2,
    build_table_3_subgroups,
    build_table_3,
    build_table_4
)

def markdown_to_dataframe(md_table_str):
    lines = md_table_str.strip().split('\n')
    # Filter out separator lines like | :--- | :---: |
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

def render_table_image(df, title, output_path, figsize=(10, 6), col_widths=None):
    fig, ax = plt.subplots(figsize=figsize)
    ax.axis('off')
    
    # Use Helvetica with fallbacks for Japanese via font.sans-serif
    plt.rcParams['font.sans-serif'] = ['Helvetica', 'Arial', 'Hiragino Sans', 'AppleGothic', 'sans-serif']
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.size'] = 10
    
    # Title - Elegant and minimal
    ax.set_title(title, fontsize=11, weight='bold', pad=20, loc='left', color='#2c3e50')
    
    # Create table
    tbl = ax.table(
        cellText=df.values,
        colLabels=df.columns,
        cellLoc='left',
        loc='center',
        colWidths=col_widths
    )
    
    # Style table
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8.5)
    tbl.scale(1.0, 1.6)  # Taller rows for a professional, breathing layout
    
    n_rows = len(df)
    border_color_main = '#2c3e50'   # Top/bottom thick rule color
    border_color_light = '#dcdde1'  # Internal thin rule color
    bg_section = '#f8f9fa'          # Subtle background for subgroup headers
    
    for (row_idx, col_idx), cell in tbl.get_celld().items():
        # Clean three-line table (booktabs) border configuration
        if row_idx == 0:
            cell.visible_edges = 'TB'
        elif row_idx == n_rows:
            cell.visible_edges = 'B'
        else:
            cell.visible_edges = ''
        
        # Check if this row is a section header (e.g. contains "**" or has empty values in other columns)
        is_section = False
        if row_idx > 0:
            row_values = df.iloc[row_idx - 1]
            first_val = str(row_values.iloc[0]).strip()
            if first_val.startswith('**') and first_val.endswith('**'):
                is_section = True
            elif col_idx == 0 and all(str(x).strip() == '' for x in row_values.iloc[1:]):
                is_section = True
        
        if row_idx == 0:
            # Table Header
            cell.set_facecolor('#ffffff')
            cell.set_text_props(weight='bold', color='#111111')
            cell.set_linewidth(1.5)  # Thick top/middle rule
            cell.set_edgecolor(border_color_main)
            cell.set_height(0.06)
        elif is_section:
            # Section/Group Header
            cell.set_facecolor(bg_section)
            cell.set_text_props(weight='bold', color='#2c3e50')
            cell.set_linewidth(0.5)
            cell.set_edgecolor(border_color_light)
            
            # Clean up markdown bold markers if present
            clean_text = cell.get_text().get_text().replace('**', '')
            cell.get_text().set_text(clean_text)
        else:
            # Normal Data Row
            cell.set_facecolor('#ffffff')
            cell.set_text_props(color='#2c3e50')
            cell.set_linewidth(0.5)
            
            if row_idx == n_rows:
                # Bottom-most line of the table is thicker (booktabs bottom rule)
                cell.set_edgecolor(border_color_main)
                cell.set_linewidth(1.5)
            else:
                cell.set_edgecolor(border_color_light)
                
            # Align non-label cells to center
            if col_idx > 0:
                cell.set_text_props(horizontalalignment='center')
                
            # Replace markdown/HTML formatting markers with clean unicode symbols
            text_obj = cell.get_text()
            orig_text = text_obj.get_text()
            clean_text = orig_text.replace('&ge;', '≥').replace('&le;', '≤')
            text_obj.set_text(clean_text)
            
    # Save figure
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Table image saved to: {output_path}")

def main():
    try:
        results_file = get_latest_results_file()
        print(f"Reading results from: {results_file}")
        
        output_dir = "results"
        os.makedirs(output_dir, exist_ok=True)
        
        # 1. Table 1
        print("Rendering Table 1...")
        table1_md = parse_markdown_table(results_file, "DESCRIPTIVE TABLE 1 BY LOG_RATIO QUARTILE")
        df1 = markdown_to_dataframe(table1_md)
        render_table_image(
            df1,
            "Table 1. Participant Characteristics by log(GBR) Quartile (NHANES 2007-2018)",
            os.path.join(output_dir, "table1_characteristics.png"),
            figsize=(11, 8.5),
            col_widths=[0.35, 0.13, 0.13, 0.13, 0.13, 0.13]
        )
        
        # 2. Table 2
        print("Rendering Table 2...")
        table2_md = build_table_2(results_file)
        df2 = markdown_to_dataframe(table2_md)
        render_table_image(
            df2,
            "Table 2. Primary and Sensitivity Analyses of log(GBR) in the Full Sample",
            os.path.join(output_dir, "table2_association.png"),
            figsize=(12, 6.5),
            col_widths=[0.36, 0.08, 0.26, 0.08, 0.11, 0.11]
        )
        
        # 3. Table 3
        print("Rendering Table 3...")
        table3_md = build_table_3_subgroups(results_file)
        df3 = markdown_to_dataframe(table3_md)
        render_table_image(
            df3,
            "Table 3. Subgroup Stratified Analyses of GBR and PHQ-9",
            os.path.join(output_dir, "table3_subgroups.png"),
            figsize=(13, 11),
            col_widths=[0.32, 0.08, 0.28, 0.08, 0.12, 0.12]
        )
        
        # 4. Table 4
        print("Rendering Table 4...")
        table4_md = build_table_3(results_file)
        df4 = markdown_to_dataframe(table4_md)
        render_table_image(
            df4,
            "Table 4. Specificity Analyses of GBR and Other Liver Enzymes (Standardized per 1 SD)",
            os.path.join(output_dir, "table4_specificity.png"),
            figsize=(12, 5.5),
            col_widths=[0.32, 0.08, 0.28, 0.08, 0.12, 0.12]
        )
        
        # 5. Table 5
        print("Rendering Table 5...")
        table5_md = build_table_4(results_file)
        df5 = markdown_to_dataframe(table5_md)
        render_table_image(
            df5,
            "Table 5. Independent and Additive Association of GBR and OBS (Standardized per 1 SD)",
            os.path.join(output_dir, "table5_gbr_obs.png"),
            figsize=(13, 5.5),
            col_widths=[0.32, 0.24, 0.08, 0.20, 0.06, 0.06, 0.04]
        )
        
        # 6. Supplementary Table 3 (hsCRP Sensitivity)
        print("Rendering Supplementary Table 3...")
        from generate_manuscript import build_table_hscrp_sensitivity
        table_hscrp_md = build_table_hscrp_sensitivity(results_file)
        df_hscrp = markdown_to_dataframe(table_hscrp_md)
        render_table_image(
            df_hscrp,
            "Supplementary Table 3. hs-CRP Sensitivity Analysis in cycles I/J",
            os.path.join(output_dir, "supp_table3_hscrp.png"),
            figsize=(11, 4.5),
            col_widths=[0.35, 0.12, 0.26, 0.12, 0.15]
        )
        
        # 7. Supplementary Table 4 (CCA vs MICE)
        print("Rendering Supplementary Table 4...")
        from generate_manuscript import build_table_cca_vs_mice
        table_ccamice_md = build_table_cca_vs_mice(results_file)
        df_ccamice = markdown_to_dataframe(table_ccamice_md)
        render_table_image(
            df_ccamice,
            "Supplementary Table 4. Complete Case Analysis (CCA) vs Multiple Imputation (MICE) Bias Assessment",
            os.path.join(output_dir, "supp_table4_ccamice.png"),
            figsize=(11, 4.5),
            col_widths=[0.25, 0.15, 0.12, 0.26, 0.12, 0.10]
        )
        
        print("All table images generated successfully.")
        
    except Exception as e:
        print(f"Error generating table images: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
