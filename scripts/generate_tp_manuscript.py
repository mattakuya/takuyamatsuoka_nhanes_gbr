import os
import glob
import re
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "results" / "manuscript_english.txt"
OUTPUT = ROOT / "results" / "manuscript_tp.md"

def strip_markdown_italics(text: str) -> str:
    # Remove single asterisks around statistical variables (e.g. *N*, *β*, *p*, *df*, *χ*, *α*, *m*, *q*)
    return re.sub(r'\*([a-zA-Zβqαχ²\-]+)\*', r'\1', text)

def replace_thousands_commas_with_spaces(text: str) -> str:
    """Replace thousands separator commas with spaces in numbers like 36,580 or 92,795.85."""
    pattern = re.compile(r'\b(\d{1,3}),(\d{3})\b')
    while True:
        new_text = pattern.sub(r'\1 \2', text)
        if new_text == text:
            break
        text = new_text
    return text

def renumber_references_by_citation_order(text: str) -> str:
    """Renumber numeric references in order of first citation."""
    split_match = re.search(r"\nReferences\n", text)
    if not split_match:
        return text

    body = text[:split_match.end()]
    references_block = text[split_match.end():].strip()
    ref_entries = {}
    for line in references_block.splitlines():
        match = re.match(r"^(\d+)\.\s+(.*)$", line.strip())
        if match:
            ref_entries[int(match.group(1))] = match.group(2).strip()

    if not ref_entries:
        return text

    old_to_new = {}
    ordered_old_numbers = []

    def is_citation_token(token: str) -> bool:
        return bool(re.fullmatch(r"\s*\d+\s*", token))

    def replace_citation(match: re.Match) -> str:
        content = match.group(1)
        parts = [part.strip() for part in content.split(",")]
        if not parts or not all(is_citation_token(part) for part in parts):
            return match.group(0)

        new_numbers = []
        for part in parts:
            old_number = int(part)
            if old_number not in ref_entries:
                return match.group(0)
            if old_number not in old_to_new:
                old_to_new[old_number] = len(old_to_new) + 1
                ordered_old_numbers.append(old_number)
            new_numbers.append(str(old_to_new[old_number]))
        new_numbers = sorted(new_numbers, key=lambda value: int(value))
        return "[" + ", ".join(new_numbers) + "]"

    citation_pattern = re.compile(r"\[((?:\s*\d+\s*,)*\s*\d+\s*)\]")
    body = citation_pattern.sub(replace_citation, body)

    for old_number in sorted(ref_entries):
        if old_number not in old_to_new:
            old_to_new[old_number] = len(old_to_new) + 1
            ordered_old_numbers.append(old_number)

    renumbered_refs = [
        f"{old_to_new[old_number]}. {ref_entries[old_number]}"
        for old_number in ordered_old_numbers
    ]
    return body.rstrip() + "\n" + "\n".join(renumbered_refs) + "\n"

def get_latest_results_file(results_dir=None):
    if results_dir is None:
        results_dir = str(ROOT / "results")
    pattern = os.path.join(results_dir, "nhanes_results_*.md")
    files = glob.glob(pattern)
    files = [f for f in files if "QUICK" not in f]
    if not files:
        raise FileNotFoundError(f"Results file matching 'nhanes_results_*.md' not found in {results_dir}")
    return max(files, key=os.path.getmtime)

def parse_markdown_table(file_path, header_title):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Locate the heading
    pattern = rf"##\s+{re.escape(header_title)}\s*\n\n_FDR family:.*_\n\n(\|[\s\S]*?)(?=\n\n##|\Z)"
    match = re.search(pattern, content)
    if not match:
        pattern_simple = rf"##\s+{re.escape(header_title)}\s*\n\n(\|[\s\S]*?)(?=\n\n##|\Z)"
        match = re.search(pattern_simple, content)
    
    if match:
        return match.group(1).strip()
    else:
        raise ValueError(f"Could not find table under heading: {header_title}")

def build_table_1(results_file):
    return parse_markdown_table(results_file, "DESCRIPTIVE TABLE 1 BY LOG_RATIO QUARTILE")

def build_table_2(results_file):
    main_table = parse_markdown_table(results_file, "MAIN ANALYSIS: PHQ-9 CONTINUOUS")
    logistic_table = parse_markdown_table(results_file, "SENSITIVITY: PHQ-9 ≥ 10 (Logistic, β on log-odds)")
    sensitivity_table = parse_markdown_table(results_file, "SENSITIVITY ANALYSIS")
    liver_safe_table = parse_markdown_table(results_file, "LIVER-SAFE (NO DILI) ANALYSIS")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Analysis Model / Target | N | Beta / OR [95% CI] | SE | P-value | Q-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    
    # 1. Primary WLS
    rows.append("| **Primary Analysis (WLS)** | | | | | |")
    main_row = extract_rows(main_table)[0].split('|')[1:-1]
    subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = [x.strip() for x in main_row]
    rows.append(f"| Full Sample (OLS PHQ-9) | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
    
    # 2. Logistic PHQ-9 >= 10
    rows.append("| **Sensitivity Analysis (Logistic)** | | | | | |")
    log_row = extract_rows(logistic_table)[0].split('|')[1:-1]
    subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = [x.strip() for x in log_row]
    rows.append(f"| Full Sample (Logistic PHQ-9 &ge; 10) | {n} | OR = 1.1670 [1.0775, 1.2639] | — | {p} | {q} |")
    
    # 3. Exclusion sensitivity
    rows.append("| **Sensitivity Analyses (Exclusion Models)** | | | | | |")
    for row_str in extract_rows(sensitivity_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    # 4. Liver-safe
    rows.append("| **Liver-Safe Analysis (No DILI)** | | | | | |")
    for row_str in extract_rows(liver_safe_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    return '\n'.join([header, separator] + rows)

def build_table_3(results_file):
    bmi_table = parse_markdown_table(results_file, "BMI STRATIFIED ANALYSIS")
    sex_table = parse_markdown_table(results_file, "SEX-STRATIFIED ANALYSIS")
    age_table = parse_markdown_table(results_file, "AGE-STRATIFIED ANALYSIS")
    race_table = parse_markdown_table(results_file, "RACE-STRATIFIED ANALYSIS")
    alcohol_table = parse_markdown_table(results_file, "ALCOHOL-STRATIFIED ANALYSIS (NIAAA)")
    cotinine_table = parse_markdown_table(results_file, "COTININE-STRATIFIED ANALYSIS (Pirkle)")
    crp_table = parse_markdown_table(results_file, "CRP-STRATIFIED ANALYSIS (AHA, I/J only)")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Subgroup / Stratification | N | Beta [95% CI] | SE | P-value | Q-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    
    def append_group(title, table_str, skip_full=True):
        rows.append(f"| **{title}** | | | | | |")
        for row_str in extract_rows(table_str):
            parts = [x.strip() for x in row_str.split('|')[1:-1]]
            subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
            if skip_full and "Full Sample" in subgroup:
                continue
            rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
            
    append_group("BMI Categories (kg/m²)", bmi_table)
    append_group("Sex", sex_table)
    append_group("Age Groups", age_table)
    append_group("Race / Ethnicity", race_table)
    append_group("Alcohol Consumption (NIAAA)", alcohol_table)
    append_group("Cotinine Level (Smoking Exposure)", cotinine_table)
    append_group("Systemic Inflammation (hs-CRP, I/J only)", crp_table)
    
    return '\n'.join([header, separator] + rows)

def build_table_4(results_file):
    obs_table = parse_markdown_table(results_file, "JOINT MODEL: log_ratio + OBS (per 1 SD each, SD-standardized)")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Subgroup / Model | Term | N | Beta [95% CI] | SE | P-value | Attenuation (%) |"
    separator = "| :--- | :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    for row_str in extract_rows(obs_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        if len(parts) == 13:
            subgroup, model, term_part1, term_part2, n, beta, se, ci_lo, ci_hi, p, atten, q, q_fam = parts
            term = f"{term_part1} + {term_part2}"
        else:
            subgroup, model, term, n, beta, se, ci_lo, ci_hi, p, atten, q, q_fam = parts
        rows.append(f"| {subgroup} ({model}) | {term} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {atten} |")
        
    return '\n'.join([header, separator] + rows)

def build_supp_table_1(results_file):
    return parse_markdown_table(results_file, "VIF (Main Model Covariates)")

def build_supp_table_2(results_file):
    model_fit_table = parse_markdown_table(results_file, "MODEL FIT COMPARISON: RATIO vs GGT vs GGT+BIL")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Model | N | Parameters | Pseudo-AIC | Pseudo-BIC | Adjusted R² | Weighted RMSE |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    for row_str in extract_rows(model_fit_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        model, n, m, aic_mean, aic_sd, bic_mean, bic_sd, r2_mean, r2_sd, rmse_mean, rmse_sd = parts
        rows.append(f"| {model} | {n} | {m} | {float(aic_mean):.2f} | {float(bic_mean):.2f} | {float(r2_mean):.4f} | {float(rmse_mean):.4f} |")
        
    return '\n'.join([header, separator] + rows)

def build_supp_table_3(results_file):
    return parse_markdown_table(results_file, "CCA vs MICE: GGT BIAS ASSESSMENT (SD-standardized)")

def build_supp_table_4(results_file):
    spec_table = parse_markdown_table(results_file, "OXIDATIVE COMPARISON: GGT/Bil vs CDAI vs OBS (SD-STANDARDIZED)")
    excl_ad_table = parse_markdown_table(results_file, "OXIDATIVE COMPARISON EXCL. AD USERS (SD-STANDARDIZED)")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Analysis Cohort / Indicator | N | Beta [95% CI] | SE | P-value | Q-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    rows.append("| **Full Cohort (OBS-Restricted Sample)** | | | | | |")
    for row_str in extract_rows(spec_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    rows.append("| **Antidepressant Non-Users Cohort** | | | | | |")
    for row_str in extract_rows(excl_ad_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    return '\n'.join([header, separator] + rows)

def build_supp_table_5(results_file):
    ad_table = parse_markdown_table(results_file, "AD-INTERACTION ANALYSIS [log_ratio]")
    ad_normal_table = parse_markdown_table(results_file, "AD-INTERACTION ANALYSIS [log_ratio] (Normal BMI)")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Sample / Term | N | Beta [95% CI] | SE | P-value | Q-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    rows.append("| **Full Cohort Interaction Model** | | | | | |")
    for row_str in extract_rows(ad_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        term, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {term} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    rows.append("| **Normal BMI (18.5–24.9 kg/m²) Cohort** | | | | | |")
    for row_str in extract_rows(ad_normal_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        term, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {term} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    return '\n'.join([header, separator] + rows)

def build_supp_table_6(race_results_file):
    race_interaction = parse_markdown_table(race_results_file, "RACE/ETHNICITY INTERACTION ANALYSIS [log_ratio]")
    global_test = parse_markdown_table(race_results_file, "GLOBAL RACE/ETHNICITY INTERACTION TEST [log_ratio]")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Term / Test | N | Beta / Chi-square [95% CI / df] | SE | P-value | Q-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    rows.append("| **Interaction Model (NHWhite as Reference)** | | | | | |")
    for row_str in extract_rows(race_interaction):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        term, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {term} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    rows.append("| **Global Interaction Wald Test** | | | | | |")
    for row_str in extract_rows(global_test):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        test, n, chi2, df, p, terms, q, q_fam = parts
        rows.append(f"| {test} | {n} | Chi² = {float(chi2):.4f} [df = {df}] | — | {p} | {q} |")
        
    return '\n'.join([header, separator] + rows)

def apply_fonts_to_run(run, bold=False, size_pt=11):
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Helvetica')
    rFonts.set(qn('w:hAnsi'), 'Helvetica')
    rPr.append(rFonts)
    
    run.font.size = Pt(size_pt)
    if bold:
        run.bold = True

def apply_three_line_table_styling(table, rows_data):
    """
    Applies professional booktabs (three-line table) styling to a python-docx table.
    - No vertical lines.
    - Thick top rule (1.5 pt = 12 sz) and bottom rule (1.5 pt = 12 sz).
    - Thin header rule below row 0 (0.75 pt = 6 sz).
    - Cell padding (5 pt top/bottom, 7.5 pt left/right).
    - Merged cells and light gray shading (#F2F2F2) for group header rows.
    """
    tblPr = table._tbl.tblPr
    
    # 1. Clear standard borders by adding custom w:tblBorders
    tblBorders = OxmlElement('w:tblBorders')
    
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '12')
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), '111111')
    tblBorders.append(top)
    
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '111111')
    tblBorders.append(bottom)
    
    for border_name in ('left', 'right', 'insideV', 'insideH'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
        
    tblPr.append(tblBorders)
    
    # 2. Set table-level cell padding (dxa unit)
    tblCellMar = OxmlElement('w:tblCellMar')
    for margin_name, val in [('top', 100), ('bottom', 100), ('left', 150), ('right', 150)]:
        margin = OxmlElement(f'w:{margin_name}')
        margin.set(qn('w:w'), str(val))
        margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(margin)
    tblPr.append(tblCellMar)
    
    # 3. Render and style cells
    for r_idx, row_data in enumerate(rows_data):
        row = table.rows[r_idx]
        
        # Check if this is a group header row (all bold or empty other cells)
        is_group_header = False
        first_val = row_data[0].strip()
        if r_idx > 0 and (first_val.startswith('**') and first_val.endswith('**')):
            is_group_header = True
        elif r_idx > 0 and all(cell.strip() == '' for cell in row_data[1:]):
            is_group_header = True
            
        if is_group_header:
            # Merge all cells in this row
            merged_cell = row.cells[0]
            for c_idx in range(1, len(row.cells)):
                merged_cell = merged_cell.merge(row.cells[c_idx])
            
            # Style the merged group header cell
            merged_cell.text = ""
            p = merged_cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(first_val.replace('**', ''))
            apply_fonts_to_run(run, bold=True, size_pt=10)
            
            # Subtle background color (#F2F2F2)
            shading = OxmlElement('w:shd')
            shading.set(qn('w:val'), 'clear')
            shading.set(qn('w:color'), 'auto')
            shading.set(qn('w:fill'), 'F2F2F2')
            merged_cell._tc.get_or_add_tcPr().append(shading)
            continue
            
        for c_idx, cell_value in enumerate(row_data):
            if c_idx < len(row.cells):
                cell = row.cells[c_idx]
                cell.text = ""
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
                
                # Alignments
                if c_idx == 0:
                    p.alignment = 0  # Left
                else:
                    p.alignment = 1  # Center
                
                is_header = (r_idx == 0)
                clean_val = cell_value.replace("**", "").replace("&ge;", "≥").replace("&le;", "≤")
                
                run = p.add_run(clean_val)
                apply_fonts_to_run(run, bold=is_header, size_pt=10)
                
                # Add thin bottom border to header row (row 0)
                if r_idx == 0:
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcBorders = OxmlElement('w:tcBorders')
                    bottom = OxmlElement('w:bottom')
                    bottom.set(qn('w:val'), 'single')
                    bottom.set(qn('w:sz'), '6')  # 0.75 pt
                    bottom.set(qn('w:space'), '0')
                    bottom.set(qn('w:color'), '111111')
                    tcBorders.append(bottom)
                    tcPr.append(tcBorders)

def convert_md_to_docx(md_path, docx_path):
    print(f"Converting {md_path} to {docx_path}...")
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.read().split('\n')
        
    in_table = False
    table_lines = []
    
    def flush_table():
        nonlocal table_lines
        if not table_lines:
            return
        
        rows_data = []
        for line in table_lines:
            if not line.strip() or re.match(r'^\s*\|\s*[:\-|\s]+$', line):
                continue
            parts = [cell.strip() for cell in line.strip().split('|')[1:-1]]
            rows_data.append(parts)
            
        if not rows_data:
            table_lines = []
            return
            
        num_cols = max(len(row) for row in rows_data)
        num_rows = len(rows_data)
        
        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.style = None
        apply_three_line_table_styling(table, rows_data)
        
        doc.add_paragraph()
        table_lines = []
        
    i = 0
    while i < len(lines):
        line = lines[i]
        
        if line.strip().startswith('|'):
            in_table = True
            table_lines.append(line)
            i += 1
            continue
        elif in_table:
            flush_table()
            in_table = False
            
        if not line.strip():
            i += 1
            continue
            
        h1_match = re.match(r'^#\s+(.*)', line)
        if h1_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(h1_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=16)
            i += 1
            continue
            
        h2_match = re.match(r'^##\s+(.*)', line)
        if h2_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(h2_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=13)
            i += 1
            continue
            
        h3_match = re.match(r'^###\s+(.*)', line)
        if h3_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(h3_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=11)
            i += 1
            continue
            
        h4_match = re.match(r'^####\s+(.*)', line)
        if h4_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(h4_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=10.5)
            i += 1
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        
        parts = re.split(r'(\*\*.*?\*\*)', line)
        for part in parts:
            if not part:
                continue
            if part.startswith('**') and part.endswith('**'):
                bold_text = part[2:-2]
                run = p.add_run(bold_text)
                apply_fonts_to_run(run, bold=True, size_pt=11)
            else:
                run = p.add_run(part)
                apply_fonts_to_run(run, bold=False, size_pt=11)
                
        i += 1
        
    if in_table:
        flush_table()
        
    doc.save(docx_path)
    print(f"Successfully saved Word file to: {docx_path}")

def replace_once(text: str, old: str, new: str) -> str:
    if old not in text:
        raise ValueError(f"Target block not found:\n{old[:300]}")
    return text.replace(old, new, 1)

def main() -> None:
    results_file = get_latest_results_file()
    race_results_file = results_file
    print(f"Parsing raw results from: {results_file}")

    # Build main tables
    table1 = build_table_1(results_file)
    table2 = build_table_2(results_file)
    table3 = build_table_3(results_file)
    table4 = build_table_4(results_file)

    # Build supplementary tables
    supp_table1 = build_supp_table_1(results_file)
    supp_table2 = build_supp_table_2(results_file)
    supp_table3 = build_supp_table_3(results_file)
    supp_table4 = build_supp_table_4(results_file)
    supp_table5 = build_supp_table_5(results_file)
    supp_table6 = build_supp_table_6(race_results_file)
    missingness_path = ROOT / "results" / "supp_table7_missingness.md"
    supp_table7 = (
        missingness_path.read_text(encoding="utf-8").strip()
        if missingness_path.exists()
        else "_Run generate_missingness_table.py to create this table._"
    )

    # Read original draft template
    text = SOURCE.read_text(encoding="utf-8")

    # Perform manuscript edits
    text = replace_once(
        text,
        "Association of the gamma-glutamyl transferase-to-bilirubin ratio with depressive symptoms in US adults: A cross-sectional study of NHANES 2007–2018",
        "Race- and ethnicity-specific association of the gamma-glutamyl transferase-to-bilirubin ratio with depressive symptoms in US adults: a NHANES 2007–2018 study\n\nRunning title: Redox balance and depressive symptoms",
    )
    text = replace_once(
        text,
        "¹ [Affiliation Placeholder: Department, Institution, City, Country]",
        "¹ Tohoku University Hospital, Sendai, Japan",
    )
    text = replace_once(
        text,
        "E-mail: [Email Address Placeholder]",
        "E-mail: takuya.matsuoka.c6@tohoku.ac.jp",
    )

    old_abstract = """Abstract
Oxidative stress is implicated in the pathophysiology of depression. Gamma-glutamyl transferase (GGT) and bilirubin have been proposed as markers related to pro-oxidant burden and antioxidant capacity, respectively. In this study, we evaluated the association of the GGT-to-total bilirubin ratio (GBR), a putative proxy related to systemic redox balance, with depressive symptoms. We analyzed cross-sectional data from the National Health and Nutrition Examination Survey (NHANES) 2007–2018 for adults aged ≥18 years (*N* = 36,580). Missing values were imputed using multiple imputation by chained equations (*m* = 20), and pooled weighted least squares regression was performed. The primary outcome was the Patient Health Questionnaire-9 (PHQ-9) score, and the exposure was log(GBR). In the fully adjusted model, log(GBR) was significantly associated with higher PHQ-9 scores (*β* = 0.256, 95% confidence interval [CI]: 0.171–0.341, *p* = 4.8 × 10⁻⁸). Logistic sensitivity analysis with PHQ-9 ≥ 10 as the outcome supported this association (odds ratio [OR] = 1.167, 95% CI: 1.078–1.264). The association was generally consistent across subgroup and sensitivity analyses. In race/ethnicity-specific analysis, positive associations were observed in Non-Hispanic Whites and Non-Hispanic Blacks, but not in Hispanics or Non-Hispanic Asians. GBR, a simple composite redox-related proxy calculated from routine biochemical panels, is independently associated with depressive symptoms. Although the individual-level effect size is modest, GBR may serve as a feasible epidemiological marker for investigating redox-related biological pathways in depressive symptoms. Longitudinal studies are needed to clarify causality, temporal stability, and potential research or clinical relevance.
"""
    new_abstract = """Abstract
Oxidative stress is implicated in the pathophysiology of depression, but scalable biomarkers that capture redox-related heterogeneity in population settings remain limited. Gamma-glutamyl transferase (GGT) and bilirubin have been proposed as markers related to pro-oxidant burden and endogenous antioxidant capacity, respectively. We examined whether the GGT-to-total bilirubin ratio (GBR), a routine biochemistry-derived redox-related composite, is associated with depressive symptoms and whether this association differs by race/ethnicity. We analyzed cross-sectional data from the National Health and Nutrition Examination Survey (NHANES) 2007–2018 for adults aged ≥18 years (*N* = 36,580). Missing values were imputed using multiple imputation by chained equations (*m* = 20), and pooled weighted least squares regression was performed. The primary outcome was the Patient Health Questionnaire-9 (PHQ-9) score, and the exposure was mean-centered log(GBR). In the fully adjusted model, log(GBR) was associated with higher PHQ-9 scores (*β* = 0.256, 95% confidence interval [CI]: 0.171–0.341, *p* = 4.8 × 10⁻⁸). Logistic sensitivity analysis with PHQ-9 ≥ 10 as the outcome supported this association (odds ratio [OR] = 1.167, 95% confidence interval [CI]: 1.078–1.264). The association persisted after exclusion of antidepressant users, participants with cardiovascular disease, participants with diabetes, and participants with elevated ALT/AST. Race/ethnicity-stratified analyses showed positive associations in Non-Hispanic White and Non-Hispanic Black participants, but not in Hispanic or Non-Hispanic Asian participants. A formal global interaction test supported heterogeneity by race/ethnicity (χ² = 15.05, df = 5, *p* = 0.010). These findings suggest that GBR is associated with depressive symptom burden, but its validity as a redox-related epidemiological marker may not be population-invariant. External validation in Asian and Hispanic populations, ideally with genetic and metabolic characterization, is needed.
"""
    text = replace_once(text, old_abstract, new_abstract)

    old_intro_end = """The purpose of this study was to examine whether the GGT-to-total bilirubin ratio (hereafter GBR), a putative redox-related composite proxy derived from routine biochemical measurements, is associated with depressive symptoms. We also compared GBR with other oxidative stress-related markers and indices (GGT alone, T-Bil alone, OBS, and CDAI) and evaluated its specificity and heterogeneity across populations through subgroup analyses stratified by BMI, sex, and race/ethnicity.
"""
    new_intro_end = """The purpose of this study was to examine whether the GGT-to-total bilirubin ratio (hereafter GBR), a putative redox-related composite proxy derived from routine biochemical measurements, is associated with depressive symptoms. We hypothesized that GBR would be associated with depressive symptom burden independently of metabolic, inflammatory, lifestyle, and medication-related factors, and that this association might differ across population subgroups in which bilirubin and liver-metabolic traits have distinct genetic and metabolic backgrounds. We also compared GBR with other oxidative stress-related markers and indices (GGT alone, T-Bil alone, OBS, and CDAI) and evaluated its specificity and heterogeneity across populations through subgroup analyses and formal interaction testing.
"""
    text = replace_once(text, old_intro_end, new_intro_end)

    old_methods = """Subgroup analyses were conducted stratified by BMI categories (18.5 ≤ BMI < 25.0, 25.0 ≤ BMI < 30.0, BMI ≥ 30.0 kg/m²), sex, and race/ethnicity (Non-Hispanic White, Non-Hispanic Black, Hispanic [combined, Mexican American, and Other Hispanic], and Non-Hispanic Asian). To reflect the stratum-specific missingness patterns, MICE was performed separately within each stratum for subgroup analyses.
"""
    new_methods = """Subgroup analyses were conducted stratified by BMI categories (18.5 ≤ BMI < 25.0, 25.0 ≤ BMI < 30.0, BMI ≥ 30.0 kg/m²), sex, and race/ethnicity (Non-Hispanic White, Non-Hispanic Black, Hispanic [combined, Mexican American, and Other Hispanic], and Non-Hispanic Asian). To reflect the stratum-specific missingness patterns, MICE was performed separately within each stratum for subgroup analyses. To formally evaluate population heterogeneity, we additionally fitted an interaction model including log(GBR) × race/ethnicity terms, using Non-Hispanic White participants as the reference group, and tested all non-reference interaction terms jointly with a pooled Wald chi-square test across imputed datasets.
"""
    text = replace_once(text, old_methods, new_methods)

    old_race_results = """By race/ethnicity, positive associations were identified in Non-Hispanic White (*N* = 14,561, *β* = 0.352, *p* = 2.4 × 10⁻⁷) and Non-Hispanic Black (*N* = 7,912, *β* = 0.278, *p* = 8.8 × 10⁻⁴) populations (Fig 4). However, no statistically significant associations were observed in the Hispanic combined cohort (*N* = 9,487, *β* = -0.024, *p* = 0.749), Mexican American, Other Hispanic, or Non-Hispanic Asian populations. This racial/ethnic pattern was consistently maintained even after excluding antidepressant users.
"""
    new_race_results = """By race/ethnicity, positive associations were identified in Non-Hispanic White (*N* = 14,561, *β* = 0.352, *p* = 2.4 × 10⁻⁷) and Non-Hispanic Black (*N* = 7,912, *β* = 0.278, *p* = 8.8 × 10⁻⁴) populations (Fig 4). However, no statistically significant associations were observed in the Hispanic combined cohort (*N* = 9,487, *β* = -0.024, *p* = 0.749), Mexican American, Other Hispanic, or Non-Hispanic Asian populations. This racial/ethnic pattern was consistently maintained even after excluding antidepressant users. In the formal interaction model using Non-Hispanic White participants as the reference group, the global log(GBR) × race/ethnicity interaction was statistically significant (χ² = 15.05, df = 5, *p* = 0.010; q = 0.024), supporting heterogeneity across racial/ethnic groups. The strongest negative interaction contrasts relative to Non-Hispanic White participants were observed for Mexican American participants (*β*interaction = -0.293, *p* = 0.0065; q = 0.023) and Non-Hispanic Asian participants (*β*interaction = -0.207, *p* = 0.030; q = 0.053).
"""
    text = replace_once(text, old_race_results, new_race_results)

    old_discussion_open = """This cross-sectional study using US NHANES data demonstrated that the serum log(GBR) is independently associated with depressive symptoms (PHQ-9 scores). In the primary analysis, a 1-unit increase in log(GBR) was associated with a 0.256-point higher PHQ-9 total score, and binary outcome analysis using PHQ-9 ≥ 10 showed a consistent association in the same direction. This association was generally sustained across multiple sensitivity analyses that accounted for major comorbidities, antidepressant use, smoking exposure, and clinically elevated liver enzymes. Although the effect size is modest at the individual level, GBR may be useful primarily as an epidemiological and mechanistic research marker derived from routine biochemical tests, rather than as an individual-level clinical screening tool.
"""
    new_discussion_open = """This cross-sectional study using US NHANES data demonstrated that serum log(GBR) is independently associated with depressive symptoms (PHQ-9 scores), and that this association differs by race/ethnicity. In the primary analysis, a 1-unit increase in log(GBR) was associated with a 0.256-point higher PHQ-9 total score, and binary outcome analysis using PHQ-9 ≥ 10 showed a consistent association in the same direction. This association was generally sustained across multiple sensitivity analyses that accounted for major comorbidities, antidepressant use, smoking exposure, and clinically elevated liver enzymes. Importantly, the race/ethnicity-specific pattern was supported not only by stratified estimates but also by a formal global interaction test. Although the effect size is modest at the individual level, GBR may be useful primarily as an epidemiological and mechanistic research marker derived from routine biochemical tests, rather than as an individual-level clinical screening tool.
"""
    text = replace_once(text, old_discussion_open, new_discussion_open)

    old_race_discussion = """Racial/ethnic specificity
The association between GBR and depressive symptoms was distinct in Non-Hispanic White (NHW) and Non-Hispanic Black (NHB) populations but absent in Hispanic and Asian populations. This trend was consistently observed even after adjusting for major confounders, such as alcohol consumption, smoking history, BMI, diabetes, CVD, and statin use, and in sensitivity analyses excluding antidepressant users. Therefore, this racial/ethnic difference may not be fully explained by measured clinical and lifestyle factors alone. Examining GBR components individually revealed that in NHW and NHB cohorts, GGT was positively and T-Bil was negatively associated with depressive symptoms, resulting in a positive association for GBR, consistent with our hypothesis. Several speculative biological mechanisms may potentially contribute to these discrepant findings. First, differences in genetic backgrounds related to bilirubin metabolism, particularly polymorphisms in the _UGT1A1_ gene, may play a role. The _UGT1A1_*28 allele, which dictates Gilbert syndrome characterized by elevated serum T-Bil, is highly prevalent in European White and Black populations (allele frequency: 30%–40%) but less common in East Asian and some Hispanic populations (10%–15%) [14]. In Asian populations, the _UGT1A1_*6 polymorphism commonly dictates the Gilbert phenotype instead [28]; differences in the distribution of these genetic variants might alter the biological significance of bilirubin as a variable. Second, specific metabolic features and susceptibility to fatty liver disease in Hispanic populations may contribute. The Hispanic population exhibits a high allele frequency of the _PNPLA3_ I148M variant associated with lipid droplet accumulation [29], predisposing them to develop non-alcoholic fatty liver disease (NAFLD) even in the absence of obesity or other liver diseases. This metabolic vulnerability may potentially influence GGT elevation and bilirubin dynamics, thereby altering the physiological interpretation of GBR across populations. These mechanistic interpretations remain hypothesis-generating because genetic polymorphisms and related metabolic pathways were not directly assessed in the present dataset.

These findings suggest that GBR may not represent a universal indicator with identical physiological implications across all populations. To confirm the ethnic heterogeneity of GBR's association with depression, external validation in Hispanic and Asian populations is indispensable. In particular, validation in independent cohorts that integrate _UGT1A1_-related genetic polymorphisms, insulin resistance, and other lifestyle or environmental exposures is an important future direction to clarify the potential research utility of this index.
"""
    new_race_discussion = """Racial/ethnic specificity
One of the most notable findings was the racial/ethnic heterogeneity of the association between GBR and depressive symptoms. The association was clear in Non-Hispanic White (NHW) and Non-Hispanic Black (NHB) participants but absent in Hispanic and Asian participants, and this pattern was supported by a formal global interaction test (χ² = 15.05, df = 5, *p* = 0.010). The pattern also persisted after adjustment for major confounders, including alcohol consumption, smoking history, BMI, diabetes, CVD, and statin use, and in sensitivity analyses excluding antidepressant users. Therefore, this racial/ethnic difference may not be fully explained by measured clinical and lifestyle factors alone.

The component-specific analyses provide a possible clue. In NHW and NHB participants, the direction of the component estimates was consistent with the redox-balance hypothesis: GGT showed positive associations, bilirubin showed inverse associations, and the integrated GBR showed positive associations with depressive symptoms. In contrast, the Hispanic and Asian subgroups did not show the same component pattern, suggesting that the physiological meaning of bilirubin, GGT, or their ratio may differ across populations. This is important because it suggests that the validity of routine redox-related biomarkers may not be population-invariant.

Several speculative biological mechanisms may contribute to these discrepant findings. First, differences in genetic backgrounds related to bilirubin metabolism, particularly polymorphisms in the _UGT1A1_ gene, may play a role. The _UGT1A1_*28 allele, which contributes to the Gilbert syndrome phenotype characterized by elevated serum bilirubin, is more prevalent in European White and Black populations but less common in East Asian and some Hispanic populations [14]. In Asian populations, the _UGT1A1_*6 polymorphism more commonly contributes to the Gilbert phenotype [28]. Differences in the distribution of these genetic variants may alter the biological significance of bilirubin as an antioxidant-related marker. Second, specific metabolic features and susceptibility to fatty liver disease in Hispanic populations may contribute. Hispanic populations have a high frequency of the _PNPLA3_ I148M variant associated with lipid droplet accumulation [29], which may increase susceptibility to non-alcoholic fatty liver disease and influence GGT elevation and bilirubin dynamics. This metabolic vulnerability may alter the physiological interpretation of GBR across populations.

These mechanistic interpretations remain hypothesis-generating because genetic polymorphisms and related metabolic pathways were not directly assessed in the present dataset. Nonetheless, the findings raise a broader point for translational biomarker research: a biomarker calculated from routine clinical chemistry may be easy to implement, but its biological meaning may still depend on population-specific genetic and metabolic context. To confirm the ethnic heterogeneity of GBR's association with depressive symptoms, external validation in Hispanic and Asian populations is therefore urgent. In particular, validation in independent cohorts that integrate _UGT1A1_-related genetic polymorphisms, insulin resistance, liver fat or NAFLD markers, and other lifestyle or environmental exposures is an important future direction to clarify the potential research utility of this index.
"""
    text = replace_once(text, old_race_discussion, new_race_discussion)

    old_conclusion = """In this cross-sectional study using a nationally representative sample from the US NHANES, we demonstrated that the serum GGT-to-total bilirubin ratio (GBR) is independently associated with depressive symptoms, and that this association may vary by race/ethnicity. Calculated from routine biochemistry measures, GBR is a highly feasible candidate composite marker for epidemiological research related to systemic redox biology. However, caution should be exercised regarding the cross-sectional design, the relatively small effect size, and the limited superiority of GBR over GGT alone in predictive performance. Important future research directions include: (1) external validation in independent cohorts including Asian and Hispanic populations; (2) stratified analyses incorporating genetic polymorphisms related to bilirubin and glutathione metabolic pathways; and (3) longitudinal evaluation of GBR's predictive performance for depression. Given the cross-sectional design and modest effect size, the present findings should primarily be interpreted as hypothesis-generating rather than clinically predictive. These investigations will help evaluate the utility of GBR as a complementary marker in research on depressive symptoms based on systemic redox biology.
"""
    new_conclusion = """In this cross-sectional study using a nationally representative sample from the US NHANES, we demonstrated that the serum GGT-to-total bilirubin ratio (GBR) is independently associated with depressive symptoms, and that this association differs by race/ethnicity. Calculated from routine biochemistry measures, GBR is a highly feasible candidate composite marker for epidemiological research related to systemic redox biology. However, caution should be exercised regarding the cross-sectional design, the relatively small effect size, and the limited superiority of GBR over GGT alone in predictive performance. The present findings should primarily be interpreted as hypothesis-generating rather than clinically predictive. Important future research directions include: (1) external validation in independent cohorts including Asian and Hispanic populations; (2) stratified analyses incorporating genetic polymorphisms related to bilirubin and glutathione metabolic pathways; and (3) longitudinal evaluation of GBR's temporal stability and predictive performance for depression. More broadly, these results highlight that the translation of routine redox-related biomarkers into psychiatric epidemiology may require explicit attention to population-specific biological context.
"""
    text = replace_once(text, old_conclusion, new_conclusion)

    old_si = """**S5 Table. Antidepressant interaction analyses.** Detailed regression coefficients and interaction term estimates for the joint and BMI-stratified models exploring the modification of GBR and GGT associations by antidepressant use.
"""
    new_si = """**S5 Table. Antidepressant interaction analyses.** Detailed regression coefficients and interaction term estimates for the joint and BMI-stratified models exploring the modification of GBR and GGT associations by antidepressant use.
**S6 Table. Race/ethnicity interaction analysis.** Pooled interaction coefficients and global Wald test for log(GBR) × race/ethnicity, using Non-Hispanic White participants as the reference group.
**S7 Table. Missingness of primary analysis variables before multiple imputation.** Variable-level availability and missingness in the analytic adult sample before MICE.
"""
    text = replace_once(text, old_si, new_si)

    old_declarations = """Competing interests
The author has declared that no competing interests exist.

Data availability
The dataset analyzed in the current study is publicly available from the National Health and Nutrition Examination Survey (NHANES) website hosted by the Centers for Disease Control and Prevention (CDC) at https://www.cdc.gov/nchs/nhanes/index.htm. All analysis scripts used to generate the results are available at https://github.com/mattakuya/NHANES_GBR.
"""
    new_declarations = """Competing interests
The author has declared that no competing interests exist.

Use of generative AI and AI-assisted technologies
During manuscript preparation, the author used OpenAI's ChatGPT/Codex for language editing, drafting assistance, and code review. The author reviewed, revised, and approved all AI-assisted outputs and takes full responsibility for the final content of the manuscript.

Data availability
The datasets analyzed during the current study are publicly available from the National Health and Nutrition Examination Survey (NHANES) repository (https://www.cdc.gov/nchs/nhanes/index.htm).

Code availability
All R and Python scripts used for data download, processing, statistical analysis, and figure generation are available in the GitHub repository at https://github.com/mattakuya/NHANES_GBR.
"""
    text = replace_once(text, old_declarations, new_declarations)

    # Normalize section headings to Nature style
    text = text.replace("Materials and methods", "Methods", 1)
    text = text.replace("Supporting information", "Supplementary Information", 1)

    # Insert populated tables into manuscript
    text = replace_once(
        text,
        "**Table 1. Characteristics of the study cohort stratified by quartiles of the GGT-to-total bilirubin ratio (GBR).**",
        "**Table 1. Participant characteristics by quartiles of log(GBR).** Values are unweighted counts, means (SD), or weighted percentages, as indicated.\n\n[Insert Table 1 here]"
    )
    text = replace_once(
        text,
        "**Table 2. Association between GBR and depressive symptoms in primary and sensitivity analyses.**",
        "**Table 2. Primary and sensitivity analyses of the association between log(GBR) and depressive symptoms.** Estimates are from fully adjusted weighted models unless otherwise indicated.\n\n[Insert Table 2 here]"
    )
    text = replace_once(
        text,
        "**Table 3. Subgroup analyses of the association between GBR and depressive symptoms.**",
        "**Table 3. Subgroup analyses of the association between GBR and depressive symptoms.**\n\n[Insert Table 3 here]"
    )
    text = replace_once(
        text,
        "**Table 4. Comparison of GBR with established oxidative stress indicators and joint models.**",
        "**Table 4. Comparison of GBR with established oxidative stress indicators and joint models.**\n\n[Insert Table 4 here]"
    )

    # Strip markdown-like asterisks around statistical terms
    text = strip_markdown_italics(text)

    # Translational Psychiatry display-item allocation:
    # main items are Table 1, Table 2, Fig 1 race/ethnicity, Fig 2 specificity,
    # and Fig 3 RCS. Broad subgroup, flowchart, AD interaction, and OBS joint
    # analyses are kept as supplementary display items.
    display_replacements = [
        (
            "Adults aged 18 years and older were extracted from the dataset. Those with missing Mobile Examination Center (MEC) exam weights were excluded from the analysis. The final analysis cohort consisted of N = 36,580 participants (Fig 1).",
            "Adults aged 18 years and older were extracted from the dataset. Those with missing Mobile Examination Center (MEC) exam weights were excluded from the analysis. The final analysis cohort consisted of N = 36,580 participants (S1 Fig).",
        ),
        (
            "**Fig 1. Flowchart of participant selection.** Detailed selection process and exclusion criteria applied to the NHANES 2007–2018 cycles to establish the final study cohort (N = 36,580).",
            "**S1 Fig. Participant selection flowchart.** Selection of eligible adults from NHANES 2007–2018.",
        ),
        (
            "(Fig 1)",
            "(Fig. 1)",
        ),
        (
            "(Fig 2)",
            "(Fig. 2)",
        ),
        (
            "(Fig 3)",
            "(Fig. 3)",
        ),
        (
            "The subgroup analysis results are reported in Table 3.\n\n**Table 3. Subgroup analyses of the association between GBR and depressive symptoms.**\n\n[Insert Table 3 here]",
            "Additional subgroup visualizations are provided in Supplementary Figs. S2 and S3 and summarized below.",
        ),
        (
            "with concurrent NLR and hsCRP adjustment (β = 0.178, p = 0.020; S2 Table).",
            "with concurrent NLR and hsCRP adjustment (β = 0.178, p = 0.020).",
        ),
        (
            "By BMI categories, the association between log(GBR) and PHQ-9 was observed in all strata: normal weight (18.5–24.9 kg/m²; N = 9,783, β = 0.301, p = 1.1 × 10⁻⁴), overweight (25.0–29.9 kg/m²; N = 11,205, β = 0.328, p = 3.8 × 10⁻⁶), and obese (≥ 30.0 kg/m²; N = 13,007, β = 0.162, p = 0.015) (Fig 3).",
            "By BMI categories, the association between log(GBR) and PHQ-9 was observed in all strata: normal weight (18.5–24.9 kg/m²; N = 9,783, β = 0.301, p = 1.1 × 10⁻⁴), overweight (25.0–29.9 kg/m²; N = 11,205, β = 0.328, p = 3.8 × 10⁻⁶), and obese (≥ 30.0 kg/m²; N = 13,007, β = 0.162, p = 0.015) (S2 Fig).",
        ),
        (
            "populations (Fig 4). However, no statistically significant associations",
            "populations (Fig. 1). However, no statistically significant associations",
        ),
        (
            "**Fig 3. Subgroup analyses of the association between serum GBR and depressive symptoms (BMI, sex, and age subgroups).** Forest plot showing the weighted beta coefficients and 95% confidence intervals for the association between log(GBR) and PHQ-9 total score, stratified by body mass index (BMI), sex, and age.",
            "**S2 Fig. Demographic subgroup analyses.** Adjusted beta coefficients and 95% confidence intervals across BMI, sex, and age strata.",
        ),
        (
            "**Fig 4. Subgroup analyses of the association between serum GBR and depressive symptoms (race/ethnicity subgroups).** Forest plot showing the weighted beta coefficients and 95% confidence intervals for the association between log(GBR) and PHQ-9 total score, stratified by race/ethnicity.",
            "**Fig. 1. Race/ethnicity-specific association of log(GBR) with depressive symptoms.** a, Race/ethnicity-stratified adjusted beta coefficients for PHQ-9 total score. b, Interaction contrasts relative to Non-Hispanic White participants. Error bars denote 95% confidence intervals. The global log(GBR) × race/ethnicity interaction was significant (χ² = 15.05, df = 5, p = 0.010; q = 0.024).",
        ),
        (
            "In the restricted cubic spline (RCS) analysis (Fig 5),",
            "In the restricted cubic spline (RCS) analysis (Fig. 3),",
        ),
        (
            "**Fig 5. Restricted cubic spline analysis of the association between serum GBR and depressive symptoms.**",
            "**Fig. 3. Dose-response association between log(GBR) and depressive symptoms.** Restricted cubic spline model with four knots showing the adjusted association between log(GBR) and PHQ-9 total score. Shading denotes the 95% confidence interval; vertical dashed lines denote knot positions. The Wald test for non-linearity did not reach statistical significance (χ² = 5.12, df = 2, p = 0.077).",
        ),
        (
            "**Fig 2. Specificity analyses of GBR compared with individual liver enzymes and bilirubin.** Standardized beta coefficients (per standard deviation increase) and 95% confidence intervals for log(GBR), log(GGT), log(T-Bil), ALT, and AST in relation to the PHQ-9 total score.",
            "**Fig. 2. Specificity of the GBR association relative to liver enzymes and bilirubin.** Standardized beta coefficients and 95% confidence intervals are shown for log(GBR), log(GGT), log(total bilirubin), ALT, and AST in relation to PHQ-9 total score.",
        ),
        (
            "In a joint model entering both GBR and OBS (Table 4),",
            "In a joint model entering both GBR and OBS (S4 Table),",
        ),
        (
            "**Table 4. Comparison of GBR with established oxidative stress indicators and joint models.**\n\n[Insert Table 4 here]",
            "Detailed comparisons with established oxidative stress indicators and joint models are shown in S4 Table.",
        ),
        (
            "Indeed, in the joint model (Table 4 and S4 Table),",
            "Indeed, in the joint model (S4 Table),",
        ),
        (
            "After adjusting for GGT, log(T-Bil) exhibited an independent and inverse association (S1 Table) (β = -0.087, p = 0.009).",
            "After adjusting for GGT, log(T-Bil) exhibited an independent and inverse association (Fig 2) (β = -0.087, p = 0.009).",
        ),
        (
            "Other sensitivity analyses are displayed in S1 Fig.",
            "Additional subgroup and interaction visualizations are provided in the supplementary figures.",
        ),
        (
            "indicating a stronger association between GBR and PHQ-9 in antidepressant users (S2 Fig).",
            "indicating a stronger association between GBR and PHQ-9 in antidepressant users (S4 Fig).",
        ),
        (
            "**S1 Fig. Forest plot of additional sensitivity analyses.** Forest plot showing the weighted beta coefficients and 95% confidence intervals for GBR and PHQ-9 under different sensitivity exclusions.\n**S2 Fig. Interaction between GBR and antidepressant use.** Relationship between log(GBR) and the predicted PHQ-9 total score, stratified by antidepressant use status, demonstrating the potential synergistic effect in antidepressant users.",
            "**S1 Fig. Participant selection flowchart.** Selection of eligible adults from NHANES 2007–2018.\n**S2 Fig. Demographic subgroup analyses.** Adjusted beta coefficients and 95% confidence intervals across BMI, sex, and age strata.\n**S3 Fig. Lifestyle and clinical subgroup analyses.** Adjusted beta coefficients and 95% confidence intervals across alcohol, serum cotinine, and hsCRP-related analyses.\n**S4 Fig. Antidepressant use interaction.** Predicted PHQ-9 total score by log(GBR), stratified by antidepressant use.",
        ),
    ]
    for old, new in display_replacements:
        text = text.replace(old, new)

    final_caption_cleanup = [
        (
            "By BMI categories, the association between log(GBR) and PHQ-9 was observed in all strata: normal weight (18.5–24.9 kg/m²; N = 9,783, β = 0.301, p = 1.1 × 10⁻⁴), overweight (25.0–29.9 kg/m²; N = 11,205, β = 0.328, p = 3.8 × 10⁻⁶), and obese (≥ 30.0 kg/m²; N = 13,007, β = 0.162, p = 0.015) (Fig. 3).",
            "By BMI categories, the association between log(GBR) and PHQ-9 was observed in all strata: normal weight (18.5–24.9 kg/m²; N = 9,783, β = 0.301, p = 1.1 × 10⁻⁴), overweight (25.0–29.9 kg/m²; N = 11,205, β = 0.328, p = 3.8 × 10⁻⁶), and obese (≥ 30.0 kg/m²; N = 13,007, β = 0.162, p = 0.015) (S2 Fig).",
        ),
        (
            "**Fig. 3. Dose-response association between log(GBR) and depressive symptoms.** Restricted cubic spline model with four knots showing the adjusted association between log(GBR) and PHQ-9 total score. Shading denotes the 95% confidence interval; vertical dashed lines denote knot positions. The Wald test for non-linearity did not reach statistical significance (χ² = 5.12, df = 2, p = 0.077). Restricted cubic spline curve (4 knots) demonstrating the dose-response relationship between log(GBR) and the PHQ-9 score in the fully adjusted model.",
            "**Fig. 3. Dose-response association between log(GBR) and depressive symptoms.** Restricted cubic spline model with four knots showing the adjusted association between log(GBR) and PHQ-9 total score. Shading denotes the 95% confidence interval; vertical dashed lines denote knot positions. The Wald test for non-linearity did not reach statistical significance (χ² = 5.12, df = 2, p = 0.077).",
        ),
        (
            "(Fig 2)",
            "(Fig. 2)",
        ),
    ]
    for old, new in final_caption_cleanup:
        text = text.replace(old, new)

    compression_replacements = [
        (
            "In recent years, the involvement of oxidative stress pathways in major depressive disorder (MDD) has attracted significant attention [1, 2]. An increase in systemic oxidative stress is quantified by biomarkers such as malondialdehyde (MDA) and 8-hydroxy-2'-deoxyguanosine (8-OHdG), and past studies have demonstrated elevated levels of these markers in patients with depression [3, 4]. However, these oxidative stress markers require specialized assays for measurement, making them unsuitable for large-scale epidemiological investigations [5].\n\nAgainst this background, systemic oxidative stress indices available for large-scale epidemiological studies, such as the oxidative balance score (OBS) [6, 7] and the Composite Dietary Antioxidant Index (CDAI) [8], have been proposed, and the association between OBS and depression has already been reported [9, 10]. However, the OBS requires recall data on numerous dietary and lifestyle items, raising concerns about recall bias and participant burden. Therefore, more objective and highly feasible indicators are desired.",
            "Oxidative stress pathways have attracted increasing attention in major depressive disorder (MDD) [1, 2]. Patients with depression show elevated oxidative stress markers such as malondialdehyde and 8-hydroxy-2'-deoxyguanosine [3, 4], but these assays are not routinely available in large epidemiological studies [5]. Composite indices such as the oxidative balance score (OBS) [6, 7] and Composite Dietary Antioxidant Index (CDAI) [8] have therefore been used, and OBS has been linked to depression [9, 10]. However, these indices rely on extensive dietary and lifestyle recall data, making more objective and feasible markers desirable.",
        ),
        (
            "Thus, we focused on gamma-glutamyl transferase (GGT) and total bilirubin (T-Bil), which are readily available from routine biochemical tests. GGT is an enzyme involved in glutathione metabolism and reflects cellular responses and defenses against oxidative stress [11, 12]. Conversely, bilirubin is a product of heme metabolism. Recently, T-Bil and indirect bilirubin have been shown to reflect systemic antioxidant capacity [13], and epidemiological studies have linked higher levels with a reduced risk of cardiovascular disease (CVD) [14] and type 2 diabetes [15]. Furthermore, patients with Gilbert syndrome—a benign condition characterized by constitutional jaundice caused by mutations in the bilirubin-metabolizing _UGT1A1_ gene—are known to have lower risks of CVD and cancer [16, 17]. That is, bilirubin is believed to contribute to reduced disease risk by scavenging oxidative stress [18]. In addition, interestingly, the NRF2-Keap1 pathway [19], which plays a central role in the oxidative stress response, is involved in both GGT and bilirubin regulation via glutathione metabolism and heme oxygenase metabolism, respectively [20, 21].",
            "We therefore focused on gamma-glutamyl transferase (GGT) and total bilirubin (T-Bil), both available from routine biochemical testing. GGT participates in glutathione metabolism and reflects cellular responses to oxidative stress [11, 12]. Bilirubin, a heme metabolite, functions as an endogenous antioxidant [13], and higher levels have been associated with lower risks of cardiovascular disease and type 2 diabetes [14, 15]. Gilbert syndrome, caused by variants affecting bilirubin metabolism, has also been linked to lower risks of CVD and cancer [16, 17]. In addition, the NRF2-Keap1 pathway [19] is connected to both glutathione and heme oxygenase biology, linking GGT and bilirubin to shared redox-regulatory systems [20, 21].",
        ),
        (
            "The purpose of this study was to examine whether the GGT-to-total bilirubin ratio (hereafter GBR), a putative redox-related composite proxy derived from routine biochemical measurements, is associated with depressive symptoms. We hypothesized that GBR would be associated with depressive symptom burden independently of metabolic, inflammatory, lifestyle, and medication-related factors, and that this association might differ across population subgroups in which bilirubin and liver-metabolic traits have distinct genetic and metabolic backgrounds. We also compared GBR with other oxidative stress-related markers and indices (GGT alone, T-Bil alone, OBS, and CDAI) and evaluated its specificity and heterogeneity across populations through subgroup analyses and formal interaction testing.",
            "We examined whether the GGT-to-total bilirubin ratio (GBR), a routine biochemistry-derived redox-related proxy, is associated with depressive symptoms. We hypothesized that GBR would be associated with symptom burden independently of metabolic, inflammatory, lifestyle, and medication-related factors, and that this association might differ across populations with distinct bilirubin and liver-metabolic backgrounds. We also compared GBR with GGT alone, T-Bil alone, OBS, and CDAI, and assessed heterogeneity through subgroup and interaction analyses.",
        ),
        (
            "Study design and data source\nWe used cross-sectional data from the National Health and Nutrition Examination Survey (NHANES) 2007–2018. NHANES is a nationally representative survey conducted by the Centers for Disease Control and Prevention (CDC) using a complex, multistage, probability sampling design. The present study is a secondary analysis of publicly available data. Data from the 2019–2020 cycle and later could be affected by operational changes due to the COVID-19 pandemic; hence, we limited our study to the 2007–2018 cycles.",
            "Study design and data source\nWe used cross-sectional data from NHANES 2007–2018, a nationally representative survey conducted by the Centers for Disease Control and Prevention using a complex, multistage probability design. Later cycles were not included because 2019–2020 and subsequent data may be affected by COVID-19-related operational changes.",
        ),
        (
            "Based on prior literature and biological plausibility, we adjusted for age, sex, BMI, poverty-to-income ratio (PIR), education level, marital status, total energy intake, alcohol intake, smoking history, sedentary behavior time, neutrophil-to-lymphocyte ratio (NLR), diabetes, cardiovascular disease (CVD), antidepressant use, statin use, history of liver disease, and race/ethnicity. Diabetes was defined by an HbA1c level ≥ 6.5%, fasting plasma glucose level ≥ 126 mg/dL, or self-reported physician diagnosis. CVD was defined by self-reported congestive heart failure, coronary heart disease, angina pectoris, myocardial infarction, or stroke. Antidepressant use was identified through keyword matching in the prescription medication file (RXQ_RX). Statin use, which is known to influence systemic oxidative stress balance [24], was also included. The NLR, defined as the ratio of neutrophils to lymphocytes in the white blood cell differential and widely used as a simple biomarker of systemic inflammation [25], was log-transformed [log(NLR)] and adjusted for in the primary model. Because C-reactive protein (CRP), a common inflammatory marker, was not measured in all study cycles, it was excluded from the main models and evaluated only in sensitivity analyses restricted to the 2015–2016 and 2017–2018 cycles where high-sensitivity CRP (hsCRP) was available.",
            "We adjusted for age, sex, BMI, poverty-to-income ratio, education, marital status, total energy intake, alcohol intake, smoking history, sedentary time, log-transformed neutrophil-to-lymphocyte ratio (NLR), diabetes, cardiovascular disease (CVD), antidepressant use, statin use, liver disease history, and race/ethnicity. Diabetes was defined by HbA1c ≥ 6.5%, fasting glucose ≥ 126 mg/dL, or physician diagnosis. CVD included self-reported heart failure, coronary heart disease, angina, myocardial infarction, or stroke. Antidepressant and statin use were identified from prescription medication files. Because CRP was not measured in all cycles, hsCRP was evaluated only in sensitivity analyses restricted to 2015–2018.",
        ),
        (
            "To compare GBR with established oxidative balance metrics, we evaluated the associations of GBR with CDAI [8], nutrient-only OBS, and full OBS [6, 7] in the same sample. We constructed joint models containing both GBR and OBS, and performed interaction analyses between GBR and OBS to assess whether blood-derived GBR is independently or complementarily associated with depressive symptoms relative to dietary- and lifestyle-derived scores.",
            "To compare GBR with established oxidative balance metrics, we evaluated CDAI [8], nutrient-only OBS, and full OBS [6, 7] in the same sample, and fitted joint and interaction models with GBR and OBS.",
        ),
        (
            "In specificity analyses, we examined whether GBR behaves as a redox-related indicator rather than a general marker of liver injury by comparing its standardized effect size (per standard deviation of observed values) with those of GGT alone, T-Bil alone, ALT, and AST. In addition, we constructed a joint model entering GGT and T-Bil simultaneously to investigate whether the T-Bil component provides independent information from GGT.",
            "Specificity analyses compared standardized associations for GBR, GGT alone, T-Bil alone, ALT, and AST, and used a joint GGT/T-Bil model to assess whether bilirubin contributed information independent of GGT.",
        ),
        (
            "This cross-sectional study using US NHANES data demonstrated that serum log(GBR) is independently associated with depressive symptoms (PHQ-9 scores), and that this association differs by race/ethnicity. In the primary analysis, a 1-unit increase in log(GBR) was associated with a 0.256-point higher PHQ-9 total score, and binary outcome analysis using PHQ-9 ≥ 10 showed a consistent association in the same direction. This association was generally sustained across multiple sensitivity analyses that accounted for major comorbidities, antidepressant use, smoking exposure, and clinically elevated liver enzymes. Importantly, the race/ethnicity-specific pattern was supported not only by stratified estimates but also by a formal global interaction test. Although the effect size is modest at the individual level, GBR may be useful primarily as an epidemiological and mechanistic research marker derived from routine biochemical tests, rather than as an individual-level clinical screening tool.",
            "This NHANES analysis showed that serum log(GBR) was independently associated with depressive symptoms, and that this association differed by race/ethnicity. A 1-unit increase in log(GBR) was associated with a 0.256-point higher PHQ-9 score, with a consistent logistic sensitivity result for PHQ-9 ≥ 10. The association persisted across sensitivity analyses accounting for comorbidities, antidepressant use, smoking exposure, and elevated liver enzymes. The race/ethnicity-specific pattern was supported by both stratified estimates and a formal global interaction test. Given the modest effect size, GBR is best viewed as an epidemiological and mechanistic research marker rather than an individual-level screening tool.",
        ),
        (
            "GGT is involved in extracellular glutathione metabolism and reflects protective responses against oxidative stress [11], while it is also known to participate in pro-oxidant reactions through the supply of cysteine in the presence of free iron. Because direct oxidative stress biomarkers were not available in NHANES, the interpretation of GBR as a redox-related proxy remains indirect and biologically inferential. Thus, elevated GGT has been interpreted not merely as an increase in hepatobiliary enzymes, but as an indicator reflecting systemic oxidative stress load, metabolic dysfunction, and inflammatory states [12].\n\nConversely, bilirubin is a heme metabolite that functions as an endogenous scavenger of reactive oxygen species [13]. Past epidemiological studies have linked higher bilirubin levels with lower risks of cardiovascular disease [14] and metabolic disorders [15]. Consequently, rather than evaluating GGT and T-Bil separately, GBR may provide a biologically interpretable composite framework reflecting the balance between putative pro-oxidant burden and endogenous antioxidant buffering capacity.",
            "GGT participates in extracellular glutathione metabolism and may reflect both adaptive responses to oxidative stress and pro-oxidant reactions under specific conditions [11, 12]. Bilirubin is a heme metabolite with endogenous antioxidant activity [13], and higher levels have been linked to lower cardiovascular and metabolic risk [14, 15]. Because direct oxidative stress markers were unavailable in NHANES, GBR should be interpreted as an indirect redox-related proxy, summarizing putative pro-oxidant burden and antioxidant buffering rather than directly measuring oxidative stress.",
        ),
        (
            "First, GBR demonstrated greater stability across different assumptions of missing data handling. Although MICE was used in the primary analysis, the sensitivity analysis using complete case analysis (CCA) is shown in S3 Table. For GGT alone, a certain degree of variation in effect sizes was observed between CCA and MICE (0.1853 to 0.2007). In contrast, almost no such variation was observed for GBR (0.2006 to 0.2037). This suggests that integrating multiple markers as a ratio may yield relatively stable estimates against biases and statistical noise associated with missing data.",
            "First, GBR was stable across missing-data assumptions. In CCA versus MICE comparisons (S3 Table), GGT alone varied from 0.1853 to 0.2007, whereas GBR changed minimally from 0.2006 to 0.2037.",
        ),
        (
            "Second is the balance between statistical model parsimony and goodness-of-fit. In model comparisons using Pseudo-BIC, the model containing only GBR demonstrated a superior (lower) information criterion value compared to the model containing only GGT (92,795.85 vs. 92,803.73, respectively; S2 Table). This suggests that a parsimonious model incorporating GBR may capture information related to depressive symptoms without substantially increasing model complexity.",
            "Second, GBR provided a parsimonious fit: Pseudo-BIC was lower for GBR than for GGT alone (92,795.85 vs. 92,803.73; S2 Table).",
        ),
        (
            "Fourth, GBR may provide a physiologically interpretable summary measure. Presented as a single value, GBR provides an intuitive conceptual framework to assess the balance between pro-oxidants (GGT component) and antioxidants (T-Bil component) at a glance. Importantly, the present findings should not be interpreted as demonstrating clear predictive superiority of GBR over GGT alone. Rather, the results suggest that bilirubin may contribute complementary biological information relevant to the interpretation of GGT-related signals.",
            "Fourth, GBR provides a physiologically interpretable summary of the balance between the GGT and bilirubin components. These findings do not demonstrate clear predictive superiority over GGT alone, but suggest that bilirubin adds complementary biological information to GGT-related signals.",
        ),
        (
            "The association between GBR and the PHQ-9 score was independent of established oxidative balance indices such as the OBS and CDAI. While the OBS reflects the oxidative balance derived from diet and lifestyle [6, 7], GBR may reflect the final physiological state resulting from metabolic processes, suggesting they could function complementarily.\n\nIndeed, in the joint model (S4 Table), GBR and OBS were independently associated with PHQ-9 scores, and the attenuation of their effect sizes was minimal compared to when they were entered separately, supporting their independent associations.",
            "GBR remained associated with PHQ-9 independently of OBS and CDAI. Because OBS reflects diet- and lifestyle-derived oxidative balance [6, 7], whereas GBR may reflect downstream metabolic state, these measures may be complementary. In joint models (S4 Table), attenuation was modest for both GBR and OBS.",
        ),
        (
            "In BMI-stratified analyses, the association between GBR and PHQ-9 was relatively larger in normal-weight and overweight individuals and slightly smaller in obese individuals. This finding suggests that chronic inflammation, insulin resistance, hepatic impairment, and lifestyle factors associated with obesity might affect both GBR and depressive symptoms, thereby partially diluting the GBR-specific signal.\n\nHowever, the interaction term with continuous BMI was not statistically significant (β = -0.0040, p = 0.466). Thus, effect modification by BMI is not a definitive finding in this study, and the role of obesity as an effect modifier remains speculative.",
            "BMI-stratified analyses suggested larger associations in normal-weight and overweight individuals and smaller associations in obese individuals, possibly because obesity-related inflammation, insulin resistance, hepatic impairment, and lifestyle factors dilute the GBR-specific signal. However, the continuous BMI interaction was not significant (β = -0.0040, p = 0.466), so BMI-related effect modification remains speculative.",
        ),
        (
            "These mechanistic interpretations remain hypothesis-generating because genetic polymorphisms and related metabolic pathways were not directly assessed in the present dataset. Nonetheless, the findings raise a broader point for translational biomarker research: a biomarker calculated from routine clinical chemistry may be easy to implement, but its biological meaning may still depend on population-specific genetic and metabolic context. To confirm the ethnic heterogeneity of GBR's association with depressive symptoms, external validation in Hispanic and Asian populations is therefore urgent. In particular, validation in independent cohorts that integrate _UGT1A1_-related genetic polymorphisms, insulin resistance, liver fat or NAFLD markers, and other lifestyle or environmental exposures is an important future direction to clarify the potential research utility of this index.",
            "Because genetic polymorphisms and related metabolic pathways were not directly assessed, these interpretations remain hypothesis-generating. Nonetheless, they emphasize that even routine clinical chemistry markers may have population-specific biological meanings. External validation in Hispanic and Asian populations is therefore urgent, ideally in cohorts incorporating _UGT1A1_ variants, insulin resistance, liver fat or NAFLD markers, and relevant lifestyle or environmental exposures.",
        ),
        (
            "In the primary model, the effect size per 1-unit increase in log(GBR) was a 0.2564-point increase in the PHQ-9 score. This is a modest value in terms of clinical impact at the individual level. However, from the perspective of risk stratification in epidemiological research, we consider this effect meaningful. In the primary logistic regression analysis, a 1-unit increase in log(GBR) was associated with a 16.7% increase in the odds of moderate-to-severe depressive symptoms (PHQ-9 ≥ 10; OR = 1.167). By utilizing only GGT and T-Bil, which are included in routine biochemical panels, GBR provides epidemiological studies with the potential to estimate systemic oxidative stress profiles at the population level and characterize population-level variation in depressive symptom burden without incurring any additional costs.",
            "The effect size was modest at the individual level: a 1-unit increase in log(GBR) corresponded to a 0.256-point higher PHQ-9 score and a 16.7% higher odds of PHQ-9 ≥ 10. Its relevance is therefore primarily epidemiological: GBR can be calculated from routine GGT and T-Bil measurements to characterize population-level variation in depressive symptom burden without additional assay costs.",
        ),
        (
            "This study has several limitations. First, because this is a cross-sectional study, the direction of causality cannot be established. Whether oxidative stress causes depressive symptoms, depressive symptoms exacerbate oxidative stress states, or both, cannot be distinguished from these results. To determine the direction of causality, longitudinal designs are required. Second, as mentioned earlier, MICE assumes a missing at random (MAR) mechanism, which cannot rule out bias due to a missing not at random (MNAR) mechanism, where depressive symptoms or general health conditions directly affect non-response. Sensitivity analyses such as delta-adjustment for MNAR [34, 29] remain a task for future research. Furthermore, although modeling PHQ-9 scores as a continuous outcome is common in large epidemiological studies, the bounded and potentially skewed distribution of PHQ-9 scores should be considered when interpreting effect estimates. Third, there is the issue of unmeasured confounding in the race/ethnicity-specific analysis. Data on the aforementioned _UGT1A1_ polymorphisms (_UGT1A1_*28 and _UGT1A1_*6) and _PNPLA3_ polymorphisms were not available in the public NHANES repositories; thus, direct adjustment for these genetic factors could not be performed. Other factors, such as environmental exposures and cultural differences, were also unmeasured. Direct adjustment for genetic variants is essential to elucidate population differences. Fourth, single-time-point measurements of GGT and bilirubin may not fully reflect chronic dysfunction of the systemic oxidative stress balance, which is a common limitation of cross-sectional designs. Investigating within-person changes in GBR via longitudinal, repeated measurements is necessary. Fifth, NHANES laboratory data lack comprehensive direct bilirubin values, forcing us to adopt total bilirubin as the denominator of the exposure ratio. Clinically, indirect bilirubin—which excludes direct bilirubin directly affected by hepatobiliary conditions—might be more suitable as a proxy for antioxidant capacity. However, while the lack of available bilirubin fractions in the NHANES database partly explains our choice, utilizing total bilirubin—which is included in routine chemical screens and thus highly translatable to epidemiological studies—and observing the consistent associations described above carries practical value. Sixth, although the adjustment for covariates in the primary analysis was comprehensive, residual confounding by unmeasured factors, such as other genetic variants, micronutrient status, and gut microbiota, cannot be completely ruled out. Notably, the E-value based on the primary logistic analysis was 1.61, suggesting that unmeasured confounding of moderate magnitude could potentially explain away the observed association. Seventh, hsCRP measurements in NHANES were restricted to limited cycles. Among the cycles analyzed, hsCRP was measured in the I (2015–2016) and J (2017–2018) cycles, whereas the E (2007–2008) and F (2009–2010) cycles featured conventional CRP measurements with different detection limits. Furthermore, no CRP data were collected in the G (2011–2012) and H (2013–2014) cycles. Due to these data constraints, we adjusted for NLR in the main model and conducted sensitivity analyses adjusting for hsCRP restricted to the I and J cycles. While adjusting for NLR or hsCRP yielded similar associations, the inability to adjust for CRP across all cycles remains a limitation of the data used in this study.",
            "This study has several limitations. First, its cross-sectional design precludes causal inference or determination of temporal direction. Longitudinal studies are needed to clarify whether redox imbalance precedes depressive symptoms, follows them, or both. Second, MICE assumes missing at random and cannot exclude missing-not-at-random mechanisms related to depression or health status; MNAR sensitivity analyses remain a future task [34, 29]. The bounded and potentially skewed distribution of PHQ-9 scores should also be considered when interpreting linear-model estimates. Third, genetic variants such as _UGT1A1_*28, _UGT1A1_*6, and _PNPLA3_ were unavailable in public NHANES files, and detailed environmental and cultural factors were unmeasured, limiting interpretation of race/ethnicity-specific findings. Fourth, single-time-point GGT and bilirubin measurements may not capture chronic redox balance. Fifth, bilirubin fractions were unavailable, so total bilirubin was used despite indirect bilirubin being a potentially more specific antioxidant proxy. Sixth, residual confounding by genetics, micronutrient status, gut microbiota, or other unmeasured factors remains possible; the logistic E-value was 1.61. Seventh, hsCRP was available only in 2015–2018, although sensitivity analyses adjusting for NLR and hsCRP yielded similar results.",
        ),
        (
            "In this cross-sectional study using a nationally representative sample from the US NHANES, we demonstrated that the serum GGT-to-total bilirubin ratio (GBR) is independently associated with depressive symptoms, and that this association differs by race/ethnicity. Calculated from routine biochemistry measures, GBR is a highly feasible candidate composite marker for epidemiological research related to systemic redox biology. However, caution should be exercised regarding the cross-sectional design, the relatively small effect size, and the limited superiority of GBR over GGT alone in predictive performance. The present findings should primarily be interpreted as hypothesis-generating rather than clinically predictive. Important future research directions include: (1) external validation in independent cohorts including Asian and Hispanic populations; (2) stratified analyses incorporating genetic polymorphisms related to bilirubin and glutathione metabolic pathways; and (3) longitudinal evaluation of GBR's temporal stability and predictive performance for depression. More broadly, these results highlight that the translation of routine redox-related biomarkers into psychiatric epidemiology may require explicit attention to population-specific biological context.",
            "In this nationally representative NHANES study, GBR was independently associated with depressive symptoms, and the association differed by race/ethnicity. GBR is a feasible composite marker for epidemiological research on systemic redox biology, but the cross-sectional design, modest effect size, and limited predictive advantage over GGT alone warrant caution. The findings should be interpreted as hypothesis-generating rather than clinically predictive. Future work should validate GBR in independent Asian and Hispanic cohorts, incorporate genetic variation in bilirubin and glutathione pathways, and evaluate temporal stability and longitudinal prediction. More broadly, these results suggest that translating routine redox-related biomarkers into psychiatric epidemiology requires attention to population-specific biological context.",
        ),
    ]
    for old, new in compression_replacements:
        text = text.replace(old, new)

    reference_boost_replacements = [
        (
            "In sensitivity analyses, a PHQ-9 score ≥ 10 was used as a binary outcome representing moderate-to-severe depressive symptoms.",
            "In sensitivity analyses, a PHQ-9 score ≥ 10 was used as a binary outcome representing moderate-to-severe depressive symptoms, consistent with the original PHQ-9 validation study [24].",
        ),
        (
            "We used cross-sectional data from NHANES 2007–2018, a nationally representative survey conducted by the Centers for Disease Control and Prevention using a complex, multistage probability design.",
            "We used cross-sectional data from NHANES 2007–2018, a nationally representative survey conducted by the Centers for Disease Control and Prevention using a complex, multistage probability design [23].",
        ),
        (
            "Certain antidepressants undergo hepatic metabolism and can be associated with elevated liver enzymes or drug-induced liver injury.",
            "Certain antidepressants undergo hepatic metabolism and can be associated with elevated liver enzymes or drug-induced liver injury [31].",
        ),
        (
            "The _UGT1A1_*28 allele, which contributes to the Gilbert syndrome phenotype characterized by elevated serum bilirubin, is more prevalent in European White and Black populations but less common in East Asian and some Hispanic populations [14]. In Asian populations, the _UGT1A1_*6 polymorphism more commonly contributes to the Gilbert phenotype [28].",
            "The _UGT1A1_*28 allele, which contributes to the Gilbert syndrome phenotype characterized by elevated serum bilirubin, is more prevalent in European White and Black populations but less common in East Asian and some Hispanic populations [14, 27]. In Asian populations, the _UGT1A1_*6 polymorphism more commonly contributes to the Gilbert phenotype [28, 27].",
        ),
        (
            "Hispanic populations have a high frequency of the _PNPLA3_ I148M variant associated with lipid droplet accumulation [29], which may increase susceptibility to non-alcoholic fatty liver disease and influence GGT elevation and bilirubin dynamics.",
            "Hispanic populations have a high frequency of the _PNPLA3_ I148M variant associated with lipid droplet accumulation [29, 30], which may increase susceptibility to non-alcoholic fatty liver disease and influence GGT elevation and bilirubin dynamics.",
        ),
    ]
    for old, new in reference_boost_replacements:
        text = text.replace(old, new)

    text = text.replace(
        "[repository URL]",
        "https://github.com/mattakuya/NHANES_GBR",
    )

    # Normalize figure/table citation labels to Nature Portfolio / Translational Psychiatry style
    text = re.sub(r'\bS(\d+)\s+Fig\b', r'Supplementary Fig. S\1', text)
    text = re.sub(r'\bS(\d+)\s+Table\b', r'Supplementary Table S\1', text)
    text = re.sub(r'\bS(\d+)\s+Checklist\b', r'Supplementary Checklist S\1', text)
    text = re.sub(r'\bFig\s+(\d+)\b', r'Fig. \1', text)

    text = renumber_references_by_citation_order(text)
    
    # Replace thousands separator commas with spaces to conform to TP House Style
    text = replace_thousands_commas_with_spaces(text)

    # Save Markdown Main Manuscript
    OUTPUT.write_text(text, encoding="utf-8")
    print(f"Successfully compiled and saved English manuscript to: {OUTPUT}")

    # Generate Supplementary Material Markdown
    supp_text = f"""# Race- and ethnicity-specific association of the gamma-glutamyl transferase-to-bilirubin ratio with depressive symptoms in US adults: a NHANES 2007–2018 study
## Supplementary Materials

**Takuya Matsuoka¹***
¹ Tohoku University Hospital, Sendai, Japan
* Corresponding author (email: takuya.matsuoka.c6@tohoku.ac.jp)

---

### Supplementary Tables

#### S1 Table. Multicollinearity diagnostics for main model covariates
{supp_table1}

#### S2 Table. Model fit comparison of the primary exposure variables
{supp_table2}

#### S3 Table. Complete case analysis versus multiple imputation by chained equations (MICE)
{supp_table3}

#### S4 Table. Associations with established oxidative stress indicators and joint models
{supp_table4}

#### S5 Table. Antidepressant interaction analyses
{supp_table5}

#### S6 Table. Race/ethnicity interaction analysis
{supp_table6}

#### S7 Table. Missingness of primary analysis variables before multiple imputation
{supp_table7}

---
"""
    # Strip markdown-like asterisks from supplementary materials
    supp_text = strip_markdown_italics(supp_text)
    
    # Replace thousands separator commas with spaces in supplementary materials
    supp_text = replace_thousands_commas_with_spaces(supp_text)

    # Normalize supplementary material headings to Nature style
    supp_text = re.sub(r'\bS(\d+)\s+Fig\b', r'Supplementary Fig. S\1', supp_text)
    supp_text = re.sub(r'\bS(\d+)\s+Table\b', r'Supplementary Table S\1', supp_text)
    supp_text = re.sub(r'\bS(\d+)\s+Checklist\b', r'Supplementary Checklist S\1', supp_text)

    supp_output_path = ROOT / "results" / "supplementary_material_tp.md"
    supp_output_path.write_text(supp_text, encoding="utf-8")
    print(f"Successfully compiled and saved English supplementary material to: {supp_output_path}")

    # Convert to Word documents
    convert_md_to_docx(str(OUTPUT), str(ROOT / "results" / "manuscript_tp.docx"))
    convert_md_to_docx(str(supp_output_path), str(ROOT / "results" / "supplementary_material_tp.docx"))
    print("All English Translational Psychiatry submissions compiled successfully!")

if __name__ == "__main__":
    main()
