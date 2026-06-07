import os
import glob
import re
import shutil
import math
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def get_latest_results_file(results_dir="results"):
    pattern = os.path.join(results_dir, "nhanes_results_*.md")
    files = glob.glob(pattern)
    # Exclude QUICK files
    files = [f for f in files if "QUICK" not in f]
    if not files:
        raise FileNotFoundError("Results file matching 'nhanes_results_*.md' not found in results directory.")
    # Return the latest file by modification time
    return max(files, key=os.path.getmtime)

def parse_markdown_table(file_path, header_title):
    """
    Parses a markdown table from the results file following a specific heading title.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Locate the heading
    pattern = rf"##\s+{re.escape(header_title)}\s*\n\n_FDR family:.*_\n\n(\|[\s\S]*?)(?=\n\n##|\Z)"
    match = re.search(pattern, content)
    if not match:
        # Try without the FDR family note
        pattern_simple = rf"##\s+{re.escape(header_title)}\s*\n\n(\|[\s\S]*?)(?=\n\n##|\Z)"
        match = re.search(pattern_simple, content)
    
    if match:
        return match.group(1).strip()
    else:
        raise ValueError(f"Could not find table under heading: {header_title}")

def build_table_2(file_path):
    """
    Constructs Table 2 (Primary and Sensitivity Analyses in the Full Sample)
    """
    main_table = parse_markdown_table(file_path, "MAIN ANALYSIS: PHQ-9 CONTINUOUS")
    logistic_table = parse_markdown_table(file_path, "SENSITIVITY: PHQ-9 ≥ 10 (Logistic, β on log-odds)")
    sensitivity_table = parse_markdown_table(file_path, "SENSITIVITY ANALYSIS")
    liver_safe_table = parse_markdown_table(file_path, "LIVER-SAFE (NO DILI) ANALYSIS")
    
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
    or_val = math.exp(float(beta))
    or_lo = math.exp(float(ci_lo))
    or_hi = math.exp(float(ci_hi))
    rows.append(f"| Full Sample (Logistic PHQ-9 &ge; 10) | {n} | OR = {or_val:.4f} [{or_lo:.4f}, {or_hi:.4f}] | — | {p} | {q} |")
    
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

def build_table_3_subgroups(file_path):
    """
    Constructs Table 3 (Comprehensive Subgroup Analyses) -> Supplementary Table 1
    """
    bmi_table = parse_markdown_table(file_path, "BMI STRATIFIED ANALYSIS")
    sex_table = parse_markdown_table(file_path, "SEX-STRATIFIED ANALYSIS")
    age_table = parse_markdown_table(file_path, "AGE-STRATIFIED ANALYSIS")
    race_table = parse_markdown_table(file_path, "RACE-STRATIFIED ANALYSIS")
    alcohol_table = parse_markdown_table(file_path, "ALCOHOL-STRATIFIED ANALYSIS (NIAAA)")
    cotinine_table = parse_markdown_table(file_path, "COTININE-STRATIFIED ANALYSIS (Pirkle)")
    crp_table = parse_markdown_table(file_path, "CRP-STRATIFIED ANALYSIS (AHA, I/J only)")
    
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

def build_table_3(file_path):
    """
    Constructs Table 4 (Specificity Analyses standardized per 1 SD) -> Supplementary Table 2
    """
    spec_table = parse_markdown_table(file_path, "SPECIFICITY ANALYSIS (SD-STANDARDIZED)")
    joint_table = parse_markdown_table(file_path, "BILIRUBIN BEYOND GGT (JOINT MODEL, per 1 SD)")
    
    def extract_rows(table_str):
        return table_str.strip().split('\n')[2:]
        
    header = "| Model / Term | N | Beta [95% CI] | SE | P-value | Q-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    
    # Univariate
    rows.append("| **Univariate Standardized Models** | | | | | |")
    for row_str in extract_rows(spec_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    # Joint
    rows.append("| **Joint Model of GGT + Bilirubin** | | | | | |")
    for row_str in extract_rows(joint_table):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        term, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {term} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} | {q} |")
        
    return '\n'.join([header, separator] + rows)

def build_table_4(file_path):
    """
    Constructs Table 5 (Independent and Additive Association GBR and OBS per 1 SD) -> Main Table 3
    """
    obs_table = parse_markdown_table(file_path, "JOINT MODEL: log_ratio + OBS (per 1 SD each, SD-standardized)")
    
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

def build_table_hscrp_sensitivity(file_path):
    table_str = parse_markdown_table(file_path, "I/J HS-CRP SENSITIVITY")
    def extract_rows(t_str):
        return t_str.strip().split('\n')[2:]
    
    header = "| Subgroup / Model | N | Beta [95% CI] | SE | P-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: |"
    
    rows = []
    for row_str in extract_rows(table_str):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        subgroup, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {subgroup} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} |")
        
    return '\n'.join([header, separator] + rows)

def build_table_cca_vs_mice(file_path):
    table_str = parse_markdown_table(file_path, "CCA vs MICE: GGT BIAS ASSESSMENT (SD-standardized)")
    def extract_rows(t_str):
        return t_str.strip().split('\n')[2:]
    
    header = "| Marker | Method | N | Beta [95% CI] | SE | P-value |"
    separator = "| :--- | :---: | :---: | :---: | :---: | :---: |"
    
    rows = []
    for row_str in extract_rows(table_str):
        parts = [x.strip() for x in row_str.split('|')[1:-1]]
        marker, method, n, beta, se, ci_lo, ci_hi, p, q, q_fam = parts
        rows.append(f"| {marker} | {method} | {n} | {beta} [{ci_lo}, {ci_hi}] | {se} | {p} |")
        
    return '\n'.join([header, separator] + rows)

def get_manuscript_text(table1, table2, table3, table4):
    text = r"""Race- and ethnicity-specific association of the GGT-to-bilirubin ratio with depressive symptoms in US adults: an NHANES 2007–2018 study

Takuya Matsuoka¹*

¹ Tohoku University Hospital, Sendai, Japan

* Corresponding author:
E-mail: takuya.matsuoka.c6@tohoku.ac.jp

Abstract
Oxidative stress is implicated in the pathophysiology of depression, but scalable biomarkers that capture redox-related heterogeneity in population settings remain limited. Gamma-glutamyl transferase (GGT) and bilirubin have been proposed as markers related to pro-oxidant burden and endogenous antioxidant capacity, respectively. We examined whether the GGT-to-total bilirubin ratio (GBR), a routine biochemistry-derived redox-related composite, is associated with depressive symptoms and whether this association differs by race/ethnicity. We analyzed cross-sectional data from the National Health and Nutrition Examination Survey (NHANES) 2007–2018 for adults aged ≥18 years (N = 36,580). Missing values were imputed using multiple imputation by chained equations (m = 20), and pooled weighted least squares regression was performed. The primary outcome was the Patient Health Questionnaire-9 (PHQ-9) score, and the exposure was mean-centered log(GBR). In the fully adjusted model, log(GBR) was associated with higher PHQ-9 scores (β = 0.256, 95% confidence interval [CI]: 0.171–0.341, p = 4.8 × 10⁻⁸). Logistic sensitivity analysis with PHQ-9 ≥ 10 as the outcome supported this association (odds ratio [OR] = 1.167, 95% CI: 1.078–1.264). The association persisted after excluding antidepressant users, those with cardiovascular disease or diabetes, and participants with elevated ALT/AST. Race/ethnicity-stratified analyses showed positive associations in Non-Hispanic White and Non-Hispanic Black participants, but not in Hispanic or Non-Hispanic Asian participants. A formal global interaction test supported heterogeneity by race/ethnicity (χ² = 15.05, df = 5, p = 0.010). These findings suggest that GBR is associated with depressive symptom burden, but its validity as a redox-related epidemiological marker may not be population-invariant. External validation in Asian and Hispanic populations, ideally with genetic and metabolic characterization, is needed.

Introduction
Oxidative stress pathways have attracted increasing attention in major depressive disorder (MDD) [1, 2]. Patients with depression show elevated oxidative stress markers such as malondialdehyde and 8-hydroxy-2'-deoxyguanosine [3, 4], but these assays are not routinely available in large epidemiological studies [5]. Composite indices such as the oxidative balance score (OBS) [6, 7] and Composite Dietary Antioxidant Index (CDAI) [8] have therefore been used, and OBS has been linked to depression [9, 10]. However, these indices rely on extensive dietary and lifestyle recall data, making more objective and feasible markers desirable.

We therefore focused on gamma-glutamyl transferase (GGT) and total bilirubin (T-Bil), both available from routine biochemical testing. GGT participates in glutathione metabolism and reflects cellular responses to oxidative stress [11, 12]. Bilirubin, a heme metabolite, functions as an endogenous antioxidant [13, 34], and higher levels have been associated with lower risks of cardiovascular disease and type 2 diabetes [14, 15]. Gilbert syndrome, caused by variants affecting bilirubin metabolism, has also been linked to lower risks of CVD and cancer [16, 17]. In addition, the NRF2-Keap1 pathway [18] is connected to both glutathione and heme oxygenase biology, linking GGT and bilirubin to shared redox-regulatory systems [19, 20].

In the past, the associations of GGT alone and T-Bil alone with depressive symptoms have been reported individually [21, 22]. However, previous reports treated GGT or bilirubin separately as markers of liver dysfunction, inflammation, or metabolic abnormalities. To our knowledge, investigations integrating the two into a single ratio as a putative redox-related composite remain limited.

We examined whether the GGT-to-total bilirubin ratio (GBR), a routine biochemistry-derived redox-related proxy, is associated with depressive symptoms. We hypothesized that GBR would be associated with symptom burden independently of metabolic, inflammatory, lifestyle, and medication-related factors, and that this association might differ across populations with distinct bilirubin and liver-metabolic backgrounds. We also compared GBR with GGT alone, T-Bil alone, OBS, and CDAI, and assessed heterogeneity through subgroup and interaction analyses.

Materials and methods
Study design and data source
We used cross-sectional data from NHANES 2007–2018, a nationally representative survey conducted by the Centers for Disease Control and Prevention using a complex, multistage probability design [23]. Later cycles were not included because 2019–2020 and subsequent data may be affected by COVID-19-related operational changes.

Ethical considerations
The NHANES study protocols were approved by the National Center for Health Statistics (NCHS) Research Ethics Review Board. All participants provided written informed consent. The present study is a secondary analysis of publicly available, de-identified data and was therefore deemed exempt from institutional review board oversight.

Study participants
Adults aged 18 years and older were extracted from the dataset. Those with missing Mobile Examination Center (MEC) exam weights were excluded from the analysis. The final analysis cohort consisted of N = 36,580 participants (S1 Fig).

**S1 Fig. Participant selection flowchart.** Selection of eligible adults from NHANES 2007–2018.

Exposure
The primary exposure variable was the natural logarithm of the GGT-to-total bilirubin ratio, log(GBR), calculated from serum GGT [U/L] and T-Bil [mg/dL]. Values of GGT or T-Bil ≤ 0 were treated as missing. In the main analyses, log(GBR) was treated as a continuous variable and mean-centered before being entered into the models.

Outcome
Depressive symptoms were assessed using the continuous total score (range: 0–27) of the Patient Health Questionnaire-9 (PHQ-9). Each item (DPQ010–DPQ090) retained only valid responses (0–3), and the PHQ-9 total score was calculated by summing the nine items after MICE imputation. In sensitivity analyses, a PHQ-9 score ≥ 10 was used as a binary outcome representing moderate-to-severe depressive symptoms, consistent with the original PHQ-9 validation study [24].

Covariates
We adjusted for age, sex, BMI, poverty-to-income ratio, education, marital status, total energy intake, alcohol intake, smoking history, sedentary time, log-transformed neutrophil-to-lymphocyte ratio (NLR), diabetes, cardiovascular disease (CVD), antidepressant use, statin use, liver disease history, and race/ethnicity. Diabetes was defined by HbA1c ≥ 6.5%, fasting glucose ≥ 126 mg/dL, or physician diagnosis. CVD included self-reported heart failure, coronary heart disease, angina, myocardial infarction, or stroke. Antidepressant and statin use were identified from prescription medication files. Statin use, which is known to influence systemic oxidative stress balance [35], was also included. The NLR, defined as the ratio of neutrophils to lymphocytes in the white blood cell differential and widely used as a simple biomarker of systemic inflammation [36], was log-transformed and adjusted for in the primary model. Because CRP was not measured in all cycles, hsCRP was evaluated only in sensitivity analyses restricted to 2015–2018.

Statistical analysis
Missing values were handled using multiple imputation by chained equations (MICE) [37, 38]. Multiple imputation was performed using the chained equations approach via the IterativeImputer framework in scikit-learn, with posterior sampling enabled and a maximum of 30 iterations. The Mobile Examination Center (MEC) exam weight (WTMEC2YR) was included as an auxiliary variable to preserve the survey structure. We generated m = 20 imputed datasets. The same analytical models were applied to each dataset, and the estimates were pooled using Rubin's rules [25, 37] with the Barnard-Rubin degrees-of-freedom correction [26].

In the primary analysis, pooled weighted least squares (WLS) regression was used to estimate the association between log(GBR) and the PHQ-9 total score. To integrate multiple cycles, the MEC weight (WTMEC2YR) was divided by the number of cycles (6). Cluster-robust standard errors (SE) based on primary sampling units (PSU: SDMVPSU) were used for variance estimation.

To evaluate multicollinearity, the variance inflation factor (VIF) among covariates in the main model was calculated, with a VIF > 10 considered indicative of high multicollinearity. For exploratory tests spanning multiple analysis families, the Benjamini-Hochberg method was applied for False Discovery Rate (FDR) correction within each family.

As sensitivity analyses, we performed logistic regression with PHQ-9 ≥ 10 as the outcome, compared results with complete case analysis (CCA), excluded antidepressant users, excluded participants with a history of CVD, excluded those with diabetes, excluded all three groups simultaneously, restricted the analysis to individuals with normal ALT and AST levels, and performed analyses stratified by serum cotinine levels. Furthermore, in the subset of participants from the I/J cycles (2015–2018) where hsCRP was measured, we evaluated the impact of adjusting for NLR and hsCRP.

To compare GBR with established oxidative balance metrics, we evaluated CDAI [8], nutrient-only OBS, and full OBS [6, 7] in the same sample, and fitted joint and interaction models with GBR and OBS.

Specificity analyses compared standardized associations for GBR, GGT alone, T-Bil alone, ALT, and AST, and used a joint GGT/T-Bil model to assess whether bilirubin contributed information independent of GGT.

In interaction analyses, interactions between log(GBR) and BMI, antidepressant use, and smoking indicators were tested in separate models.

In dose-response analyses, to evaluate the linearity of the association between log(GBR) and the PHQ-9 score, restricted cubic spline (RCS) regression with four knots was implemented within the MICE integration framework, and non-linearity was assessed using a Wald test for non-linear terms.

Subgroup analyses were conducted stratified by BMI categories (18.5 ≤ BMI < 25.0, 25.0 ≤ BMI < 30.0, BMI ≥ 30.0 kg/m²), sex, and race/ethnicity (Non-Hispanic White, Non-Hispanic Black, Hispanic [combined, Mexican American, and Other Hispanic], and Non-Hispanic Asian). To reflect the stratum-specific missingness patterns, MICE was performed separately within each stratum for subgroup analyses. To formally evaluate population heterogeneity, we additionally fitted an interaction model including log(GBR) × race/ethnicity terms, using Non-Hispanic White participants as the reference group, and tested all non-reference interaction terms jointly with a pooled Wald chi-square test across imputed datasets.

All statistical analyses were performed using Python (version 3.11) with pandas, scikit-learn, and statsmodels. The significance level was set at two-tailed α = 0.05.

Results
Sample characteristics
The final analysis cohort included N = 36,580 adults aged 18 years and older (Table 1). The unweighted number of participants in each log(GBR) quartile was 8,201. With increasing log(GBR) quartiles, BMI, GGT, smoking history, antidepressant use, diabetes, CVD, and history of liver disease generally increased, while T-Bil decreased monotonically. In contrast, log(NLR) showed no substantial differences across quartiles, suggesting that GBR contains information distinct from simple systemic inflammatory markers.

**Table 1. Participant characteristics by quartiles of log(GBR).** Values are unweighted counts, means (SD), or weighted percentages, as indicated.

{table1}

Primary analysis
In the fully adjusted model (N = 36,580), log(GBR) was positively associated with higher PHQ-9 total scores (β = 0.256, SE = 0.043, 95% CI: 0.171–0.341, p = 4.8 × 10⁻⁸; q = 1.0 × 10⁻⁶). A logistic sensitivity analysis using PHQ-9 ≥ 10 as the outcome demonstrated a consistent association (Table 2) (OR = 1.167, 95% CI: 1.078–1.264, E-value = 1.61).

**Table 2. Primary and sensitivity analyses of the association between log(GBR) and depressive symptoms.** Estimates are from fully adjusted weighted models unless otherwise indicated.

{table2}

The VIF among all covariates was less than 2 (with the maximum being 1.74 for age), indicating no multicollinearity (S1 Table).

Specificity analysis
In specificity analyses, both log(GBR) (β = 0.204, p = 4.7 × 10⁻⁸) and log(GGT) (β = 0.201, p = 6.5 × 10⁻⁸) showed positive associations of similar magnitude with PHQ-9 scores. In contrast, log(T-Bil) was inversely associated (β = -0.073, p = 0.029), whereas ALT and AST showed no significant associations. In the joint model including both log(GGT) and log(T-Bil), log(T-Bil) remained independently and inversely associated with the outcome (β = -0.087, p = 0.009) (Fig. 2).

**Fig. 2. Specificity of the GBR association relative to liver enzymes and bilirubin.** Standardized beta coefficients and 95% confidence intervals are shown for log(GBR), log(GGT), log(total bilirubin), ALT, and AST in relation to PHQ-9 total score.

Regarding model fit, the AIC was lowest (best) for the GGT+Bil joint model, and the ratio-only model was slightly better than the GGT-only model. Conversely, the BIC was lowest (best) for the GBR model, followed by the GGT+T-Bil joint and GGT-only models (S2 Table).

Sensitivity analyses
In key sensitivity analyses, the positive association remained statistically significant after excluding antidepressant users (N = 32,801, β = 0.216, p = 1.7 × 10⁻⁶), individuals with a history of CVD (N = 30,875, β = 0.241, p = 1.2 × 10⁻⁶), patients with diabetes (N = 31,911, β = 0.258, p = 7.7 × 10⁻⁸), or all three groups simultaneously (N = 25,132, β = 0.218, p = 1.2 × 10⁻⁵). The association also remained significant in the cohort with normal ALT/AST levels (N = 28,990, β = 0.241, p = 1.8 × 10⁻⁵). Estimates from complete case analysis (CCA) and MICE were highly comparable (CCA β = 0.201 vs. MICE β = 0.204 for GBR; S3 Table).

In the hsCRP sensitivity analysis restricted to the 2015–2016 and 2017–2018 cycles (N = 11,848), a significant positive association was observed across all adjustments: without inflammatory covariates (β = 0.179, p = 0.019), with NLR adjustment (β = 0.181, p = 0.018), with hsCRP adjustment (β = 0.175, p = 0.021), and with concurrent NLR and hsCRP adjustment (β = 0.178, p = 0.020).

In analyses stratified by serum cotinine levels, positive associations were found in all strata: non-exposed (< 3 ng/mL: β = 0.186, p = 0.00013), transitional (3–30 ng/mL: β = 0.457, p = 0.024), and active smoking equivalent (> 30 ng/mL: β = 0.314, p = 0.002).

Additional subgroup and interaction visualizations are provided in the supplementary figures.

Subgroup analyses
Additional subgroup visualizations are provided in S2 and S3 Figs and summarized below.

By BMI categories, the association between log(GBR) and PHQ-9 was observed in all strata: normal weight (18.5–24.9 kg/m²; N = 9,783, β = 0.301, p = 1.1 × 10⁻⁴), overweight (25.0–29.9 kg/m²; N = 11,205, β = 0.328, p = 3.8 × 10⁻⁶), and obese (≥ 30.0 kg/m²; N = 13,007, β = 0.162, p = 0.015) (S2 Fig). Although the effect size was relatively smaller in the obese group, the interaction with continuous BMI was not significant (β = -0.0040, p = 0.466).

By sex, positive associations of similar magnitude were observed in both males (N = 17,783, β = 0.258, p = 3.1 × 10⁻⁶) and females (N = 18,797, β = 0.267, p = 1.4 × 10⁻⁵).

By age, positive associations were found in all strata: 18–39 years (β = 0.207, p = 7.4 × 10⁻⁴), 40–64 years (β = 0.267, p = 1.3 × 10⁻⁴), and 65 years and older (β = 0.183, p = 0.023).

By race/ethnicity, positive associations were identified in Non-Hispanic White (N = 14,561, β = 0.352, p = 2.4 × 10⁻⁷) and Non-Hispanic Black (N = 7,912, β = 0.278, p = 8.8 × 10⁻⁴) populations (Fig. 1). However, no statistically significant associations were observed in the Hispanic combined cohort (N = 9,487, β = -0.024, p = 0.749), Mexican American, Other Hispanic, or Non-Hispanic Asian populations. This racial/ethnic pattern was consistently maintained even after excluding antidepressant users. In the formal interaction model using Non-Hispanic White participants as the reference group, the global log(GBR) × race/ethnicity interaction was statistically significant (χ² = 15.05, df = 5, p = 0.010; q = 0.024), supporting heterogeneity across racial/ethnic groups. The strongest negative interaction contrasts relative to Non-Hispanic White participants were observed for Mexican American participants (βinteraction = -0.293, p = 0.0065; q = 0.023) and Non-Hispanic Asian participants (βinteraction = -0.207, p = 0.030; q = 0.053).

**S2 Fig. Demographic subgroup analyses.** Adjusted beta coefficients and 95% confidence intervals across BMI, sex, and age strata.

**Fig. 1. Race/ethnicity-specific association of log(GBR) with depressive symptoms.** a, Race/ethnicity-stratified adjusted beta coefficients for PHQ-9 total score. b, Interaction contrasts relative to Non-Hispanic White participants. Error bars denote 95% confidence intervals. The global log(GBR) × race/ethnicity interaction was significant (χ² = 15.05, df = 5, p = 0.010; q = 0.024).

Interaction analyses
In interaction analyses, log(GBR) was positively associated with PHQ-9 scores even among antidepressant non-users (β = 0.189, p = 1.4 × 10⁻⁵). Furthermore, a significant interaction was found between log(GBR) and antidepressant use (β = 0.480, p = 6.9 × 10⁻⁴), indicating a stronger association between GBR and PHQ-9 in antidepressant users (S4 Fig). In contrast, this interaction was not clear in the normal BMI group (S5 Table) (β = 0.163, p = 0.530).

Dose-response analyses
In the restricted cubic spline (RCS) analysis (Fig. 3), the Wald test for the non-linearity of the association between log(GBR) and PHQ-9 was borderline (χ² = 5.12, df = 2, p = 0.077) and did not reach statistical significance. Therefore, although the primary analysis model did not provide clear evidence of non-linearity, the possibility of some non-linear relationship cannot be entirely ruled out.

**Fig. 3. Dose-response association between log(GBR) and depressive symptoms.** Restricted cubic spline model with four knots showing the adjusted association between log(GBR) and PHQ-9 total score. Shading denotes the 95% confidence interval; vertical dashed lines denote knot positions. The Wald test for non-linearity did not reach statistical significance (χ² = 5.12, df = 2, p = 0.077).

Comparison with other established oxidative stress indices
In the subset where the OBS could be calculated (N = 30,335), comparing standardized effect sizes (per 1 SD) showed that GBR (β = 0.216), GGT alone (β = 0.213), CDAI (β = -0.115), nutrient-only OBS (β = -0.288), and full OBS (β = -0.435) were all associated with PHQ-9 scores in the expected directions.

In a joint model entering both GBR and OBS (S4 Table), both GBR (β = 0.261, p = 2.1 × 10⁻¹¹) and OBS (β = -0.382, p = 2.5 × 10⁻¹⁵) remained independently associated with the outcome, with only minimal attenuation of their effects compared to univariate models (15.3% attenuation for GBR and 12.1% for OBS). The positive association of log(GBR) was maintained across all OBS tertiles, and the log(GBR) × OBS interaction term was not significant (S4 Table) (β = -0.023, p = 0.477).

Detailed comparisons with established oxidative stress indicators and joint models are shown in S4 Table.

Discussion
This NHANES analysis showed that serum log(GBR) was independently associated with depressive symptoms, and that this association differed by race/ethnicity. A 1-unit increase in log(GBR) was associated with a 0.256-point higher PHQ-9 score, with a consistent logistic sensitivity result for PHQ-9 ≥ 10. The association persisted across sensitivity analyses accounting for comorbidities, antidepressant use, smoking exposure, and elevated liver enzymes. The race/ethnicity-specific pattern was supported by both stratified estimates and a formal global interaction test. Given the modest effect size, GBR is best viewed as an epidemiological and mechanistic research marker rather than an individual-level screening tool.

Biological interpretation of GBR
GGT participates in extracellular glutathione metabolism and may reflect both adaptive responses to oxidative stress and pro-oxidant reactions under specific conditions [11, 12]. Bilirubin is a heme metabolite with endogenous antioxidant activity [13, 34], and higher levels have been linked to lower cardiovascular and metabolic risk [14, 15]. Because direct oxidative stress markers were unavailable in NHANES, GBR should be interpreted as an indirect redox-related proxy, summarizing putative pro-oxidant burden and antioxidant buffering rather than directly measuring oxidative stress.

GBR relative to GGT alone
In this study, the effect size of GBR in the primary model was highly comparable to that of GGT alone. However, from the perspectives of statistical stability and model interpretation, integrating GGT and T-Bil into a single composite may provide several interpretative and statistical characteristics distinct from evaluating GGT alone:

First, GBR was stable across missing-data assumptions. In CCA versus MICE comparisons (S3 Table), GGT alone varied from 0.1853 to 0.2007, whereas GBR changed minimally from 0.2006 to 0.2037.

Second, GBR provided a parsimonious fit: Pseudo-BIC was lower for GBR than for GGT alone (92,795.85 vs. 92,803.73; S2 Table).

Third, the bilirubin component provides complementary biological information. After adjusting for GGT, log(T-Bil) exhibited an independent and inverse association (Fig. 2) (β = -0.087, p = 0.009). This supports the rationale of integrating GGT and T-Bil as GBR.

Fourth, GBR provides a physiologically interpretable summary of the balance between the GGT and bilirubin components. These findings do not demonstrate clear predictive superiority over GGT alone, but suggest that bilirubin adds complementary biological information to GGT-related signals.

Independence and complementarity relative to existing oxidative balance indices
GBR remained associated with PHQ-9 independently of OBS and CDAI. Because OBS reflects diet- and lifestyle-derived oxidative balance [6, 7], whereas GBR may reflect downstream metabolic state, these measures may be complementary. In joint models (S4 Table), attenuation was modest for both GBR and OBS.

Heterogeneity of effects by BMI
BMI-stratified analyses suggested larger associations in normal-weight and overweight individuals and smaller associations in obese individuals, possibly because obesity-related inflammation, insulin resistance, hepatic impairment, and lifestyle factors dilute the GBR-specific signal. However, the continuous BMI interaction was not significant (β = -0.0040, p = 0.466), so BMI-related effect modification remains speculative.

Racial/ethnic specificity
One of the most notable findings was the racial/ethnic heterogeneity of the association between GBR and depressive symptoms. The association was clear in Non-Hispanic White (NHW) and Non-Hispanic Black (NHB) participants but absent in Hispanic and Asian participants, and this pattern was supported by a formal global interaction test (χ² = 15.05, df = 5, p = 0.010). The pattern also persisted after adjustment for major confounders, including alcohol consumption, smoking history, BMI, diabetes, CVD, and statin use, and in sensitivity analyses excluding antidepressant users. Therefore, this racial/ethnic difference may not be fully explained by measured clinical and lifestyle factors alone.

The component-specific analyses provide a possible clue. In NHW and NHB participants, the direction of the component estimates was consistent with the redox-balance hypothesis: GGT showed positive associations, bilirubin showed inverse associations, and the integrated GBR showed positive associations with depressive symptoms. In contrast, the Hispanic and Asian subgroups did not show the same component pattern, suggesting that the physiological meaning of bilirubin, GGT, or their ratio may differ across populations. This is important because it suggests that the validity of routine redox-related biomarkers may not be population-invariant.

Several speculative biological mechanisms may contribute to these discrepant findings. First, differences in genetic backgrounds related to bilirubin metabolism, particularly polymorphisms in the _UGT1A1_ gene, may play a role. The _UGT1A1_*28 allele, which contributes to the Gilbert syndrome phenotype characterized by elevated serum bilirubin, is more prevalent in European White and Black populations but less common in East Asian and some Hispanic populations [14, 27]. In Asian populations, the _UGT1A1_*6 polymorphism more commonly contributes to the Gilbert phenotype [27, 28]. Differences in the distribution of these genetic variants may alter the biological significance of bilirubin as an antioxidant-related marker. Second, specific metabolic features and susceptibility to fatty liver disease in Hispanic populations may contribute. Hispanic populations have a high frequency of the _PNPLA3_ I148M variant associated with lipid droplet accumulation [29, 30], which may increase susceptibility to non-alcoholic fatty liver disease and influence GGT elevation and bilirubin dynamics. This metabolic vulnerability may alter the physiological interpretation of GBR across populations.

Because genetic polymorphisms and related metabolic pathways were not directly assessed, these interpretations remain hypothesis-generating. Nonetheless, they emphasize that even routine clinical chemistry markers may have population-specific biological meanings. External validation in Hispanic and Asian populations is therefore urgent, ideally in cohorts incorporating _UGT1A1_ variants, insulin resistance, liver fat or NAFLD markers, and relevant lifestyle or environmental exposures.

Interaction between GBR and antidepressant use
The significant positive interaction between log(GBR) and antidepressant use is noteworthy (S5 Table). Although the association between GBR and PHQ-9 was significant among antidepressant non-users, it was significantly stronger in antidepressant users. Several interpretations may explain this finding. First, antidepressant use may serve as a proxy not only for drug exposure itself but also for depressive symptom severity, chronicity, healthcare access, and treatment resistance. In populations with these clinical backgrounds, the link between the PHQ-9 score and the oxidative stress balance reflected by GBR may become more pronounced. Second, drug-induced liver enzyme fluctuations may be involved. Certain antidepressants undergo hepatic metabolism and can be associated with elevated liver enzymes or drug-induced liver injury [31]. Thus, a portion of the stronger association observed in antidepressant users might be explained by drug exposure and hepatic metabolism signals contained within the GGT component of GBR. However, caution is warranted before interpreting this result simply as GBR capturing antidepressant-related liver injury. The association between GBR and PHQ-9 was maintained after excluding antidepressant users (β = 0.216, p = 1.7 × 10⁻⁶) and remained clearly apparent when restricting the analysis to participants with normal ALT/AST levels (β = 0.241, p = 1.8 × 10⁻⁵). Nonetheless, these interpretations remain exploratory and hypothesis-generating, and require validation using longitudinal data incorporating specific drug classes, dosages, treatment duration, and repeated liver enzyme measurements.

Effect size and epidemiological relevance
The effect size was modest at the individual level: a 1-unit increase in log(GBR) corresponded to a 0.256-point higher PHQ-9 score and a 16.7% higher odds of PHQ-9 ≥ 10. Its relevance is therefore primarily epidemiological: GBR can be calculated from routine GGT and T-Bil measurements to characterize population-level variation in depressive symptom burden without additional assay costs.

Limitations
This study has several limitations. First, its cross-sectional design precludes causal inference or determination of temporal direction. Longitudinal studies are needed to clarify whether redox imbalance precedes depressive symptoms, follows them, or both. Second, MICE assumes missing at random and cannot exclude missing-not-at-random mechanisms related to depression or health status; MNAR sensitivity analyses represent an important direction for future research [32, 33, 39]. The bounded and potentially skewed distribution of PHQ-9 scores should also be considered when interpreting linear-model estimates. Third, genetic variants such as _UGT1A1_*28, _UGT1A1_*6, and _PNPLA3_ were unavailable in public NHANES files, and detailed environmental and cultural factors were unmeasured, limiting interpretation of race/ethnicity-specific findings. Fourth, single-time-point GGT and bilirubin measurements may not capture chronic redox balance. Fifth, bilirubin fractions were unavailable, so total bilirubin was used despite indirect bilirubin being a potentially more specific antioxidant proxy. Sixth, residual confounding by genetics, micronutrient status, gut microbiota, or other unmeasured factors remains possible; the logistic E-value was 1.61. Seventh, hsCRP was available only in 2015–2018, although sensitivity analyses adjusting for NLR and hsCRP yielded similar results.

In addition, although sampling weights and PSU-based robust variance estimation were incorporated, the analyses did not fully account for the complete NHANES complex survey stratification structure. Therefore, some degree of residual variance estimation bias cannot be excluded.

Conclusions
In this nationally representative NHANES study, GBR was independently associated with depressive symptoms, and the association differed by race/ethnicity. GBR is a feasible composite marker for epidemiological research on systemic redox biology, but the cross-sectional design, modest effect size, and limited predictive advantage over GGT alone warrant caution. The findings should be interpreted as hypothesis-generating rather than clinically predictive. Future work should validate GBR in independent Asian and Hispanic cohorts, incorporate genetic variation in bilirubin and glutathione pathways, and evaluate temporal stability and longitudinal prediction. More broadly, these results suggest that translating routine redox-related biomarkers into psychiatric epidemiology requires attention to population-specific biological context.

Author contributions
Conceptualization: Takuya Matsuoka.
Data curation: Takuya Matsuoka.
Formal analysis: Takuya Matsuoka.
Methodology: Takuya Matsuoka.
Software: Takuya Matsuoka.
Visualization: Takuya Matsuoka.
Writing – original draft: Takuya Matsuoka.
Writing – review & editing: Takuya Matsuoka.

Funding
The author received no specific funding for this work.

Competing interests
The author has declared that no competing interests exist.

Use of generative AI and AI-assisted technologies
During manuscript preparation, the author used OpenAI's ChatGPT/Codex for language editing, drafting assistance, and code review. The author reviewed, revised, and approved all AI-assisted outputs and takes full responsibility for the final content of the manuscript.

Data availability
The dataset analyzed in the current study is publicly available from the National Health and Nutrition Examination Survey (NHANES) website hosted by the Centers for Disease Control and Prevention (CDC) at https://www.cdc.gov/nchs/nhanes/index.htm. All analysis scripts used to generate the results are available at https://github.com/mattakuya/NHANES_GBR.

Supporting information
**S1 Checklist. STROBE checklist.** Completed reporting checklist for cross-sectional studies.
**S1 Fig. Participant selection flowchart.** Selection of eligible adults from NHANES 2007–2018.
**S2 Fig. Demographic subgroup analyses.** Adjusted beta coefficients and 95% confidence intervals across BMI, sex, and age strata.
**S3 Fig. Lifestyle and clinical subgroup analyses.** Adjusted beta coefficients and 95% confidence intervals across alcohol, serum cotinine, and hsCRP-related analyses.
**S4 Fig. Antidepressant use interaction.** Predicted PHQ-9 total score by log(GBR), stratified by antidepressant use.
**S1 Table. Multicollinearity diagnostics for main model covariates.** Variance inflation factors (VIF) for all covariates adjusted in the primary analysis model.
**S2 Table. Model fit comparison of the primary exposure variables.** Comparison of Akaike Information Criterion (AIC), Bayesian Information Criterion (BIC), Adjusted R², and Weighted Root Mean Square Error (RMSE) across models containing GBR only, GGT only, or GGT and total bilirubin concurrently.
**S3 Table. Complete case analysis versus multiple imputation by chained equations (MICE).** Assessment of potential selection bias by comparing standardized beta coefficients for GGT alone and the GGT/bilirubin ratio under complete case analysis and multiple imputation frameworks.
**S4 Table. Associations with established oxidative stress indicators and joint models.** Standardized beta coefficients for the GGT/bilirubin ratio, GGT alone, CDAI, nutrient-only OBS, and full OBS, alongside joint model analysis and interaction testing in the OBS-restricted cohort.
**S5 Table. Antidepressant interaction analyses.** Detailed regression coefficients and interaction term estimates for the joint and BMI-stratified models exploring the modification of GBR and GGT associations by antidepressant use.
**S6 Table. Race/ethnicity interaction analysis.** Pooled interaction coefficients and global Wald test for log(GBR) × race/ethnicity, using Non-Hispanic White participants as the reference group.
**S7 Table. Missingness of primary analysis variables before multiple imputation.** Variable-level availability and missingness in the analytic adult sample before MICE.

References
1. Maes M, Galecki P, Chang YS, Berk M. A review on the oxidative and nitrosative stress (O&NS) pathways in major depression and their possible contribution to the (neuro)degenerative processes in that illness. Prog Neuropsychopharmacol Biol Psychiatry. 2011;35(3):676-692. doi:10.1016/j.pnpbp.2010.05.004.
2. Black CN, Bot M, Scheffer PG, Cuijpers P, Penninx BWJH. Is depression associated with increased oxidative stress? A systematic review and meta-analysis. Psychoneuroendocrinology. 2015;51:164-175. doi:10.1016/j.psyneuen.2014.09.025.
3. Liu T, Zhong S, Liao X, Chen J, He T, Lai S, Jia Y. A meta-analysis of oxidative stress markers in depression. PLoS One. 2015;10(10):e0138904. doi:10.1371/journal.pone.0138904.
4. Forlenza MJ, Miller GE. Increased serum levels of 8-hydroxy-2'-deoxyguanosine in clinical depression. Psychosom Med. 2006;68(1):1-7. doi:10.1097/01.psy.0000195780.37277.2a.
5. Frijhoff J, Winyard PG, Zarkovic N, Davies SS, Stocker R, Cheng D, et al. Clinical relevance of biomarkers of oxidative stress. Antioxid Redox Signal. 2015;23(14):1144-1170. doi:10.1089/ars.2015.6317.
6. Hernandez-Ruiz A, Garcia-Villanova B, Guerra-Hernandez EJ, Amiano P, Azpiri M, Molina-Montes E. A review of a priori defined oxidative balance scores relative to their components and impact on health outcomes. Nutrients. 2019;11(4):774. doi:10.3390/nu11040774.
7. Goodman M, Bostick RM, Dash C, Terry P, Flanders WD, Mandel J. A summary measure of pro- and anti-oxidant exposures and risk of incident, sporadic, colorectal adenomas. Cancer Causes Control. 2008;19(10):1051-1064. doi:10.1007/s10552-008-9169-y.
8. Wright ME, Mayne ST, Stolzenberg-Solomon RZ, Li Z, Pietinen P, Taylor PR, et al. Development of a comprehensive dietary antioxidant index and application to lung cancer risk in a cohort of male smokers. Am J Epidemiol. 2004;160(1):68-76. doi:10.1093/aje/kwh173.
9. Liu X, Liu X, Wang Y, Zeng B, Zhu B, Dai F. Association between depression and oxidative balance score: National Health and Nutrition Examination Survey (NHANES) 2005-2018. J Affect Disord. 2023;337:57-65. doi:10.1016/j.jad.2023.05.071.
10. Li H, Song L, Cen M, Fu X, Gao X, Zuo Q, Wu J. Oxidative balance scores and depressive symptoms: Mediating effects of oxidative stress and inflammatory factors. J Affect Disord. 2023;334:205-212. doi:10.1016/j.jad.2023.04.134.
11. Lee DH, Blomhoff R, Jacobs DR Jr. Is serum gamma glutamyltransferase a marker of oxidative stress? Free Radic Res. 2004;38(6):535-539. doi:10.1080/10715760410001694026.
12. Takigawa T, Hibino Y, Kimura S, Yamauchi H, Wang B, Wang D, Ogino K. Association between serum gamma-glutamyltransferase and oxidative stress related factors. Hepatogastroenterology. 2008;55(81):50-53.
13. Stocker R, Yamamoto Y, McDonagh AF, Glazer AN, Ames BN. Bilirubin is an antioxidant of possible physiological importance. Science. 1987;235(4792):1043-1046. doi:10.1126/science.3029864.
14. Schwertner HA, Vitek L. Gilbert syndrome, UGT1A1*28 allele, and cardiovascular disease risk: possible protective effects and therapeutic applications of bilirubin. Atherosclerosis. 2008;198(1):1-11. doi:10.1016/j.atherosclerosis.2008.01.001.
15. Abbasi A, Deetman PE, Corpeleijn E, Gansevoort RT, Gans ROB, Hillege HL, et al. Bilirubin as a potential causal factor in type 2 diabetes risk: a Mendelian randomization study. Diabetes. 2015;64(4):1459-1469. doi:10.2337/db14-0228.
16. Vitek L, Jirsa M, Brodanova M, Kalab M, Marecek Z, Danzig V, et al. Gilbert syndrome and ischemic heart disease: a protective effect of elevated bilirubin levels. Atherosclerosis. 2002;160(2):449-456. doi:10.1016/S0021-9150(01)00601-3.
17. Horsfall LJ, Rait G, Walters K, Swallow DM, Pereira SP, Nazareth I, Petersen I. Serum bilirubin and risk of respiratory disease and death. JAMA. 2011;305(7):691-697. doi:10.1001/jama.2011.124.
18. Itoh K, Wakabayashi N, Katoh Y, Ishii T, Igarashi K, Engel JD, Yamamoto M. Keap1 represses nuclear activation of antioxidant responsive elements by Nrf2 through binding to the amino-terminal Neh2 domain. Genes Dev. 1999;13(1):76-86. doi:10.1101/gad.13.1.76.
19. Kensler TW, Wakabayashi N, Biswal S. Cell survival responses to environmental stresses via the Keap1-Nrf2-ARE pathway. Annu Rev Pharmacol Toxicol. 2007;47:89-116. doi:10.1146/annurev.pharmtox.46.120604.141046.
20. Loboda A, Damulewicz M, Pyza E, Jozkowicz A, Dulak J. Role of Nrf2/HO-1 system in development, oxidative stress response and diseases: an evolutionarily conserved mechanism. Cell Mol Life Sci. 2016;73(17):3221-3247. doi:10.1007/s00018-016-2223-0.
21. Rengasamy M, Price R. Replicable and robust cellular and biochemical blood marker signatures of depression and depressive symptoms. Psychiatry Res. 2024;342:116190. doi:10.1016/j.psychres.2024.116190.
22. Peng YF, Xiang Y, Wei YS. The significance of routine biochemical markers in patients with major depressive disorder. Sci Rep. 2016;6:34402. doi:10.1038/srep34402.
23. National Center for Health Statistics. National Health and Nutrition Examination Survey: Analytic Guidelines, 1999-2010. Hyattsville, MD: Centers for Disease Control and Prevention; 2013.
24. Kroenke K, Spitzer RL, Williams JBW. The PHQ-9: Validity of a brief depression severity measure. J Gen Intern Med. 2001;16(9):606-613. doi:10.1046/j.1525-1497.2001.016009606.x.
25. Rubin DB. Multiple Imputation for Nonresponse in Surveys. New York: John Wiley & Sons; 1987. doi:10.1002/9780470316696.
26. Barnard J, Rubin DB. Small-sample degrees of freedom with multiple imputation. Biometrika. 1999;86(4):948-955. doi:10.1093/biomet/86.4.948.
27. Gammal RS, Court MH, Haidar CE, Iwuchukwu OF, Gaur AH, Alvarellos M, et al. Clinical Pharmacogenetics Implementation Consortium (CPIC) guideline for UGT1A1 and atazanavir prescribing. Clin Pharmacol Ther. 2016;99(4):363-369. doi:10.1002/cpt.269.
28. Shimoyama S. Pharmacogenetics of irinotecan: An ethnicity-based prediction of irinotecan adverse events. World J Gastrointest Surg. 2010;2(1):14-21. doi:10.4240/wjgs.v2.i1.14.
29. Romeo S, Kozlitina J, Xing C, Pertsemlidis A, Cox D, Pennacchio LA, et al. Genetic variation in PNPLA3 confers susceptibility to nonalcoholic fatty liver disease. Nat Genet. 2008;40(12):1461-1465. doi:10.1038/ng.257.
30. Sulaiman SA, Dorairaj V, Adrus MNH. Genetic polymorphisms and diversity in nonalcoholic fatty liver disease (NAFLD): a mini review. Biomedicines. 2023;11(1):106. doi:10.3390/biomedicines11010106.
31. Voican CS, Corruble E, Naveau S, Perlemuter G. Antidepressant-induced liver injury: a review for clinicians. Am J Psychiatry. 2014;171(4):404-415. doi:10.1176/appi.ajp.2013.13050709.
32. National Research Council (US) Panel on Handling Missing Data in Clinical Trials. The Prevention and Treatment of Missing Data in Clinical Trials. Washington (DC): National Academies Press (US); 2010.
33. Sterne JAC, White IR, Carlin JB, Spratt M, Royston P, Kenward MG, et al. Multiple imputation for missing data in epidemiological and clinical research: potential and pitfalls. BMJ. 2009;338:b2393. doi:10.1136/bmj.b2393.
34. Sedlak TW, Snyder SH. Bilirubin benefits: cellular protection by a biliverdin reductase antioxidant cycle. Pediatrics. 2004;113(6):1776-1782. doi:10.1542/peds.113.6.1776.
35. Davignon J. Beneficial cardiovascular pleiotropic effects of statins. Circulation. 2004;109(23 Suppl 1):III39-III43. doi:10.1161/01.CIR.0000131517.20177.5a.
36. Zahorec R. Ratio of neutrophil to lymphocyte counts--rapid and simple parameter of systemic inflammation and stress in critically ill. Bratisl Lek Listy. 2001;102(1):5-14.
37. Little RJA, Rubin DB. Statistical Analysis with Missing Data. 2nd ed. Hoboken, NJ: Wiley; 2002. doi:10.1002/9781119013563.
38. White IR, Royston P, Wood AM. Multiple imputation using chained equations: issues and guidance for practice. Stat Med. 2011;30(4):377-399. doi:10.1002/sim.4067.
39. Carpenter JR, Kenward MG, White IR. Sensitivity analysis after multiple imputation under missing at random: a weighting approach. Stat Methods Med Res. 2007;16(3):259-275. doi:10.1177/0962280206075303.
"""
    # Replace table placeholders (if they aren't already replaced or need extra appending)
    text = text.replace("{table1}", table1)
    text = text.replace("{table2}", table2)
    
    # In manuscript_tp.md, Table 3 and Table 4 might need to be appended at the end of Results
    # or replaced if placeholders are present. Since results/manuscript_tp.md doesn't have [Insert Table 3 here],
    # let's append Table 3 and Table 4 definitions at the end of Results section or just before Discussion.
    # Alternatively, we can find where to replace them.
    # Let's check if they are in en_manuscript:
    if "{table3}" not in text:
        # Append them before Discussion
        split_text = text.split("## Discussion")
        if len(split_text) == 2:
            middle = """

### Tables

#### Table 3. Subgroup-stratified analyses of the association between log(GBR) and depressive symptoms (WLS models)
{table3}

#### Table 4. Independent and additive associations of GBR and OBS (standardized per 1 SD)
{table4}

"""
            text = split_text[0] + middle + "## Discussion" + split_text[1]
        else:
            text += """

#### Table 3. Subgroup-stratified analyses of the association between log(GBR) and depressive symptoms (WLS models)
{table3}

#### Table 4. Independent and additive associations of GBR and OBS (standardized per 1 SD)
{table4}"""
    else:
        text = text.replace("{table3}", table3)
        text = text.replace("{table4}", table4)

    return text


def get_supplementary_text(table1, table2, table3):
    text = f'''# Association of the GGT-to-bilirubin ratio with depressive symptoms in US adults: an NHANES 2007–2018 study
## Supplementary Material

**Takuya Matsuoka**

---

### Supplementary Tables

#### S1 Table. Specificity analysis of GBR and other liver enzymes (standardized per 1 SD)
{table1}

#### S2 Table. High-sensitivity C-reactive protein (hs-CRP) sensitivity analysis in cycles I/J
{table2}

#### S3 Table. Bias assessment of complete case analysis (CCA) versus multiple imputation by chained equations (MICE) (standardized per 1 SD)
{table3}

---
'''
    return text

def apply_fonts_to_run(run, bold=False, size_pt=11):
    """
    Maps alphanumeric characters to Helvetica and East Asian characters to Meiryo.
    """
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Helvetica')
    rFonts.set(qn('w:hAnsi'), 'Helvetica')
    rFonts.set(qn('w:eastAsia'), 'Meiryo')
    rPr.append(rFonts)
    
    run.font.size = Pt(size_pt)
    if bold:
        run.bold = True

def convert_md_to_docx(md_path, docx_path):
    print(f"Converting {md_path} to {docx_path}...")
    doc = Document()
    
    # Margin settings (1 inch = Inches(1))
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
        table.style = 'Table Grid'
        
        for r_idx, row_data in enumerate(rows_data):
            for c_idx, cell_value in enumerate(row_data):
                if c_idx < len(table.rows[r_idx].cells):
                    cell = table.rows[r_idx].cells[c_idx]
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)
                    
                    is_header = (r_idx == 0) or cell_value.startswith("**")
                    clean_val = cell_value.replace("**", "")
                    
                    run = p.add_run(clean_val)
                    apply_fonts_to_run(run, bold=is_header, size_pt=10)
                    
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
            
        # Heading 1
        h1_match = re.match(r'^#\s+(.*)', line)
        if h1_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(h1_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=18)
            i += 1
            continue
            
        # Heading 2
        h2_match = re.match(r'^##\s+(.*)', line)
        if h2_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(h2_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=14)
            i += 1
            continue
            
        # Heading 3
        h3_match = re.match(r'^###\s+(.*)', line)
        if h3_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(h3_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=12)
            i += 1
            continue
            
        # Heading 4
        h4_match = re.match(r'^####\s+(.*)', line)
        if h4_match:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(h4_match.group(1))
            apply_fonts_to_run(run, bold=True, size_pt=11)
            i += 1
            continue
            
        # Normal Paragraph with bold parsing
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

def main():
    try:
        results_file = get_latest_results_file()
        print(f"Reading results from: {results_file}")
        
        output_dir = "results"
        os.makedirs(output_dir, exist_ok=True)
        
        # 1. Parse table 1
        table1 = parse_markdown_table(results_file, "DESCRIPTIVE TABLE 1 BY LOG_RATIO QUARTILE")
        
        # 2. Build table 2 (Primary & Sensitivity)
        table2 = build_table_2(results_file)
        
        # 3. Build table 3 (Subgroups) -> Supplementary Table 1
        table3_sub = build_table_3_subgroups(results_file)
        
        # 4. Build table 4 (Specificity) -> Supplementary Table 2
        table4_spec = build_table_3(results_file)
        
        # 5. Build table 5 (Joint model GBR + OBS) -> Main Table 3
        table5_joint = build_table_4(results_file)
        
        # 6. Build additional supplementary tables (hsCRP and CCA vs MICE)
        table_hscrp = build_table_hscrp_sensitivity(results_file)
        table_ccamice = build_table_cca_vs_mice(results_file)

        # 6.5. Compile main manuscript.md
        manuscript_content = get_manuscript_text(table1, table2, table3_sub, table5_joint)
        output_path = "manuscript.md"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(manuscript_content)
        print(f"Main manuscript successfully compiled and saved to: {os.path.abspath(output_path)}")
        
        # 7. Compile supplementary_material.md
        supp_content = get_supplementary_text(table4_spec, table_hscrp, table_ccamice)
        supp_path = "supplementary_material.md"
        with open(supp_path, "w", encoding="utf-8") as f:
            f.write(supp_content)
        print(f"Supplementary materials successfully compiled and saved to: {os.path.abspath(supp_path)}")
        
        # 7.5. Convert markdown files to Word documents
        convert_md_to_docx("manuscript.md", "manuscript.docx")
        convert_md_to_docx("supplementary_material.md", "supplementary_material.docx")
        
        # 8. Copy and rename professional table/figure PNGs under results/publication_assets
        assets_dir = os.path.join(output_dir, "publication_assets")
        os.makedirs(assets_dir, exist_ok=True)
        
        # Mapping for main figures and tables
        file_mapping = {
            # Tables
            "table1_characteristics.png": "table1_characteristics.png",
            "table2_association.png": "table2_primary_sensitivity.png",
            "table3_subgroups.png": "table3_subgroups.png",
            "table5_gbr_obs.png": "table4_gbr_obs.png",
            "table4_specificity.png": "supp_table1_specificity.png",
            "supp_table3_hscrp.png": "supp_table2_hscrp.png",
            "supp_table4_ccamice.png": "supp_table3_ccamice.png",
            # Figures
            "figure1_forest_plot.png": "figure1_subgroups.png",
            "figure2_specificity_plot.png": "figure2_specificity.png",
            "rcs_log_ratio_phq9.png": "figure3_rcs_spline.png",
            "figure4_interaction_plot.png": "figure4_interaction.png",
            "figure5_forest_plot_latter.png": "figure5_subgroups_latter.png",
            "figure6_flowchart.png": "figure6_flowchart.png"
        }
        
        for orig, new_name in file_mapping.items():
            orig_path = os.path.join(output_dir, orig)
            new_path = os.path.join(assets_dir, new_name)
            if os.path.exists(orig_path):
                shutil.copy2(orig_path, new_path)
                print(f"Asset copied and renamed: {orig} -> publication_assets/{new_name}")
            else:
                print(f"Warning: Original asset not found: {orig_path}")
                
        print("All publication-ready assets successfully compiled and organized!")
        
    except Exception as e:
        print(f"Error compiling manuscript: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
