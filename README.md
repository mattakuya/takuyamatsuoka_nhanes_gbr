# NHANES GBR Depression Analysis

[![Python Version](https://img.shields.io/badge/python-3.14+-blue.svg)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

This repository contains the official, fully reproducible research-grade code for investigating the association between the **gamma-glutamyl transferase-to-total bilirubin ratio (GBR)**—a routine biochemistry-derived proxy for systemic redox balance—and **depressive symptoms (PHQ-9)** in US adults.

The analysis is based on composite cross-sectional data from the **National Health and Nutrition Examination Survey (NHANES) 2007–2018** (*N* = 36,580). 

---

## 1. Study Background & Methodology

Oxidative stress is widely implicated in the pathophysiology of depression. GBR represents a composite proxy where:
* **Gamma-glutamyl transferase (GGT)** serves as an indicator of pro-oxidant burden and extracellular glutathione turnover.
* **Bilirubin (T-Bil)** serves as an endogenous, lipophilic antioxidant.

### Statistical Framework
1. **Complex Survey Design**: Handles cluster-robust standard errors using cycle-specific composite Primary Sampling Units (`SDMVPSU_c`) and Strata (`SDMVSTRA_c`) with composite MEC examination weights (`WTMEC2YR`), compliant with CDC guidelines.
2. **Missing Data (MICE)**: Addresses missingness via Multiple Imputation by Chained Equations (*m* = 20, max iterations = 30) using scikit-learn's `IterativeImputer`.
3. **Pooling**: Pools estimates and standard errors using **Rubin's Rules** combined with the **Barnard-Rubin (1999)** degrees-of-freedom correction for small-sample/survey settings.
4. **Primary Models**: Weighted Least Squares (WLS) regression of continuous PHQ-9 scores.
5. **Sensitivity & Subgroup Analyses**:
   * Stratifications by BMI, Sex, Age, and Race/Ethnicity (with formal Wald interaction test).
   * Exclusions of antidepressant users, cardiovascular disease, diabetes, and participants with elevated ALT/AST (liver-safe restriction).
   * Direct comparison with established oxidative/antioxidant scores: Oxidative Balance Score (OBS) and Clinically Derived Antioxidant Index (CDAI).

---

## 2. Repository Structure

```tree
.
├── pyproject.toml              # Project dependency declarations (PEP 621 compliant)
├── uv.lock                     # Lockfile for precise dependency resolution
├── README.md                   # Repository documentation
├── .gitignore                  # Git ignore definitions (ignores large raw files/caches)
├── .python-version             # Local python target (3.14)
├── scripts/                    # Python pipeline and visual execution scripts
│   ├── download_nhanes.py      # Automated CDC NHANES raw dataset downloader
│   ├── analysis_pipeline.py    # Main statistical pipeline (Preprocess -> MICE -> WLS -> Pooling)
│   ├── analysis_pipeline_quick.py # Dry-run pipeline helper (low m and max_iter for testing)
│   ├── generate_flowchart.py   # Renders participant inclusion flowchart (S1 Fig)
│   ├── generate_forest_plot.py # Renders Table 2 / Table 3 forest plot (Figure 1)
│   ├── generate_rcs_plot.py    # Renders Restricted Cubic Spline (RCS) curves (Figure 3)
│   ├── generate_docx_tables.py # Compiles Markdown clinical tables into formatted Word tables
│   ├── generate_manuscript.py  # Builds final Word manuscript with embedded tables/captions
│   ├── generate_strobe_checklist.py # Generates cross-sectional STROBE reporting checklist
│   └── generate_missingness_table.py # Evaluates raw missingness before imputation (S7 Table)
└── results/                    # Output directory for plots, tables, and manuscripts (ignored in git)
```

---

## 3. Installation & Setup

We recommend using [uv](https://github.com/astral-sh/uv) (a fast Python package installer and resolver) to run the code in a virtual environment.

### Using `uv` (Recommended)
```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/NHANES_GBR.git
cd NHANES_GBR

# Create a virtual environment and install all dependencies
uv sync
```

### Using standard `venv` & `pip`
```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/NHANES_GBR.git
cd NHANES_GBR

# Create and activate virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install dependencies
pip install .
```

---

## 4. Pipeline Execution Workflow

Follow these steps sequentially to reproduce the entire analysis, including all manuscripts, supplementary documents, tables, and figures.

### Step 1: Raw Data Acquisition
Execute the batch downloader to retrieve raw `.XPT` SAS Transport files directly from the CDC website:
```bash
python scripts/download_nhanes.py
```
*Raw files will be saved under `data/raw/[cycle_years]/`.*

### Step 2: Running the Statistical Pipeline
Run the main pipeline to preprocess datasets, perform MICE multiple imputation (*m* = 20), run survey WLS models, pool results, and output raw markdown tables:
```bash
python scripts/analysis_pipeline.py
```
* **Performance Control**: By default, the script dynamically scales parallel jobs (`n_jobs`) based on your system's RAM and CPU core count:
  * On high-end systems (>= 60 GB RAM, such as M-series Max/Ultra Macs), it utilizes `cpu_count - 2` cores to maximize speed while maintaining system UI responsiveness.
  * On smaller systems (< 60 GB RAM), it caps jobs at `cpu_count - 1` (up to 4 cores) to prevent memory peaks.
  * You can override this by setting the `NHANES_JOBS` environment variable:
    ```bash
    # Example: Run sequentially (maximum stability, lowest memory)
    export NHANES_JOBS=1
    python scripts/analysis_pipeline.py
    ```
* *For a quick test run to verify the pipeline runs without errors (using m = 1, max_iter = 1), execute `python scripts/analysis_pipeline_quick.py` instead.*

### Step 3: Generating Visualizations & Publication Assets
Once the pipeline has completed and written the statistical findings to `results/`, generate all paper figures and tables:

```bash
# Renders Figure 1 (Primary & Subgroup Forest Plot)
python scripts/generate_forest_plot.py

# Renders Figure 3 (Restricted Cubic Spline dose-response curve)
python scripts/generate_rcs_plot.py

# Generates Figure 6 (Participant Selection Flowchart)
python scripts/generate_flowchart.py

# Generates S7 Table (Pre-imputation missingness table)
python scripts/generate_missingness_table.py

# Generates clinical Tables document (tables_tp.docx) styled with APA booktabs format
python scripts/generate_docx_tables.py

# Combines all figures, tables, and STROBE checklists into a submission-ready manuscript
python scripts/generate_tp_manuscript.py
python scripts/generate_strobe_checklist.py
```

---

## 5. Licensing & Citation

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

If you find this codebase or study useful in your research, please cite:
```bibtex
@article{matsuoka2026nhanesgbr,
  title={Race- and ethnicity-specific association of the gamma-glutamyl transferase-to-bilirubin ratio with depressive symptoms in US adults: a NHANES 2007–2018 study},
  author={Matsuoka, Takuya},
  journal={Translational Psychiatry},
  year={2026},
  publisher={Nature Publishing Group}
}
```
