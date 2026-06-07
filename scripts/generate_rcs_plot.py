import os
import logging
from analysis_pipeline import NHANESDepressionAnalysis

logging.basicConfig(level=logging.INFO)

def main():
    try:
        dataset_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "raw")
        
        # Initialize pipeline; n_jobs=None resolves dynamically based on system memory and core counts
        pipeline = NHANESDepressionAnalysis(
            dataset_dir=dataset_path, n_jobs=4, imputer_max_iter=30)
        
        print("Loading data...")
        pipeline.load_multicycle_data(cycles=('E', 'F', 'G', 'H', 'I', 'J'))
        
        print("Preprocessing data...")
        pipeline.preprocess()
        
        print("Running RCS analysis (m=20) to render Figure 3 in Helvetica...")
        pipeline.run_rcs_analysis(m=20, n_knots=4, output_dir="results")
        
        # Also copy to results/tp_figures/figure3_rcs.png for medrxiv single docx
        tp_fig_dir = os.path.join("results", "tp_figures")
        os.makedirs(tp_fig_dir, exist_ok=True)
        import shutil
        shutil.copy2(
            os.path.join("results", "rcs_log_ratio_phq9.png"),
            os.path.join(tp_fig_dir, "figure3_rcs.png")
        )
        print("Copied to results/tp_figures/figure3_rcs.png")
        
        print("Figure 3 successfully regenerated!")
        
    except Exception as e:
        print(f"Error generating Figure 3: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
