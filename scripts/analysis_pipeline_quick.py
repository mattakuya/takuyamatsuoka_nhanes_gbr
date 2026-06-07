"""
Dry-run runner: Fast version with limited MICE m and IterativeImputer max_iter.
For the main production run, execute analysis_pipeline.py directly (m=20, max_iter=50).
"""
import logging
import os

from analysis_pipeline import NHANESDepressionAnalysis, run_full_pipeline

M_QUICK = 1
MAX_ITER_QUICK = 1

if __name__ == "__main__":
    try:
        dataset_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "raw")
        pipeline = NHANESDepressionAnalysis(
            dataset_dir=dataset_path,
            n_jobs=2,
            imputer_max_iter=MAX_ITER_QUICK,
        )
        run_full_pipeline(
            pipeline,
            m=M_QUICK,
            output_dir='results',
            output_prefix='nhanes_results_QUICK',
        )
    except Exception as e:
        logging.error(f"Quick analysis failed: {e}", exc_info=True)
