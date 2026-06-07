library(survey)
library(mitools)

# 1. Load imputed datasets exported by Python
message("Loading 20 imputed datasets from data/imputed/...")
imp_list <- lapply(0:19, function(i) {
  read.csv(paste0("data/imputed/imputed_", i, ".csv"))
})

# 2. Define the complex survey design
# SDMVPSU_c: composite cluster ID, SDMVSTRA_c: composite strata ID, WTMEC2YR: survey weights
message("Defining survey design using mitools::imputationList...")
designs <- svydesign(
  id = ~SDMVPSU_c,
  strata = ~SDMVSTRA_c,
  weights = ~WTMEC2YR,
  data = imputationList(imp_list),
  nest = TRUE
)

# 3. Fit survey-weighted generalized linear model (svyglm) for WLS
message("Fitting survey GLM with Taylor Series Linearization across imputations...")

sample_df <- imp_list[[1]]
race_cols <- grep("^race_", names(sample_df), value = TRUE)
edu_cols <- grep("^education_cat_", names(sample_df), value = TRUE)

covariates <- c(
  "RIDAGEYR", "is_female", "pir", "is_married", "total_calories", 
  "alcohol_drinks", "is_ever_smoker", "log_nlr", "sedentary_min", 
  "is_diabetic", "has_cvd", "is_ad", "is_statin", "has_liver_disease",
  race_cols, edu_cols
)

# WLS Model with GBR * BMI interaction
formula_str <- paste(
  "phq9_score ~ log_ratio_c * bmi_c +", 
  paste(covariates, collapse = " + ")
)
formula <- as.formula(formula_str)

# Run svyglm on each imputed dataset and pool estimates using Rubin's Rules
fit <- with(designs, svyglm(formula, family = gaussian()))
pooled <- MIcombine(fit)

message("\n============================================================")
message("=== R survey (Taylor Series Linearization) Pooled Results ===")
message("============================================================")
summary_fit <- summary(pooled)
print(summary_fit)

# Highlight key exposure variables
cat("\n--- Key Exposure: log(GBR) (log_ratio_c) ---\n")
gbr_idx <- which(rownames(summary_fit) == "log_ratio_c")
if (length(gbr_idx) > 0) {
  est <- summary_fit[gbr_idx, "results"]
  se <- summary_fit[gbr_idx, "se"]
  ci_lo <- est - 1.96 * se
  ci_hi <- est + 1.96 * se
  p_val <- summary_fit[gbr_idx, "(p)"]
  cat(sprintf("Beta: %.4f, SE: %.4f, 95%% CI: [%.4f, %.4f], p-value: %.2e\n", 
              est, se, ci_lo, ci_hi, p_val))
}

cat("\n--- Interaction: log(GBR) x BMI (log_ratio_c:bmi_c) ---\n")
int_idx <- which(rownames(summary_fit) == "log_ratio_c:bmi_c")
if (length(int_idx) > 0) {
  est <- summary_fit[int_idx, "results"]
  se <- summary_fit[int_idx, "se"]
  ci_lo <- est - 1.96 * se
  ci_hi <- est + 1.96 * se
  p_val <- summary_fit[int_idx, "(p)"]
  cat(sprintf("Beta: %.4f, SE: %.4f, 95%% CI: [%.4f, %.4f], p-value: %.2e\n", 
              est, se, ci_lo, ci_hi, p_val))
}
