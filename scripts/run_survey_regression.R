library(survey)
library(mitools)

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 5) {
  stop("Missing arguments. Usage: Rscript run_survey_regression.R <imputed_dir> <m> <formula> <family> <output> [<output_cov> <output_coef> <output_metrics>]")
}

imputed_dir <- args[1]
m <- as.integer(args[2])
formula_str <- args[3]
family_str <- args[4]
output_file <- args[5]

output_cov <- NULL
if (length(args) >= 6 && args[6] != "NULL") {
  output_cov <- args[6]
}
output_coef <- NULL
if (length(args) >= 7 && args[7] != "NULL") {
  output_coef <- args[7]
}
output_metrics <- NULL
if (length(args) >= 8 && args[8] != "NULL") {
  output_metrics <- args[8]
}

# Set options for single-PSU strata
options(survey.lonely.psu = "adjust")

# Load imputed datasets
imp_list <- lapply(0:(m-1), function(i) {
  read.csv(file.path(imputed_dir, paste0("imputed_", i, ".csv")))
})

# Define survey design using mitools::imputationList
designs <- svydesign(
  id = ~SDMVPSU_c,
  strata = ~SDMVSTRA_c,
  weights = ~WTMEC2YR,
  data = imputationList(imp_list),
  nest = TRUE
)

# Fit GLM across imputations
formula_obj <- as.formula(formula_str)
if (family_str == "binomial") {
  fam <- binomial()
} else {
  fam <- gaussian()
}

# Run svyglm on each imputed dataset
fit_list <- with(designs, svyglm(formula_obj, family = fam))
pooled <- MIcombine(fit_list)

# Extract summary and CI
sum_pooled <- summary(pooled)
ci <- confint(pooled)

# Calculate p-value manually using pooled degrees of freedom
t_stats <- sum_pooled$results / sum_pooled$se
p_vals <- 2 * pt(-abs(t_stats), df = pooled$df)

# Build results data frame
results_df <- data.frame(
  term = rownames(sum_pooled),
  beta = sum_pooled$results,
  se = sum_pooled$se,
  ci_lo = ci[, 1],
  ci_hi = ci[, 2],
  p = p_vals,
  stringsAsFactors = FALSE
)

# Save results summary
write.csv(results_df, output_file, row.names = FALSE)

# Save covariance matrix if requested
if (!is.null(output_cov)) {
  cov_mat <- as.matrix(vcov(pooled))
  write.csv(cov_mat, output_cov, row.names = TRUE)
}

# Save coefficient vector if requested
if (!is.null(output_coef)) {
  coef_vec <- coef(pooled)
  coef_df <- data.frame(
    term = names(coef_vec),
    beta = as.numeric(coef_vec),
    stringsAsFactors = FALSE
  )
  write.csv(coef_df, output_coef, row.names = FALSE)
}

# Save model fit metrics if requested (only makes sense for gaussian OLS)
if (!is.null(output_metrics)) {
  # Since fit_list is a list of svyglm objects from mitools::with
  # We extract individual fits to compute fit metrics per imputation
  metrics <- lapply(fit_list, function(f) {
    resid <- residuals(f, type = "response")
    w <- weights(f, type = "prior")
    
    # Valid elements
    valid <- !is.na(w) & w > 0 & !is.na(resid)
    wv <- w[valid]
    rv <- resid[valid]
    
    n_eff <- sum(valid)
    k <- length(coef(f))
    
    rss_w <- sum(wv * (rv ^ 2))
    sum_w <- sum(wv)
    sigma2_w <- if (sum_w > 0) rss_w / sum_w else NA
    
    pseudo_aic <- n_eff * log(sigma2_w) + 2 * k
    pseudo_bic <- n_eff * log(sigma2_w) + k * log(n_eff)
    rmse_w <- sqrt(sum(wv * (rv ^ 2)) / sum_w)
    
    # Adjusted R2 calculation
    y <- f$y
    if (is.null(y)) {
      y <- f$model[[1]]
    }
    yv <- y[valid]
    y_mean <- sum(wv * yv) / sum_w
    tss_w <- sum(wv * (yv - y_mean)^2)
    
    r2 <- 1 - (rss_w / tss_w)
    adj_r2 <- 1 - ((1 - r2) * (n_eff - 1) / (n_eff - k))
    
    data.frame(
      Pseudo_AIC = pseudo_aic,
      Pseudo_BIC = pseudo_bic,
      Adj_R2 = adj_r2,
      Weighted_RMSE = rmse_w
    )
  })
  
  metrics_df <- do.call(rbind, metrics)
  
  # Format output
  metrics_summary <- data.frame(
    Metric = c("Pseudo_AIC", "Pseudo_BIC", "Adj_R2", "Weighted_RMSE"),
    Mean = colMeans(metrics_df, na.rm = TRUE),
    SD = apply(metrics_df, 2, sd, na.rm = TRUE),
    stringsAsFactors = FALSE
  )
  write.csv(metrics_summary, output_metrics, row.names = FALSE)
}

cat("R REGRESSION SUCCESS\n")
