lapply(c("janitor", "readxl", "tidyverse", "skimr","patchwork"), require, character.only = TRUE)
file.choose()
# Change to your file directory accordingly
file_path <- "C:\\Users\\Dickson Woo Teng Wan\\Documents\\R programming\\Asymposium Climate Risk Case Study 2025\\additional_dataset.xlsx"

excel_sheets(file_path)

crop_data <- read_excel(file_path, sheet = "agriculture products production") %>%
  tibble() %>%
  clean_names() %>%
  rename(
    production = production_000_metric_tons
  )

crop_gdp_data <- read_excel(file_path, sheet = "oil palm export volume & value") %>%
  tibble() %>%
  clean_names() %>%
  filter(str_detect(product, regex("palm", ignore_case = T)) & year >= 2016 & year <= 2020) %>%
  select(!sub_product) %>%
  rename(
    volume = volume_tonnes,
    value = value_rm_million
  ) %>%
  mutate(
    year = as.numeric(year),
    volume = as.numeric(volume),
    value = as.numeric(value)
  )


flood_sheets <- excel_sheets(file_path)[6:10]

flood_data <- setNames(
  lapply(flood_sheets, function(sheet) {
    df <- read_excel(file_path, sheet = sheet, skip = 1) %>%
      clean_names()
    
    if ("maximum_flood_duration_days" %in% names(df)) {
      df <- df %>%
        mutate(
          maximum_flood_duration_days = suppressWarnings(as.numeric(maximum_flood_duration_days)),
          maximum_flood_duration_hours = maximum_flood_duration_days * 24
        )
    }
    return(df)
  }),
  flood_sheets
)

# Combine all tables
combined_floods <- imap_dfr(flood_data, ~ {
  year <- str_extract(.y, "\\d{4}")  # Extract 4-digit year from sheet name
  
  .x %>%
    # Remove summary rows (case-insensitive)
    filter(!grepl("total|jumlah", state, ignore.case = TRUE)) %>%
    # Standardize column names and types
    transmute(
      state = as.character(state),
      year = as.integer(year),
      flood_incidents = as.numeric(number_of_flood_incidents),
      avg_rainfall_mm = as.numeric(average_maximum_rainfall_mm),
      flood_duration_hrs = as.numeric(maximum_flood_duration_hours),
      flood_depth_m = as.numeric(maximum_flood_depth_m)
    )
}) %>%
  arrange(year, state)  %>% # Sort chronologically by state
  na.omit()

# Summarize total planted area and production by crop type
pct_share <- crop_data %>%
  group_by(year, agriculture_product) %>%
  summarise(
    national_production = sum(production, na.rm = T),
    .groups = "drop"
  ) %>%
  group_by(year) %>%
  mutate(
    total_production = sum(national_production, na.rm = T),
    production_pct = (national_production / total_production) * 100
  ) %>%
  ungroup() %>%
  filter(year == max(year, na.rm = T)) %>%
  arrange(desc(production_pct)) %>%
  ggplot(aes(x = reorder(agriculture_product, production_pct), y = production_pct)) + 
  geom_col(fill = "steelblue") + 
  coord_flip() +
  geom_text(aes(label = paste(round(production_pct, 2), "%")),
            hjust = -0.1,
            size = 3.5) +
  labs(title = paste("Top Agricultural Products by Production Share (2010 - 2018)"),
       x = "Agriculture Products", 
       y = "Production Share (%)") +
  theme_minimal()

trend_analysis <- crop_data %>%
  group_by(year, agriculture_product) %>%
  summarise(
    national_production = sum(production, na.rm = TRUE), 
    .groups = "drop"
    ) %>%
  group_by(year) %>%
  mutate(
    total_production = sum(national_production, na.rm = TRUE),
    production_pct = (national_production / total_production) * 100
  ) %>%
  mutate(
    agriculture_grouped = ifelse(production_pct < 1, "Other Agriculture", agriculture_product)
  ) %>%
  group_by(year, agriculture_grouped) %>%
  summarise(
    production_pct = sum(production_pct, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    agriculture_grouped = factor(agriculture_grouped,
                                 levels = c("Crude Palm Oil", "Palm Kernel",
                                            "Rice", "Natural Rubber",
                                            "Pineapple", "Other Agriculture"))
  ) %>%
  ggplot(aes(x = year, y = production_pct, color = agriculture_grouped)) +
  geom_line(size = 1) +
  geom_point(size = 2) +
  labs(
    title = "Trend of Agricultural Product Production Share (2010–2018)",
    x = "Year",
    y = "Production Share (%)",
    caption = "Source: https://ekonomi.gov.my/sites/default/files/2020-09/3.5.2.pdf",
    color = "Agriculture Product"
  ) +
  theme_minimal()

pct_share + trend_analysis +
  plot_layout()


crop_gdp_data <- crop_gdp_data %>%
  group_by(year) %>%
  summarise(
    avg_value = mean(value),
    avg_volume = mean(volume),
    .groups = "drop"
  )

# Visualize Palm Oil and Palm Kernel Value and Volume from 2016 - 2020
crop_gdp_data %>%
  pivot_longer(
    cols = -year,
    names_to = "variable",
    values_to = "value"
  ) %>%
  ggplot(aes(x = year, y = value, color = variable)) +
  geom_line() +
  facet_wrap(~ variable, scales = "free_y") +
  labs(
    x = "Year",
    y = "Value & Volume",
    title = "Palm Oil and Palm Kernel Value and Volume from 2016 - 2020"
  ) +
  theme_minimal()

combined_floods <- combined_floods %>% 
  select(!state) %>%
  group_by(year) %>%
  summarise(
    total_flood = mean(flood_incidents),
    avg_rainfall = mean(avg_rainfall_mm),
    flood_duration = mean(flood_duration_hrs),
    flood_depth = mean(flood_depth_m),
    .groups = "drop"
  )

# Visualize Annual Trends of Flood-Related Variables from 2016 to 2020
combined_floods %>%
  pivot_longer(
    cols = -year, 
    names_to = "variable",
    values_to = "value"
  ) %>%
  ggplot(aes(x = year, y = value, color = variable)) +
  geom_line(size = 1.2) +
  geom_point(size = 2) +
  facet_wrap(~variable, scales = "free_y") +
  labs(
    title = "Annual Trends of Flood-Related Variables from 2016 to 2020",
    x = "Year",
    y = "Value",
    color = "Variable"
  ) +
  theme_minimal() +
  scale_x_continuous(breaks = 2016:2020)

cor.test(crop_gdp_data$avg_value, combined_floods$avg_rainfall, method = "spearman")
cor.test(crop_gdp_data$avg_value, combined_floods$flood_depth, method = "spearman")


# Merge data sets
merged_data <- combined_floods %>%
  inner_join(crop_gdp_data, by = "year")


# Model 1: Gaussian GLM
freq_model <- glm(flood_depth ~ avg_rainfall, 
                  data = merged_data,
                  family = gaussian(link = "identity"))

# Model 2:  Gamma GLM
sev_model <- glm(avg_value ~ flood_depth + avg_rainfall,
                 data = merged_data,
                 family = Gamma(link = "log"))

# Model summaries
summary(freq_model)
summary(sev_model)


# Monte Carlo Simulation to calculate premium price
monte_carlo_sim <- function(
    sum_insured, # coverage cap
    deductible_flood_depth, #flood depth trigger
    payout_factor, # % of loss paid per unit above deductible 
    loading_factor 
){
  set.seed(123)
  
  # simulate rainfall (normal distribution)
  mu_rain <- mean(merged_data$avg_rainfall) 
  sd_rain <- sd(merged_data$avg_rainfall)   
  sim_rain <- rnorm(10000, mu_rain, sd_rain)
  
  # simulate flood depth (using Gaussian GLM)
  sim_flood_depth <- predict(freq_model, 
                             newdata = data.frame(avg_rainfall = sim_rain),
                             type = "response") + rnorm(10000, 0, sigma(freq_model))
  
  # predict loss severity (using Gamma GLM)
  sim_loss <- predict(sev_model, 
                      newdata = data.frame(
                        flood_depth = sim_flood_depth,
                        avg_rainfall = sim_rain
                      ),
                      type = "response")
  
  # calculate payout per scenario
  payout <- ifelse(sim_flood_depth > deductible_flood_depth,
                   pmin(sum_insured, 
                        payout_factor * (sim_flood_depth - deductible_flood_depth) * mean(sim_loss)),
                   0)
  
  # pure premium = expected annual payout
  pure_premium <- mean(payout)
  
  # loadings (Additional charged to the premium)
  market_premium <- pure_premium * loading_factor
  
  # profitability metrics
  loss_ratio <- pure_premium / market_premium
  
  return(data.frame(
    pure_premium,
    market_premium,
    loss_ratio
  ))
  }

# Enter key parameters value
res = monte_carlo_sim(sum_insured = 4e6, 
                      deductible_flood_depth = 1.7,
                      payout_factor = 0.8,
                      loading_factor = 1.4)

res
as.integer(4e6)
summary(combined_floods$flood_depth)

# Sensitivity testing for different deductible thresholds
deductibles <- seq(1.2, 2.2, by = 0.1)

scenario_results <- lapply(deductibles, function(d) {
  monte_carlo_sim(sum_insured = 4e6,
                  deductible_flood_depth = d,
                  payout_factor = 0.8,
                  loading_factor = 1.4) %>%
    mutate(deductible_flood_depth = d)
}) %>% bind_rows()

# Premium vs Deductible
ggplot(scenario_results, aes(x = deductible_flood_depth)) +
  geom_line(aes(y = pure_premium, color = "Pure Premium"), size = 1.2) +
  geom_line(aes(y = market_premium, color = "Market Premium"), size = 1.2) +
  labs(
    title = "Premium vs Deductible Flood Depth",
    x = "Deductible Flood Depth (meters)",
    y = "Premium (RM)",
    color = "Premium Type"
  ) +
  theme_minimal()

scenario_results %>%
  select(!loss_ratio) %>%
  summary()
