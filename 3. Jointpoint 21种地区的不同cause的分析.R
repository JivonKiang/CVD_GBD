# 初始化环境 ----------------------------------------------------------------
rm(list = ls())
# 安装必要包（若未安装）
# install.packages(c("openxlsx", "nih.joinpoint", "purrr", "dplyr", "stringr"))

library(tools)
library(openxlsx)
library(nih.joinpoint)
library(dplyr)
library(purrr)
library(stringr)

# 设置工作目录
target_dir <- "E:/20250414 Wenping Gong gaint mission/E"
setwd(target_dir)  

# 自动获取ZIP文件并处理路径
zip_file <- list.files(pattern = "\\.zip$")[1]
stopifnot("未找到ZIP文件" = !is.na(zip_file))
zip_name <- tools::file_path_sans_ext(zip_file)

# 创建解压目录
if (dir.exists(zip_name)) {
  unlink(file.path(zip_name, "*"), recursive = TRUE, force = TRUE)
} else {
  dir.create(zip_name)
}

# 解压文件
unzip(zip_file, exdir = zip_name)
setwd(zip_name)

# 读取原始数据
csv_file <- list.files(pattern = "\\.csv$")[1]
data_all <- read.csv(csv_file)

# 获取所有location
locations <- unique(data_all$location)

# 主循环处理每个location
walk(locations, function(loc) {
  # 创建location专用目录
  loc_dir <- file.path(getwd(), gsub("[^[:alnum:]]", "_", loc))
  dir.create(loc_dir, showWarnings = FALSE, recursive = TRUE)
  
  # 筛选并处理数据
  data_processed <- data_all %>%
    filter(location == loc) %>%
    mutate(across(c(measure, cause), as.character)) %>%
    mutate(
      measure = case_when(
        measure == "DALYs (Disability-Adjusted Life Years)" ~ "DALYs",
        TRUE ~ measure
      )#,
      #cause = str_replace_all(cause, c(
      #  "Multidrug-resistant tuberculosis without extensive drug resistance" = "MDR-TB without XDR",
      #  "Extensively drug-resistant tuberculosis" = "XDR-TB",
      #  "Respiratory infections and tuberculosis" = "RI and TB"
      #)
  #)
    ) %>%
    filter(val != 0) %>%
    arrange(desc(cause))
  
  # 保存处理后的数据
  write.xlsx(data_processed, file.path(loc_dir, "combined_data.xlsx"))
  
  # 设置Joinpoint输出目录
  output_dir <- file.path(loc_dir, "join_point")
  dir.create(file.path(output_dir, "plots"), recursive = TRUE)
  dir.create(file.path(output_dir, "tables"), recursive = TRUE)
  
  # 核心分析函数（修改版）
  run_measure_analysis <- function(measure_type) {
    analysis_data <- read.xlsx(file.path(loc_dir, "combined_data.xlsx")) %>% 
      filter(measure == measure_type) %>%
      mutate(
        se = (upper - lower)/(2*qnorm(0.975)),
        year = as.numeric(year),
        value = val
      ) %>%
      group_by(cause, year) %>%
      slice(1) %>%
      ungroup() %>%
      arrange(cause, year) %>% 
      filter(!is.na(value) & !is.na(year))
    
    cl <- parallel::makeCluster(parallel::detectCores() - 1)
    jp <- joinpoint(
      data = analysis_data,
      x = "year",
      y = "value",
      by = "cause",
      se = "se",
      run_opts = run_options(model = "ln", max_joinpoints = 5)
    )
    parallel::stopCluster(cl)
    
    # 可视化输出
    plot_name <- paste0(gsub("[^[:alnum:]]", "_", measure_type), "_trend.pdf")
    pdf(file.path(output_dir, "plots", plot_name), 
        width = 10, height = max(8, length(unique(analysis_data$cause)) * 3))
    print(jp_plot(jp, ncol = 1))
    dev.off()
    
    # 结果处理
    list(
      apc = jp$apc %>% mutate(measure = measure_type),
      report = jp$report %>% mutate(measure = measure_type),
      metadata = tibble(
        measure = measure_type,
        Analysis_Date = Sys.Date(),
        Data_Points = nrow(analysis_data),
        Joinpoints = max(jp$report$segment)
      )
    )
  }
  
  # 执行分析
  measures <- unique(data_processed$measure)
  analysis_results <- map(measures, safely(run_measure_analysis))
  
  # 保存结果
  wb <- createWorkbook()
  walk2(analysis_results, measures, ~{
    if (!is.null(.x$result)) {
      addWorksheet(wb, str_sub(paste0(.y, "_APC"), 1, 31))
      writeDataTable(wb, sheet = str_sub(paste0(.y, "_APC"), 1, 31), 
                     x = .x$result$apc, startRow = 2)
      
      addWorksheet(wb, str_sub(paste0(.y, "_Report"), 1, 31))
      writeDataTable(wb, sheet = str_sub(paste0(.y, "_Report"), 1, 31),
                     x = .x$result$report, startRow = 2)
    }
  })
  saveWorkbook(wb, file.path(output_dir, "tables", "Joinpoint_Results.xlsx"), overwrite = TRUE)
  
  # 保存元数据
  metadata <- map_dfr(analysis_results, ~.x$result$metadata)
  write.xlsx(metadata, file.path(output_dir, "tables", "Analysis_Metadata.xlsx"))
})

# 可视化部分（跨location整合）
setwd(file.path(target_dir, zip_name))

combined_apc <- map_dfr(locations, ~{
  loc_dir <- file.path(getwd(), gsub("[^[:alnum:]]", "_", .x))
  read.xlsx(file.path(loc_dir, "join_point/tables/Joinpoint_Results.xlsx"), sheet = 1) %>%
    mutate(location = .x)
}) %>%
  mutate(
    time_interval = paste(segment_start, segment_end, sep = "-"),
    significance = case_when(
      p_value < 0.001 ~ "​**​*",
      p_value < 0.01 ~ "​**​",
      p_value < 0.05 ~ "*",
      TRUE ~ ""
    ),
    trend = ifelse(apc > 0, "Increase", "Decrease")
  )

# 生成最终可视化
p <- ggplot(combined_apc, aes(x = time_interval, y = cause, fill = trend)) +
  geom_tile() +
  facet_grid(location ~ measure, scales = "free") +
  scale_fill_manual(values = c("Increase" = "#C44E52", "Decrease" = "#4C72B0")) +
  theme_minimal() +
  labs(x = "Time Segment", y = "Cause", title = "APC Trends by Location")

ggsave("Combined_APC_Trends.pdf", p, width = 16, height = 12)
ggsave("Combined_APC_Trends.png", p, width = 16, height = 12, dpi = 300)