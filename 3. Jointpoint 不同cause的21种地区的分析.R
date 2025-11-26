# 初始化环境 ----------------------------------------------------------------
rm(list = ls())
library(tools)
library(openxlsx)
library(nih.joinpoint)
library(dplyr)
library(purrr)
library(stringr)
library(ggplot2)

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

colnames(data_all)[2] <- c("measure")

colnames(data_all)[4] <- c("location")

colnames(data_all)[6] <- c("sex")

colnames(data_all)[8] <- c("age")

colnames(data_all)[10] <- c("cause")

colnames(data_all)[12] <- c("metric")
      

# 获取所有cause（原location改为cause）
causes <- unique(data_all$cause)

# 主循环处理每个cause（添加数据验证） ----------------------------
walk(causes, function(current_cause) {
  # 创建cause专用目录
  cause_dir <- file.path(getwd(), gsub("[^[:alnum:]]", "_", current_cause))
  dir.create(cause_dir, showWarnings = FALSE, recursive = TRUE)
  
  # 筛选并处理数据（添加数据验证）
  data_processed <- data_all %>%
    filter(cause == current_cause) %>%
    mutate(across(c(measure, location), as.character)) %>%
    mutate(
      measure = case_when(
        measure == "DALYs (Disability-Adjusted Life Years)" ~ "DALYs",
        TRUE ~ measure
      )
    ) %>%
    filter(val != 0) %>%
    arrange(desc(location))
  
  # 验证数据是否有效（新增数据验证）
  if (nrow(data_processed) == 0) {
    message(sprintf("跳过原因 [%s]: 无有效数据", current_cause))
    return()
  }
  
  # 保存处理后的数据
  write.xlsx(data_processed, file.path(cause_dir, "combined_data.xlsx"))
  
  # 设置Joinpoint输出目录
  output_dir <- file.path(cause_dir, "join_point")
  dir.create(file.path(output_dir, "plots"), recursive = TRUE, showWarnings = FALSE)
  dir.create(file.path(output_dir, "tables"), recursive = TRUE, showWarnings = FALSE)
  
  # 核心分析函数（添加多重验证） --------------------------------
  run_measure_analysis <- function(measure_type) {
    # 读取并预处理数据
    analysis_data <- tryCatch({
      df <- read.xlsx(file.path(cause_dir, "combined_data.xlsx")) %>% 
        filter(measure == measure_type) %>%
        mutate(
          se = (upper - lower)/(2*qnorm(0.975)),
          year = as.numeric(year),
          value = val
        ) %>%
        group_by(location, year) %>%
        slice(1) %>%
        ungroup() %>%
        arrange(location, year) %>%
        filter(!is.na(value) & !is.na(year))
      
      # 验证每个location的数据量（新增分组验证）
      valid_locations <- df %>%
        group_by(location) %>%
        summarise(n = n(), .groups = "drop") %>%
        filter(n >= 6)  # Joinpoint至少需要4个数据点
      
      df %>% filter(location %in% valid_locations$location)
    }, error = function(e) {
      message(sprintf("数据加载失败 [%s-%s]: %s", current_cause, measure_type, e$message))
      return(NULL)
    })
    
    # 检查数据是否有效
    if (is.null(analysis_data) || nrow(analysis_data) == 0) {
      message(sprintf("跳过分析 [%s-%s]: 无有效数据", current_cause, measure_type))
      return(NULL)
    }
    
    # 并行计算配置
    cl <- parallel::makeCluster(parallel::detectCores() - 1)
    on.exit(parallel::stopCluster(cl))
    
    # 执行Joinpoint分析（添加错误处理）
    jp_result <- tryCatch({
      joinpoint(
        data = analysis_data,
        x = "year",
        y = "value",
        by = "location",
        se = "se",
        run_opts = run_options(model = "ln", max_joinpoints = 5)
      )
    }, error = function(e) {
      message(sprintf("Joinpoint分析失败 [%s-%s]: %s", current_cause, measure_type, e$message))
      return(NULL)
    })
    
    if (is.null(jp_result)) return(NULL)
    
    # 可视化输出（添加空数据保护）
    if (length(unique(analysis_data$location)) > 0) {
      plot_name <- paste0(gsub("[^[:alnum:]]", "_", measure_type), "_trend.pdf")
      pdf(file.path(output_dir, "plots", plot_name), 
          width = 10, height = max(8, length(unique(analysis_data$location)) * 3))
      print(jp_plot(jp_result, ncol = 1))
      dev.off()
    }
    
    # 结果处理
    list(
      apc = jp_result$apc %>% mutate(measure = measure_type),
      report = jp_result$report %>% mutate(measure = measure_type),
      metadata = tibble(
        measure = measure_type,
        Analysis_Date = Sys.Date(),
        Data_Points = nrow(analysis_data),
        Joinpoints = max(jp_result$report$segment)
      )
    )
  }
  
  # 执行分析（添加错误日志记录）
  measures <- unique(data_processed$measure)
  analysis_results <- map(measures, function(m) {
    result <- safely(run_measure_analysis)(m)
    if (!is.null(result$error)) {
      message(sprintf("分析错误 [%s-%s]: %s", current_cause, m, result$error$message))
    }
    result
  })
  
  # 保存结果（跳过空结果）
  wb <- createWorkbook()
  walk2(analysis_results, measures, ~{
    if (!is.null(.x$result)) {
      sheet_name_apc <- str_sub(paste0(.y, "_APC"), 1, 31)
      sheet_name_report <- str_sub(paste0(.y, "_Report"), 1, 31)
      
      if (!is.null(.x$result$apc)) {
        addWorksheet(wb, sheet_name_apc)
        writeDataTable(wb, sheet = sheet_name_apc, x = .x$result$apc, startRow = 2)
      }
      
      if (!is.null(.x$result$report)) {
        addWorksheet(wb, sheet_name_report)
        writeDataTable(wb, sheet = sheet_name_report, x = .x$result$report, startRow = 2)
      }
    }
  })
  
  if (length(names(wb)) > 0) {
    saveWorkbook(wb, file.path(output_dir, "tables", "Joinpoint_Results.xlsx"), overwrite = TRUE)
  } else {
    message(sprintf("无有效结果可保存 [%s]", current_cause))
  }
  
  # 保存元数据（添加空值保护）
  metadata <- map_dfr(analysis_results, ~ if (!is.null(.x$result)) .x$result$metadata else NULL)
  
  if (nrow(metadata) > 0) {
    write.xlsx(metadata, file.path(output_dir, "tables", "Analysis_Metadata.xlsx"))
  }
} )


# 可视化部分（添加空数据过滤） ----------------------------------------
setwd(file.path(target_dir, zip_name))

combined_apc <- map_dfr(causes, ~{
  cause_dir <- file.path(getwd(), gsub("[^[:alnum:]]", "_", .x))
  result_file <- file.path(cause_dir, "join_point/tables/Joinpoint_Results.xlsx")
  
  if (!file.exists(result_file)) return(NULL)
  
  # 获取所有工作表名称
  sheets <- openxlsx::getSheetNames(result_file)
  apc_sheets <- sheets[str_detect(sheets, "_APC$")]  # 正则匹配APC结尾的工作表
  
  if (length(apc_sheets) == 0) return(NULL)
  
  # 读取所有APC工作表并合并
  map_dfr(apc_sheets, function(sht) {
    tryCatch({
      read.xlsx(result_file, sheet = sht) %>%
        mutate(
          cause = .x,
          measure = str_remove(sht, "_APC")  # 从工作表名提取measure类型
        )
    }, error = function(e) {
      message(sprintf("读取失败 [%s-%s]: %s", .x, sht, e$message))
      NULL
    })
  }) 
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
  ) %>%
  filter(!is.na(apc))  # 最终过滤无效记录

###通过形状编码测量指标(measure)实现紧凑布局

# 创建位置偏移映射（假设有4个measure）
# 修正后的偏移映射（示例）
measure_offset <- c(
  "Deaths" = -0.3, 
  "DALYs" = -0.1,
  "Prevalence" = 0.1, 
  "Incidence" = 0.3
)

combined_apc <- combined_apc %>% 
  mutate(
    x_mid = (segment_start + segment_end)/2 + measure_offset[measure]
  )

# 为每个cause创建measure子层级
combined_apc <- combined_apc %>%
  group_by(cause) %>%
  mutate(
    y_pos = as.numeric(factor(cause)) + (as.numeric(factor(measure)) - 2.5)/10  # 在y轴上为每个measure创建微调位置
  ) %>%
  ungroup()

# 检查measure值的实际类型
unique(combined_apc$measure)  # 返回结果应与measure_offset的键名一致

# 检查x_mid是否存在NA值
sum(is.na(combined_apc$x_mid))  # 若返回381则验证假设成立

str(combined_apc)

# 加载必要包
library(tidyr)

# 生成唯一时间分割点标签数据（最终修正版）
label_points <- combined_apc %>%
  group_by(measure, cause, location) %>%
  reframe(
    time_points = unique(c(segment_start, segment_end))  # 确保时间点去重
  ) %>%
  group_by(measure, cause) %>%
  mutate(
    hjust = case_when(
      time_points == min(time_points) ~ 0,   # 首时间点左对齐
      time_points == max(time_points) ~ 1,   # 末时间点右对齐
      TRUE ~ 0.5                             # 中间节点居中对齐
    ),
    label = as.character(round(time_points)) # 确保整数年份
  ) %>%
  ungroup()

str(label_points)

# 生成分cause的可视化 --------------------------------------------------------
# 生成分cause的可视化 --------------------------------------------------------
walk(causes, function(current_cause) {
  # 创建安全目录名（限制长度并替换特殊字符）
  safe_cause_name <- substr(gsub("[^[:alnum:]]", "_", current_cause), 1, 80)
  cause_dir <- file.path(getwd(), safe_cause_name)
  vis_dir <- file.path(cause_dir, "vis")
  dir.create(vis_dir, showWarnings = FALSE, recursive = TRUE)
  
  # 数据预处理 ------------------------------------------------------------
  cause_data <- combined_apc %>%
    filter(cause == current_cause) %>%
    mutate(
      segment_start = as.numeric(segment_start),
      segment_end = as.numeric(segment_end),
      significance = str_remove_all(significance, "\\u200b"),  # 移除零宽空格
      time_interval = paste(segment_start, segment_end, sep = "-"),
      location = reorder(location, desc(location))
    ) %>%
    distinct()
  
  label_points <- cause_data %>%
    group_by(location, measure) %>%
    summarise(
      time_points = c(segment_start, segment_end),
      label = as.character(c(segment_start, segment_end)),
      .groups = 'drop'
    ) %>%
    mutate(
      hjust = ifelse(label == first(label), 0, 0.5)
    )
  
  # 跳过无数据的情况
  if (nrow(cause_data) == 0) {
    message(sprintf("跳过可视化 [%s]: 无有效数据", current_cause))
    return()
  }
  
  # 动态计算图表尺寸
  n_measures <- n_distinct(cause_data$measure)
  base_height <- 6 + ceiling(n_distinct(cause_data$location) * 0.6)
  plot_width <- 10 + n_measures * 2
  
  # 可视化核心代码 --------------------------------------------------------
  p <- ggplot(cause_data) +
    # 趋势色块（双层次显示）
    geom_segment(
      aes(x = segment_start, xend = segment_end,
          y = location, yend = location,
          color = trend),
      linewidth = 4, lineend = "butt", alpha = 0.3
    ) +
    geom_segment(
      aes(x = segment_start, xend = segment_end,
          y = location, yend = location,
          color = trend),
      linewidth = 3, lineend = "butt"
    ) +
    # 时间分割点标签
    geom_text(
      data = label_points,
      aes(x = time_points, y = location, 
          label = label, hjust = hjust),
      color = "black", size = 3, vjust = -0.8,
      fontface = "bold", check_overlap = TRUE
    ) +
    # 显著性标记
    geom_text(
      aes(x = (segment_start + segment_end)/2, 
          y = location, label = significance),
      color = "white", size = 4, 
      fontface = "bold", vjust = 0.6
    ) +
    # 分面设置
    facet_wrap(
      ~ measure, 
      ncol = ifelse(n_measures > 2, 2, 1),
      labeller = labeller(measure = label_wrap_gen(15))
    ) +
    # 坐标轴设置
    scale_x_continuous(
      name = "Time Segment",
      breaks = seq(min(cause_data$segment_start), max(cause_data$segment_end), by = 5),
      limits = c(min(cause_data$segment_start)-1, max(cause_data$segment_end)+1),
      expand = expansion(add = 0.5)
    ) +
    scale_color_manual(
      values = c("Increase" = "#C44E52", "Decrease" = "#4C72B0"),
      guide = guide_legend(title = "Trend Direction")
    ) +
    labs(
      y = "Location",
      title = paste("APC Trends:", current_cause),
      subtitle = "Segmented by Measure Type"
    ) +
    theme_minimal(base_size = 12) +
    theme(
      legend.position = "top",
      axis.text.x = element_text(angle = 30, hjust = 1),
      strip.background = element_rect(fill = "#F7F7F7", color = NA),
      panel.spacing = unit(1.5, "lines"),
      plot.title = element_text(face = "bold", size = rel(1.2))
    )
  
  # 安全保存文件 ----------------------------------------------------------
  tryCatch({
    ggsave(
      file.path(vis_dir, paste0(safe_cause_name, "_APC.pdf")),
      plot = p,
      width = plot_width,
      height = base_height,
      limitsize = FALSE
    )
    ggsave(
      file.path(vis_dir, paste0(safe_cause_name, "_APC.png")),
      plot = p,
      width = plot_width,
      height = base_height,
      dpi = 300,
      limitsize = FALSE
    )
  }, error = function(e) {
    message(sprintf("保存失败 [%s]: %s", current_cause, e$message))
  })
})
