# 初始化环境 ----------------------------------------------------------------
rm(list = ls())
# 安装必要包（若未安装）
# install.packages(c("openxlsx", "nih.joinpoint", "purrr", "dplyr"))

rm(list = ls())
library(tools)

# 设置工作目录
# 设置固定工作目录路径（根据你的需求修改此处）
target_dir <- "E:/20250414 Wenping Gong gaint mission/A"  # 注意使用正斜杠或双反斜杠
setwd(target_dir)  

library(dplyr)

# 自动获取ZIP文件并处理路径
list.files(pattern = "\\.zip$")
zip_file <- list.files(pattern = "\\.zip$")[3]          # 取当前目录第一个ZIP文件
stopifnot("未找到ZIP文件" = !is.na(zip_file))          # 确保文件存在
zip_name <- file_path_sans_ext(zip_file)               # 自动提取文件夹名称

# 0. 清理或创建文件夹（保留文件夹结构）
if (dir.exists(zip_name)) {
  # 清空文件夹内容但保留结构（关键修改）
  unlink(file.path(zip_name, "*"), recursive = TRUE, force = TRUE)
  message("已清空文件夹内容：", zip_name)
} else {
  dir.create(zip_name)
  message("已创建新文件夹：", zip_name)
}

# 1. 创建并解压到新文件夹
dir.create(zip_name)                                   # 新建空文件夹
unzip(zip_file, exdir = zip_name)                     # 解压到指定文件夹

# 2. 设置新工作目录
setwd(zip_name)                                        # 切换到解压文件夹
message("当前工作目录已设置为：", getwd())

# 3. 读取数据
csv_file <- list.files(pattern = "\\.csv$")[1]        # 取第一个CSV文件
data1 <- read.csv(csv_file)

library(stringr)
# 转换列类型并执行替换（网页2、网页5方法优化）
data <- data1 %>%
  # 将可能为因子型的列转为字符型（防止因子水平问题）
  mutate(across(c(measure, cause), as.character)) %>% 
  
  # measure_name列替换（网页1、网页4方法）
  mutate(measure = case_when(
    measure == "DALYs (Disability-Adjusted Life Years)" ~ "DALYs",
    TRUE ~ measure
  )) %>%
  
  # cause_name列替换（网页7、网页8方法优化）
  mutate(cause = str_replace(
    cause,
    pattern = fixed("Multidrug-resistant tuberculosis without extensive drug resistance"),
    replacement = "MDR-TB without XDR"
  )) %>%
  
  # cause_name列替换（网页7、网页8方法优化）
  mutate(cause = str_replace(
    cause,
    pattern = fixed("Extensively drug-resistant tuberculosis"),
    replacement = "XDR-TB"
  ))%>%
  
  # cause_name列替换（网页7、网页8方法优化）
  mutate(cause = str_replace(
    cause,
    pattern = fixed("Respiratory infections and tuberculosis"),
    replacement = "RI and TB"
  ))

data <- data %>% filter(val != 0)

data <- data %>% arrange(desc(cause))

library(openxlsx)
# 将data数据框的第1列到第7列转换为因子型
data[1:7] <- lapply(data[1:7], factor)
write.xlsx(data,"combined_data.xlsx")

summary(as.factor(data$location))

# 加载库 --------------------------------------------------------------------
library(openxlsx)
library(nih.joinpoint)
library(dplyr)
library(purrr)
library(stringr)

# 配置Joinpoint路径 ---------------------------------------------------------
options(joinpoint_path = "C:/Program Files (x86)/Joinpoint Command/jpCommand.exe") 

# 创建输出目录结构 ----------------------------------------------------------
output_dir <- "join_point"
dir.create(file.path(output_dir, "plots"), recursive = TRUE, showWarnings = FALSE)
dir.create(file.path(output_dir, "tables"), recursive = TRUE, showWarnings = FALSE)

# 核心分析函数 --------------------------------------------------------------
run_measure_analysis <- function(measure_type) {
  # 数据预处理（根据网页3、7建议优化）
  analysis_data <- read.xlsx("combined_data.xlsx") %>% 
    filter(measure == measure_type) %>%  # 关键修改：measure_type替换原measure
    mutate(
      se = (upper - lower)/(2*qnorm(0.975)),  # 网页9标准误计算方法
      year = as.numeric(year),
      value = val
    ) %>%
    group_by(cause, year) %>%
    slice(1) %>%  # 网页3去重建议
    ungroup() %>%
    arrange(cause, year) %>% 
    filter(!is.na(value) & !is.na(year))
  
  # 在核心分析函数前添加集群初始化
  library(parallel)
  cl <- makeCluster(detectCores() - 1)  # 保留1个核心给系统
  
  # Joinpoint分析（网页9、11配置）
  jp <- joinpoint(
    data = analysis_data,
    x = "year",
    y = "value",
    by = "cause",
    se = "se",
    run_opts = run_options(
      model = "ln",
      max_joinpoints = 5,
      n_cores = cl  # 保留1个核心给系统
    )
  )
  
  # 分析结束后添加
  stopCluster(cl)
  
  # 获取该测量类型下的原因数量
  num_causes <- length(unique(analysis_data$cause))
  
  # 动态计算绘图高度（假设每个原因需要3英寸高度，最小保证8英寸）
  plot_height <- max(8, num_causes * 3) 
  
  # 可视化输出（网页5、7建议）
  plot_name <- paste0(gsub("[^[:alnum:]]", "_", measure_type), "_trend.pdf")
  pdf(file.path(output_dir, "plots", plot_name), 
      width = 10, height = plot_height, family = "GB1")
  print(jp_plot(jp, ncol = 1#, 
                #title = paste("Joinpoint Analysis for", measure_type)
  )
  )
  dev.off()
  
  # 结果处理：正确添加measure_type列
  list(
    apc = jp$apc %>% mutate(measure = measure_type),  # 修改为赋值操作
    report = jp$report %>% mutate(measure = measure_type),
    metadata = tibble(
      measure = measure_type,  # 关键修改
      Analysis_Date = Sys.Date(),
      Data_Points = nrow(analysis_data),
      Joinpoints = max(jp$report$segment)
    )
  )
}

# 主执行流程 ----------------------------------------------------------------
# 获取所有测量类型（网页1、4数据读取方法）
measures <- unique(read.xlsx("combined_data.xlsx")$measure)

# 批量执行分析（网页7建议的purrr应用）
analysis_results <- map(measures, safely(run_measure_analysis))

# 结果整合与输出 ------------------------------------------------------------
# Excel多分页输出（网页2、8建议）
wb <- createWorkbook()
walk2(analysis_results, measures, ~{
  if (!is.null(.x$result)) {
    addWorksheet(wb, str_sub(paste0(.y, "_APC"), 1, 31))  # Excel表名长度限制
    writeDataTable(wb, sheet = str_sub(paste0(.y, "_APC"), 1, 31), 
                   x = .x$result$apc, startRow = 2,
                   tableStyle = "TableStyleMedium2")
    
    addWorksheet(wb, str_sub(paste0(.y, "_Report"), 1, 31))
    writeDataTable(wb, sheet = str_sub(paste0(.y, "_Report"), 1, 31),
                   x = .x$result$report, startRow = 2,
                   tableStyle = "TableStyleMedium2")
  }
})
saveWorkbook(wb, file.path(output_dir, "tables", "Joinpoint_Results.xlsx"), overwrite = TRUE)

# 元数据记录（网页3、6建议）
metadata <- map_dfr(analysis_results, ~.x$result$metadata)
write.xlsx(metadata, file.path(output_dir, "tables", "Analysis_Metadata.xlsx"))

# 错误日志记录（如果有）
errors <- compact(map(analysis_results, "error"))
if (length(errors) > 0) {
  saveRDS(errors, file.path(output_dir, "tables", "Error_Logs.rds"))
}

save.image("result.RData")

###  ---------------------------------------------------------------------------
###                                   可视化
###  ---------------------------------------------------------------------------

library(openxlsx)
rm(list = ls())

load("result.RData")

str(analysis_results)

library(dplyr)

# 提取所有apc子集并合并
combined_apc <- bind_rows(
  analysis_results[[1]]$result$apc,
  analysis_results[[2]]$result$apc,
  analysis_results[[3]]$result$apc,
  analysis_results[[4]]$result$apc
) %>% 
  select(cause, measure, segment_start, segment_end, 
         apc, apc_95_lcl, apc_95_ucl, p_value)

###整合热图、形状与文字注释的创新可视化方案
combined_apc <- combined_apc %>% 
  mutate(
    time_interval = paste(segment_start, segment_end, sep = "-"),  # 时间段标签
    significance = case_when(
      p_value < 0.001 ~ "***",
      p_value < 0.01 ~ "**",
      p_value < 0.05 ~ "*",
      TRUE ~ ""
    ),            # 显著性标记
    trend = ifelse(apc > 0, "Increase", "Decrease")              # 变化方向
  )

str(combined_apc)

library(ggplot2)
library(dplyr)

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
  group_by(measure, cause) %>%
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

# 修正后的可视化代码
p <- ggplot() +
  facet_wrap(~ measure, ncol = 2,
             labeller = labeller(measure = label_value),
             strip.position = "top") +
  
  # 趋势色块（分两层绘制）
  geom_segment(
    data = combined_apc,
    aes(x = segment_start, xend = segment_end, 
        y = cause, yend = cause,
        color = trend),
    color = "black", linewidth = 3.5, lineend = "butt", alpha = 0.2
  ) +
  geom_segment(
    data = combined_apc,
    aes(x = segment_start, xend = segment_end,
        y = cause, yend = cause,
        color = trend),
    linewidth = 3, alpha = 0.8, lineend = "butt"
  ) +
  
  # 时间分割点标签（独立数据层）
  geom_text(
    data = label_points,
    aes(x = time_points, y = cause, 
        label = label, hjust = hjust),
    color = "black", size = 3, vjust = -0.8,
    fontface = "bold", check_overlap = TRUE
  ) +
  
  # 显著性标记
  geom_text(
    data = combined_apc,
    aes(x = (segment_start + segment_end)/2, 
        y = cause, label = significance),
    color = "white", size = 4, 
    fontface = "bold", vjust = 0.6
  ) +
  
  # 视觉样式设置
  scale_color_manual(
    values = c("Increase" = "#C44E52", "Decrease" = "#4C72B0"),
    guide = guide_legend(title = "Trend Direction")
  ) +
  scale_x_continuous(
    breaks = seq(1980, 2020, by = 5),
    limits = c(min(label_points$time_points)-0.5, 
               max(label_points$time_points)+0.5),
    expand = expansion(add = 0.5)
  ) +
  theme_minimal(base_size = 11) +
  theme(
    legend.position = "top",
    axis.text.x = element_text(angle = 30, hjust = 1, color = "grey30"),
    strip.background = element_rect(fill = "#F8F9FA", color = "grey80"),
    panel.grid.major.x = element_line(color = "grey90", linewidth = 0.3),
    panel.spacing = unit(1.2, "lines")
  ) +
  labs(
    x = "Time Segment (Year)", 
    y = "Digestive Disease Category",
    title = "Annual Percent Change Trends with Joinpoint Segmentation"
  )

# 输出图像
ggsave("APC_trend_final.pdf", p, 
       width = 12, height = 8, units = "in",
       device = "pdf", limitsize = FALSE)
ggsave("APC_trend_final.jpg", p, 
       width = 12, height = 8, units = "in", 
       dpi = 300)
