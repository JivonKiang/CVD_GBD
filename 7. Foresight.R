rm(list = ls())

####  前期数据准备-------------------------------------------------------------
library(openxlsx)
library(dplyr)
library(tidyverse)
library(ggsci)
library(ggrepel)
library(ggpubr)
library(tools)

# 设置工作目录
# 设置固定工作目录路径（根据你的需求修改此处）
target_dir <- "E:/20250414 Wenping Gong gaint mission/E"  # 注意使用正斜杠或双反斜杠
setwd(target_dir)  

# 定义主文件夹路径
main_folder <- "6. Foresight"

# 列出所有.csv文件的路径
csv_files <- list.files(main_folder, pattern = "\\.csv$", full.names = TRUE, recursive = TRUE)

# 读取所有.csv文件
all_data <- lapply(csv_files, read.csv)

# 如果需要将所有数据合并为一个数据框
# 假设所有文件的列结构相同
combined_data <- do.call(rbind, all_data)

# 查看结果
print(combined_data)

combined_data <- na.omit(combined_data)
combined_data <- combined_data[combined_data$Value != 0, ]

colnames(combined_data)

library(dplyr)
library(tidyr)
library(openxlsx)

# 前期数据准备及处理步骤保持不变...

# 数据预处理和格式化
processed_data <- combined_data %>%
  filter(Scenario != "Past") %>%
  mutate(
    across(c(Value, Lower.bound, Upper.bound), ~ sprintf("%.4f", .x)),
    Value_and_UI = paste0(Value, " [", Lower.bound, ", ", Upper.bound, "]")
  ) %>%
  select(Cause.of.death.or.injury, Measure, Year, Scenario, Value_and_UI)

colnames(processed_data)[1] <- c("cause")

str(processed_data)

library(tidyr)

# 长格式转宽格式
processed_wide <- processed_data %>% 
  pivot_wider(
    names_from = Scenario,          # 将Scenario的值作为新列名
    values_from = Value_and_UI     # 填充Value_and_UI的值
  )

library(openxlsx)
library(dplyr)

# 创建新工作簿
wb <- createWorkbook()

# 按疾病原因分组处理
processed_wide %>% 
  # 按cause分组并排序年份
  group_by(cause) %>% 
  #arrange(Year, .by_group = TRUE) %>% 
  # 拆分成子数据集列表
  group_split() %>% 
  # 遍历每个子数据集
  walk(~{
    # 获取当前疾病名称
    current_cause <- unique(.x$cause)
    
    # 创建合法工作表名称（处理特殊字符和长度）
    sheet_name <- substr(gsub("[[:punct:]]", "", current_cause), 1, 31)
    
    # 添加工作表
    addWorksheet(wb, sheetName = sheet_name)
    
    # 写入数据（移除分组信息）
    writeData(wb, 
              sheet = sheet_name, 
              x = as.data.frame(.x) %>% select(-cause),
              startRow = 1,
              headerStyle = createStyle(textDecoration = "bold"))
    
    # 设置自动列宽
    setColWidths(wb, sheet_name, cols = 1:ncol(.x), widths = "auto")
  })


# 保存工作簿
saveWorkbook(wb, file.path(main_folder, "Formatted_Output.xlsx"), overwrite = TRUE)

###    可视化   ----------------------------------------------------------------
library(tidyverse)
library(patchwork)
library(RColorBrewer)  # 新增颜色包

# 定义颜色方案
scenario_colors <- rev(brewer.pal(5, "RdYlBu"))

# 预处理数据
processed_plot <- processed_data %>%
  mutate(
    Value = as.numeric(str_extract(Value_and_UI, "^\\d+\\.?\\d*")),
    Lower = as.numeric(str_extract(Value_and_UI, "(?<=\\[)\\d+\\.?\\d*")),
    Upper = as.numeric(str_extract(Value_and_UI, "(?<=, )\\d+\\.?\\d*")),
    Year = as.numeric(Year)
  ) %>%
  filter(!is.na(Value))

# 创建单个Measure的绘图函数
create_measure_plot <- function(measure_name) {
  processed_plot %>%
    filter(Measure == measure_name) %>%
    ggplot(aes(x = Year, y = Value, color = Scenario)) +
    geom_line(linewidth = 0.8) +
    geom_point(size = 1.2) +
    geom_ribbon(aes(ymin = Lower, ymax = Upper, fill = Scenario), 
                alpha = 0.15, color = NA) +
    facet_wrap(~ cause, scales = "free_y", ncol = 1) + ####分为几行
    labs(title = paste("Measure:", measure_name),
         x = "Year", y = "Rate per 100,000") +
    scale_color_manual(values = scenario_colors) +  # 修改颜色设置
    scale_fill_manual(values = scenario_colors) +   # 修改填充色
    scale_x_continuous(breaks = seq(2020, 2050, 5)) +
    theme_bw() +
    theme(
      axis.text.x = element_text(angle = 45, hjust = 1),
      legend.position = "bottom",
      strip.text = element_text(size = 8, face = "bold"),
      plot.margin = unit(c(1,1,1,1), "cm")
    ) +
    guides(
      color = guide_legend(nrow = 1),  # 颜色图注3行
      fill = guide_legend(nrow = 1)    # 填充图注3行
    )
}

# 生成并保存图表
measure_plots <- unique(processed_plot$Measure) %>% 
  set_names() %>% 
  map(create_measure_plot)


# 保存结果
walk2(measure_plots, names(measure_plots),
      ~ ggsave(file.path(main_folder, 
                         paste0("Measure_", gsub("\\s+", "_", .y), ".png")),
               plot = .x,
               width = 10,
               height = 6,
               dpi = 300,
               limitsize = FALSE))  # 新增关键参数

# 修改后的组合图定义
combined_plot <- wrap_plots(measure_plots, ncol = 1, guides = "collect") + 
  plot_layout(guides = 'collect') &  # 确保收集所有图例
  theme(
    legend.position = "bottom",     # 强制图例在底部
    legend.direction = "horizontal", # 水平排列图例项
    legend.box = "vertical",         # 垂直堆叠图例框
    legend.text = element_text(size = 9),
    legend.title = element_blank()   # 可选隐藏图例标题
  )

# 保存组合图（调整高度计算）
ggsave(file.path(main_folder, "Combined_Measures.png"),
       plot = combined_plot,
       width = 10,
       height = 15,  # +3英寸给图例预留空间
       dpi = 300,
       limitsize = FALSE)
