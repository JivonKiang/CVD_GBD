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
target_dir <- "E:/20250414 Wenping Gong gaint mission/A"  # 注意使用正斜杠或双反斜杠
setwd(target_dir)  

library(dplyr)

# 自动获取ZIP文件并处理路径
list.files(pattern = "\\.zip$")
zip_file <- list.files(pattern = "\\.zip$")[5]          # 取当前目录第一个ZIP文件
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

big_hat <- c("RI and TB")

df <- data
###   数据读取完毕--------------------------------------------------------------

# 预处理：确保年龄因子顺序
age_levels <- c(
  "<1 year", "12-23 months", "2-4 years", "5-9 years", 
  "10-14 years", "15-19 years", "20-24 years", "25-29 years",
  "30-34 years", "35-39 years", "40-44 years", "45-49 years",
  "50-54 years", "55-59 years", "60-64 years", "65-69 years",
  "70-74 years", "75-79 years", "80-84 years", "85-89 years",
  "90-94 years", "95+ years"
)

# 重新定义因子顺序
df$age <- factor(df$age, levels = age_levels)

# 获取当前工作目录
current_dir <- getwd()
print(paste("当前工作目录为:", current_dir))

# 创建可视化输出目录
output_dir <- file.path(current_dir, "visualization")
if (!dir.exists(output_dir)) {
  dir.create(output_dir)
}

# 定义图形参数
plot_theme <- theme_minimal(base_size = 12) +
  theme(
    panel.grid.minor = element_blank(),
    legend.position = "top",
    strip.text = element_text(size = 8),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

# 获取所有measure类型
measures <- levels(factor(df$measure))  # 包含"DALYs", "Deaths", "Incidence", "Prevalence"

# 更新图形主题设置（在所有可视化代码段之前）
plot_theme <- theme_minimal(base_size = 20) +  # 基础字号从12增大到16
  theme(
    panel.grid.minor = element_blank(),
    legend.position = "top",
    legend.text = element_text(size = 20),       # 图例文字
    legend.title = element_text(size = 20),      # 图例标题
    strip.text = element_text(size = 20),         # 分面标题保持原大小
    axis.title = element_text(size = 20),         # 坐标轴标题
    axis.text.y = element_text(size = 20),        # y轴标签
    axis.text.x = element_text(                   # x轴标签保持原大小
      angle = 45, 
      hjust = 1,
      size = 14  # 保持原始字号
    ),
    plot.title = element_text(size = 20, face = "bold")  # 主标题
  )

#### 第1个可视化：按年龄组分面，展示不同性别在各疾病类型的指标值分布-----------
for (measure_type in measures) {
  plot_data <- df %>% 
    filter(measure == measure_type,
           year == max(df$year),#筛选最大年份
           !is.na(age))  # 确保年龄非空
  
  p <- ggplot(plot_data, aes(x = cause, y = val, color = sex, group = sex)) +  # x轴改为疾病类型
    geom_line(linewidth = 0.8) + 
    geom_point(size = 1.5) +
    scale_color_jama() +
    labs(
      title = paste0(measure_type, " by Disease Type and Sex (", max(df$year), ")"),  # 标题优化
      x = "Disease Type",  # 轴标签更新
      y = "Rate per 100,000",
      color = "Sex"
    ) +
    plot_theme +
    theme(
      axis.text.x = element_text(angle = 20, hjust = 1, vjust = 1),  # 倾斜疾病名称
      strip.text = element_text(size = 20, face = "bold")  # 强化分面标题
    ) +  
    facet_wrap(~ age,  # 分面变量改为年龄
               ncol = 5,  # 动态列数
               labeller = labeller(age = label_both),
               scales = "free_y")  # 显示年龄组标识
  
  # 设置输出文件路径
  filename <- gsub(" ", "_", measure_type)
  
  # 保存PDF
  pdf_file <- file.path(output_dir, paste0(filename, " facet age - x cause.pdf"))
  ggsave(pdf_file, p, 
         width = 25, height = 20,  
         units = "in", dpi = 600)
  
  # 保存JPEG
  jpeg_file <- file.path(output_dir, paste0(filename, " facet age - x cause.jpg"))
  ggsave(jpeg_file, p, 
         width = 25, height = 20, 
         units = "in", dpi = 600)
  
  # 打印进度
  cat("已生成:", measure_type, "的可视化文件\n")
}


#### 第2个可视化：按年龄组分面，展示不同性别在各年份的指标值趋势---------------
for (measure_type in measures) {
  # 过滤数据
  plot_data <- df %>% 
    filter(measure == measure_type,
           !is.na(age),# 移除缺失的年龄组
           cause == big_hat) #筛选大帽子进行可视化
  
  # 创建基础图形
  p <- ggplot(plot_data, aes(x = year, y = val, color = sex)) +
    geom_line(linewidth = 0.8) + 
    geom_point(size = 1.5) +
    scale_color_jama() +
    scale_x_continuous(breaks = unique(plot_data$year))+
    scale_fill_jama() +        # 填充颜色使用jama调色板（与线条颜色一致）
    labs(
      title = paste0(measure_type, " Trends by Year and Sex"),
      x = "Year",
      y = "Rate per 100,000",
      color = "Sex"
    ) +
    plot_theme +
    theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 1),  # 倾斜显示年龄标签
          strip.text = element_text(size = 20, face = "bold")  # 强化分面标题
    )+ 
    facet_wrap(~ age, 
               ncol = 4,  # 每行显示5个年龄组
               scales = "free_y"#, # y轴独立缩放
               #switch = "x"
    )  # 强制每个分面显示x轴
  
  # 设置输出文件路径
  filename <- gsub(" ", "_", measure_type)
  
  # 保存PDF
  pdf_file <- file.path(output_dir, paste0(filename, " facet age - x year.pdf"))
  ggsave(pdf_file, p, 
         width = 25, height = 20, 
         units = "in", dpi = 600)
  
  # 保存JPEG
  jpeg_file <- file.path(output_dir, paste0(filename, " facet age - x year.jpg"))
  ggsave(jpeg_file, p, 
         width = 25, height = 20, 
         units = "in", dpi = 600)
  
  # 打印进度
  cat("已生成:", measure_type, "的可视化文件\n")
}

#### 第3个可视化：按疾病类型分面，展示不同性别在各年龄组的指标值分布-----------
for (measure_type in measures) {
  plot_data <- df %>% 
    filter(measure == measure_type,
           year == max(df$year),  # 保持筛选最大年份
           !is.na(age))       # 保留原始值不聚合
  
  p <- ggplot(plot_data, aes(x = age, y = val, color = sex, group = sex)) +  # 直接使用val
    geom_line(linewidth = 0.8) +
    geom_point(size = 1.5) +
    scale_color_jama() +
    labs(
      title = paste0(measure_type, " by Age and Sex (", max(df$year), ")"),  # 添加年份标注
      x = "Age Group",
      y = "Rate per 100,000",
      color = "Sex"
    ) +
    plot_theme +
    theme(
      axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),  # 优化标签对齐
      strip.text = element_text(size = 20)  # 调整分面标题字号
    ) +  
    facet_wrap(~ cause, ncol = 4, scales = "free_y")  # 调整列数和轴范围
  
  # 设置输出文件路径
  filename <- gsub(" ", "_", measure_type)
  
  # 保存PDF
  pdf_file <- file.path(output_dir, paste0(filename, " facet cause - x age.pdf"))
  ggsave(pdf_file, p, 
         width = 25, height = 20, 
         units = "in", dpi = 600)
  
  # 保存JPEG
  jpeg_file <- file.path(output_dir, paste0(filename, " facet cause - x age.jpg"))
  ggsave(jpeg_file, p, 
         width = 25, height = 20,  
         units = "in", dpi = 600)
  
  # 打印进度
  cat("已生成:", measure_type, "的可视化文件\n")
}


#### 第4个可视化：按年份分面，展示不同性别在各年龄组的指标值平均分布-----------
for (measure_type in measures) {
  plot_data <- df %>%
    filter(measure == measure_type,
           !is.na(age),
           cause == big_hat) %>%#筛选大帽子进行可视化
    group_by(age, year, sex) %>%
    summarise(mean_val = mean(val, na.rm = TRUE))
  
  p <- ggplot(plot_data, aes(x = age, y = mean_val, color = sex, group = sex)) +
    geom_line(linewidth = 0.8) +  # 改用折线图展示年龄分布
    geom_point(size = 1.5) +
    scale_color_jama() +
    labs(
      title = paste0(measure_type, " by Age and Sex Over Years"),
      x = "Age Group",
      y = "Rate per 100,000",
      color = "Sex"
    ) +
    plot_theme +
    theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 1),  # 倾斜显示年龄标签
          strip.text = element_text(size = 20, face = "bold")  # 强化分面标题
          )  + 
    facet_wrap(~ year, ncol = 6, scales = "free_y")  # 分面布局调整
  
  # 设置输出文件路径
  filename <- gsub(" ", "_", measure_type)
  
  # 保存PDF
  pdf_file <- file.path(output_dir, paste0(filename, " facet year - x age.pdf"))
  ggsave(pdf_file, p, 
         width = 25, height = 20, 
         units = "in", dpi = 600)
  
  # 保存JPEG
  jpeg_file <- file.path(output_dir, paste0(filename, " facet year - x age.jpg"))
  ggsave(jpeg_file, p, 
         width = 25, height = 20, 
         units = "in", dpi = 600)
  
  # 打印进度
  cat("已生成:", measure_type, "的可视化文件\n")
}

# 完成提示
cat("\n所有可视化文件已保存至:", output_dir)


###    合并小图----------------------------------------------------------------
rm(list = ls())
library(magick)
library(purrr)

# 设置路径
input_dir <- "visualization"  # 输入文件夹路径
output_dir <- "combined_plots"  # 输出文件夹路径
if (!dir.exists(output_dir)) dir.create(output_dir)

# 定义要处理的四个measure
measures <- c("DALYs", "Deaths", "Incidence", "Prevalence")

# 遍历每个measure进行合并
walk(measures, ~{
  # 获取当前measure对应的所有PNG文件（按数字排序）
  img_files <- list.files(
    path = input_dir,
    pattern = paste0("^", .x, " .*\\.jpg$"),  # 匹配以measure前缀开头的PNG文件
    full.names = TRUE
  ) %>% 
    stringr::str_sort(numeric = TRUE)  # 按数字顺序排序（如DALYs_1, DALYs_2,...）
  
  # 确保找到4个文件
  if (length(img_files) != 4) {
    message("Warning: Found ", length(img_files), " images for ", .x)
    return()
  }
  
  # 读取并添加ABCD标签
  images <- map(1:4, ~{
    img <- image_read(img_files[.x]) %>%
      image_annotate(
        text = LETTERS[.x],  # 使用字母A-D
        gravity = "northwest",  # 左上角
        location = "+20+20",    # 偏移量
        size = 300,              # 字体大小
        color = "grey30",          # 字体颜色
        weight = 700            # 字体粗细
      )
    return(img)
  })
  
  # 将四张图片合并为2x2布局（两列）
  combined <- image_montage(
    image_join(images),
    tile = "2x2",  # 2列2行
    geometry = "x1200+50+50",  # 设置高度为1200像素，间距50
    bg = "white"
  )
  
  # 定义输出文件名
  filename <- paste0(.x, "_combined")
  
  # 保存为PNG
  image_write(combined, 
              path = file.path(output_dir, paste0(filename, ".png")),
              format = "png", quality = 100)
  
  # 保存为PDF
  image_write(combined, 
              path = file.path(output_dir, paste0(filename, ".pdf")),
              format = "pdf")
  
  message("Saved combined plot for: ", .x)
})

