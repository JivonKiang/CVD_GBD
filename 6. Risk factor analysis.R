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
zip_file <- list.files(pattern = "\\.zip$")[6]          # 取当前目录第一个ZIP文件
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
  ))#%>%
  
  # cause_name列替换（网页7、网页8方法优化）
  #mutate(cause = str_replace(
  #  cause,
  #  pattern = fixed("Respiratory infections and tuberculosis"),
  #  replacement = "RI and TB"
  #))


df <- data

df <- df %>% filter(val != 0)

df <- df %>% arrange(desc(cause))


###   数据读取完毕--------------------------------------------------------------
# 获取当前工作目录
current_dir <- getwd()
print(paste("当前工作目录为:", current_dir))

# 加载必要包 ----------------------------------------------------------------
# 创建可视化主目录
vis_root <- file.path(current_dir, "visualization")
if (!dir.exists(vis_root)) dir.create(vis_root)

# 获取所有measure类型
measures <- unique(df$measure)

# 修改后的热图绘制函数
# 修改后的热图绘制函数
create_heatmap <- function(data, measure_name, cause_name) {
  ggplot(data, aes(x = factor(year), y = rei)) +  
    geom_tile(aes(fill = val), color = "white", linewidth = 0.3) +
    scale_fill_gradientn(
      colours = c("#2A5CAA",  # 深蓝色(Dark Blue)
                  "#6CACE4",  # 浅蓝色(Light Blue) 
                  "#9BCBEB",  # 淡蓝色(Pale Blue)
                  "#FFD200",  # 明黄色(Yellow)
                  "#F68D2E",  # 橙色(Orange)
                  "#E03C31",   # 红色(Red)
                  "#AB2929"),
      breaks = seq(min(data$val), max(data$val), length.out = 7),
      name = "Value",
      guide = guide_legend(
        direction = "horizontal",
        nrow = 1,
        title.position = "top",
        label.position = "bottom",
        label.theme = element_text(angle = 15, hjust = 0.25, vjust = 0.25),
        position = "bottom"
      )
    ) +
    labs(
      title = paste(measure_name, "-", cause_name),
      x = "Year", 
      y = "Risk Factor"
    ) +
    theme_minimal(base_size = 16) +  # 基础字体放大至12pt
    theme(
      axis.text.x = element_text(
        angle = 45, 
        hjust = 1, 
        vjust = 1.1,  # 微调垂直对齐
        size = 16,    # X轴标签字体
        margin = margin(t = 5)  # 增加顶部间距
      ),
      axis.text.y = element_text(
        size = 16,    # Y轴标签字体
        margin = margin(r = 5)  # 增加右侧间距
      ),
      plot.title = element_text(
        size = 16, 
        face = "bold",
        margin = margin(b = 10)  # 标题与图形间距
      ),
      legend.text = element_text(size = 16),
      legend.title = element_text(size = 16),
      panel.grid = element_blank()
    )
}

library(patchwork)
library(ggplot2)

for (m in measures) {
  # 创建measure子目录
  measure_dir <- file.path(vis_root, m)
  if (!dir.exists(measure_dir)) dir.create(measure_dir)
  
  # 筛选当前measure数据
  df_measure <- df %>% filter(measure == m)
  causes <- unique(df_measure$cause)
  
  plot_list <- list()
  
  # 生成单个cause的图
  for (c in causes) {
    df_cause <- df_measure %>% filter(cause == c)
    p <- create_heatmap(df_cause, m, c)
    
    # 保存单个图
    ggsave(
      filename = file.path(measure_dir, paste0(gsub("/", "_", c), ".png")),
      plot = p,
      width = 8,   # 宽度缩小25%
      height = 6,  # 高度缩小25%
      dpi = 600,   # 提高分辨率
      units = "in"
    )
    
    plot_list[[c]] <- p
  }
  
  # 合并所有cause的图
  combined_plot <- wrap_plots(plot_list, ncol = 4) + 
    plot_annotation(title = m, theme = theme(plot.title = element_text(hjust = 0.5)))
  
  ggsave(
    filename = file.path(measure_dir, paste0("Combined_", m, ".png")),
    plot = combined_plot,
    width = 60,  # 7列需更宽画布
    height = 10 * ceiling(length(causes)/7),  # 动态计算高度
    limitsize = FALSE
  )
}

###  合并小图-------------------------------------------------------------------
rm(list = ls())
library(magick)

# ---- 参数设置 ----
a4_width_mm <- 210   # A4纵向宽度（单位：mm）
dpi <- 600           # 分辨率（建议≥300dpi）
a4_width_px <- round(a4_width_mm / 25.4 * dpi)  # 转换为像素宽度（4960px）

# ---- 加载图片并缩放 ----
measure_folders <- c("Deaths", "DALYs")
img_paths <- unlist(lapply(measure_folders, function(folder) {
  list.files(file.path("visualization", folder), 
             pattern = "^Combined_.*\\.(jpg|jpeg|png)$", 
             full.names = TRUE, ignore.case = TRUE)
}))
if (length(img_paths) == 0) stop("未找到符合条件的图片文件")

imgs <- lapply(img_paths, function(path) {
  img <- image_read(path)
  # 关键点：仅缩放宽度，高度按比例自动调整（保护宽高比）
  img_scaled <- image_scale(img, geometry = a4_width_px)
  return(img_scaled)
})

# ---- 合并图片 ----
combined_img <- image_append(image_join(imgs), stack = TRUE)

# ---- 输出PDF（适配A4尺寸） ----
image_write(
  combined_img,
  path = "combined_output.pdf",
  format = "pdf",
  density = dpi  # 根据DPI和像素尺寸自动计算物理尺寸[7,8](@ref)
)

# ---- 输出JPG ----
image_write(
  combined_img,
  path = "combined_output.jpg",
  format = "jpeg",
  quality = 100
)