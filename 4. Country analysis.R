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
zip_file <- list.files(pattern = "\\.zip$")[4]          # 取当前目录第一个ZIP文件
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

# 加载必要包 ----------------------------------------------------------------
library(tidyverse)
library(ggsci)
library(ggrepel)
library(ggpubr)

measures <- c("Deaths", "DALYs", "Incidence", "Prevalence")

df <- data

library(ggmap)
library(maps)
library(dplyr)
# 创建可视化目录
if (!dir.exists("visualization")) dir.create("visualization")

# 加载必要的包
library(tidyverse)
library(ggrepel)
library(ggpubr)
library(ggsci)
library(maps)
library(ggmap)

# 数据预处理 ----------------------------------------------------------------
# 筛选最大和最小年份
min_year <- min(df$year)
max_year <- max(df$year)

# 创建合并地图数据
world_map <- map_data("world") %>%
  mutate(region = ifelse(region == "USA", "United States", region))

# 获取df$location的所有类别
levels_location <- levels(as.factor(df$location))
print(levels_location)

# 获取world_map$region的所有类别
levels_region <- levels(as.factor(world_map$region))
print(levels_region)


# 处理国家名称差异
country_name_mapping <- c(
  # 处理长格式国家名称
  "United States of America" = "United States",
  "Russian Federation" = "Russia",
  "United Kingdom" = "UK",
  "Bolivia (Plurinational State of)" = "Bolivia",
  "Venezuela (Bolivarian Republic of)" = "Venezuela",
  "Iran (Islamic Republic of)" = "Iran",
  "Syrian Arab Republic" = "Syria",
  "Democratic People's Republic of Korea" = "North Korea",
  
  # 处理特殊字符国家
  "Côte d'Ivoire" = "Ivory Coast",
  "Lao People's Democratic Republic" = "Laos",
  "Türkiye" = "Turkey",
  
  # 处理地区名称差异
  "Taiwan (Province of China)" = "China",
  "Republic of Korea" = "South Korea",
  "United Republic of Tanzania" = "Tanzania",
  "United States Virgin Islands" = "Virgin Islands",
  
  # 处理国际标准名称差异
  "Czechia" = "Czech Republic",
  "Republic of Moldova" = "Moldova",
  "North Macedonia" = "Macedonia",
  "Viet Nam" = "Vietnam",
  
  # 处理特殊行政区
  "Hong Kong SAR China" = "Hong Kong",
  "Macao SAR China" = "Macau",
  
  # 处理重复名称
  "Congo" = "Republic of Congo",
  "Democratic Republic of the Congo" = "Democratic Republic of the Congo"
)

# 创建输出目录
if (!dir.exists("visualization")) dir.create("visualization")

library(purrr)
library(cowplot)
library(magick)

# 创建分图存储目录
if (!dir.exists("visualization/single_maps")) dir.create("visualization/single_maps", recursive = TRUE)

# 修改后的可视化函数 ----------------------------------------------------------
library(purrr)
library(cowplot)
library(magick)
library(stringr)

# 安全文件名生成函数
safe_name <- function(string) {
  str_replace_all(string, "[^[:alnum:]]", "_") %>% 
    str_trunc(50, "right")
}

# 增强版可视化函数 ----------------------------------------------------------
create_enhanced_report <- function(output_format) {
  # 筛选最大年份数据
  df_max_year <- df %>%
    filter(year == max_year) %>%
    mutate(location = recode(location, !!!country_name_mapping))
  
  # 颜色映射方案
  color_pal <- c("#2A5CAA", "#6CACE4", "#9BCBEB", "#FFD200", "#F68D2E", "#E03C31", "#AB2929")
  
  # 按measure循环生成报告
  walk(unique(df_max_year$measure), function(m) {
    # 创建measure专属目录
    measure_dir <- file.path("visualization", m)
    if (!dir.exists(measure_dir)) dir.create(measure_dir, recursive = TRUE)
    
    # 筛选当前measure数据
    df_measure <- df_max_year %>% 
      filter(measure == m) %>%
      drop_na(val)
    
    # 生成所有cause的增强地图
    walk(unique(df_measure$cause), function(c) {
      # 动态计算色阶范围
      val_range <- df_measure %>% 
        filter(cause == c) %>% 
        pull(val) %>% 
        range(na.rm = TRUE)
      
      # 生成分级色带
      breaks <- signif(seq(val_range[1], val_range[2], length.out = 7), 3)
      
      # 构建地图数据
      plot_data <- world_map %>%
        left_join(df_measure %>% filter(cause == c), 
                  by = c("region" = "location"))
      
      # 获取TOP3国家标注
      top3 <- plot_data %>%
        group_by(region) %>%
        summarise(val = first(val)) %>%
        top_n(3, val) %>%
        left_join(world_map %>% group_by(region) %>%
                    summarise(long = median(long), lat = median(lat)))
      
      # 构建主地图
      p <- ggplot(plot_data) +
        geom_polygon(aes(x = long, y = lat, group = group, fill = val),
                     color = "white", size = 0.1) +
        geom_label_repel(
          data = top3,
          aes(x = long, y = lat, label = str_wrap(sprintf("%s=%.1f", 
                                                       region, val), width = 20)),
          size = 6,                    # 优化标签字号
          box.padding = 0.5,           # 增大标签间距
          point.padding = 50,         # 增加标签与点的距离
          segment.size = 0.2,          # 适当的引导线粗细
          segment.alpha = 0.9,         # 添加透明度
          max.overlaps = 10,           # 允许更多重叠
          direction = "both"#,          # 灵活调整标签位置
          #nudge_x = 0,                 # 横向微调
          #nudge_y = 1                  # 纵向微调
        ) +
        scale_fill_gradientn(
          colors = color_pal,
          na.value = "grey90",
          limits = val_range,
          breaks = breaks,
          guide = guide_colorbar(
            title = "Rate per 100k",
            barwidth = unit(16, "cm"),  # 缩短颜色条
            barheight = unit(0.2, "cm"),  # 压扁颜色条
            title.position = "top",
            title.theme = element_text(size = 12),  # 缩小标题字号
            label.theme = element_text(size = 12)  # 缩小标签字号
          )
        ) +
        theme(
          legend.position = c(0.5, 0),  # 从原0.15上调到0.25
          legend.box.just = "top"
        )+
        coord_fixed(1.2) +  # 调整地图宽高比
        labs(title = NULL) +        # 移除标题留空
        labs(title = str_wrap(c, width = 40)) +
        theme_void() +
        theme(
          plot.title = element_text(size = 12, face = "bold", hjust = 0.5, 
                                    margin = margin(t = 0, b = 0)),  # 减少标题下边距
          legend.position = "bottom",
          legend.title = element_text(size = 10),
          legend.text = element_text(size = 10),
          legend.margin = margin(t = -30, b = 0),  # 上移图例
          legend.box.margin = margin(t = -30),  # 减少图例容器边距
          plot.margin = margin(t = -50, r = 0, b = -50, l = 0, unit = "pt")     # 四边负边距
        )
      
      # 保存独立图表
      ggsave(
        filename = file.path(measure_dir, paste0(safe_name(c), ".", output_format)),
        plot = p,
        width = ifelse(output_format=="pdf", 8, 20),
        height = ifelse(output_format=="pdf", 6, 12),
        units = "cm",
        dpi = 300
      )
    })
  })
}

# 执行生成 --------------------------------------------------------------

# 清除历史生成文件
unlink("visualization", recursive = TRUE)

walk(c("jpg", "pdf"), function(fmt) {
  tryCatch({
    create_enhanced_report(fmt)
  }, error = function(e) {
    message("生成", fmt, "时遇到问题:", e$message)
  })
})

# 小图拼接 --------------------------------------------------------------
# 小图拼接 --------------------------------------------------------------
# 小图拼接 --------------------------------------------------------------

rm(list = ls())
# 加载必要的包
library(jpeg)
library(grDevices)

# 设置参数
base_dir <- "visualization"
measures <- c("Incidence", "Deaths", "DALYs", "Prevalence")
output_res <- 300  # 输出分辨率（DPI）

# 创建通用绘图函数
plot_images <- function(img_paths, columns = 2, rows = 5) {
  # 设置绘图区域
  par(mfrow = c(rows, columns), 
      mar = rep(0.1, 4),  # 减少边距
      oma = rep(0.1, 4),
      xaxs = "i", 
      yaxs = "i")
  
  # 循环绘制所有图片
  for (img in img_paths) {
    tryCatch({
      # 读取并绘制图片
      img_data <- readJPEG(img)
      plot(NA, xlim = c(0,1), ylim = c(0,1), 
           axes = FALSE, xlab = "", ylab = "")
      rasterImage(img_data, 0, 0, 1, 1)
    }, error = function(e) message("Error processing: ", img))
  }
}

# 主处理循环
for (measure in measures) {
  measure_dir <- file.path(base_dir, measure)
  
  # 跳过不存在的目录
  if (!dir.exists(measure_dir)) {
    message("跳过不存在的目录: ", measure)
    next
  }
  
  # 获取并排序图片文件
  img_files <- list.files(measure_dir, 
                          pattern = "\\.jpg$", 
                          full.names = TRUE,
                          ignore.case = TRUE)
  
  if (length(img_files) == 0) {
    message(measure, " 文件夹中没有JPG图片")
    next
  }
  
  # 自然排序文件名（处理数字编号的情况）
  img_files <- img_files[order(gsub("[^0-9]", "", img_files))]
  
  # 设置输出文件名
  output_base <- paste0(measure, "_combined")
  
  # PDF输出参数
  pdf(file = paste0(output_base, ".pdf"),
      width = 8.27,   # A4纸的英寸宽度（纵向）
      height = 11.69) # A4纸的英寸高度
  plot_images(img_files)
  dev.off()
  
  # JPG输出参数
  jpeg(file = paste0(output_base, ".jpg"),
       width = 8.27 * output_res,
       height = 11.69 * output_res,
       res = output_res,
       quality = 100)
  plot_images(img_files)
  dev.off()
  
  message("已完成处理: ", measure)
}

message("所有处理完成！")