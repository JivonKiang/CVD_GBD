# 初始化环境 ----------------------------------------------------------------
rm(list = ls())
# 安装必要包（若未安装）
# install.packages(c("openxlsx", "nih.joinpoint", "purrr", "dplyr"))

####----------------------------------------------------------------------------
###                有一个内嵌文件global_map_example是需要载入的，要复制过来
####----------------------------------------------------------------------------

####----------------------------------------------------------------------------
###                记得把原本的country analysis的文件名更换，否则会覆写
####----------------------------------------------------------------------------

rm(list = ls())
library(tools)

measures <- c("Deaths", "DALYs (Disability-Adjusted Life Years)",
              "Incidence", "Prevalence")

# 设置工作目录
# 设置固定工作目录路径（根据你的需求修改此处）
target_dir <- "E:/20250414 Wenping Gong gaint mission/E"  # 注意使用正斜杠或双反斜杠
setwd(target_dir)  

load("global_map_example.RData")

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
df <- data1 %>%
  # 将可能为因子型的列转为字符型（防止因子水平问题）
  mutate(across(c(measure, cause), as.character))# %>% 
  
  # measure_name列替换（网页1、网页4方法）
  #mutate(measure = case_when(
  #  measure == "DALYs (Disability-Adjusted Life Years)" ~ "DALYs",
  #  TRUE ~ measure
  #)) %>%
  
  # cause_name列替换（网页7、网页8方法优化）
  #mutate(cause = str_replace(
  #  cause,
  #  pattern = fixed("Multidrug-resistant tuberculosis without extensive drug resistance"),
  #  replacement = "MDR-TB without XDR"
  #)) %>%
  
  # cause_name列替换（网页7、网页8方法优化）
  #mutate(cause = str_replace(
  #  cause,
  #  pattern = fixed("Extensively drug-resistant tuberculosis"),
  #  replacement = "XDR-TB"
  #))%>%
  
  # cause_name列替换（网页7、网页8方法优化）
  #mutate(cause = str_replace(
  #  cause,
  #  pattern = fixed("Respiratory infections and tuberculosis"),
  #  replacement = "RI and TB"
  #))

# 假设 df 和 input_data 是你的两个数据框
colnames(input_data)

# df 中的 location 列和 input_data 中的 location_id 列用于匹配
df <- merge(df, input_data[,3:4], by.x = "location", by.y = "location_name", all.x = TRUE)

colnames(df)[2] <- c("measure_name")

df <- df[df$val != 0, ]

# 加载必要包 ----------------------------------------------------------------
library(tidyverse)
library(ggsci)
library(ggrepel)
library(ggpubr)
library(ggmap)
library(maps)
library(dplyr)
library(RColorBrewer)
library(sf)
library(tidyverse)
library(vroom)
library(pacman)
library(data.table)
library(ggplot2)
library(stringr)
library(sf)
library(RJSONIO)
library(RColorBrewer)
library(stringi)
library(purrr)
library(cowplot)
library(magick)
library(stringr)

measures <- c("Deaths", "DALYs (Disability-Adjusted Life Years)",
              "Incidence", "Prevalence")

# 数据预处理 ----------------------------------------------------------------


# 重构为主函数
generate_measure_maps <- function(measure){
  ###   遍历所有病因，替换原本的单病因筛选
  unique_causes <- unique(df$cause)
  
  # 遍历所有病因
  for (current_cause in unique_causes) {
    # 创建病因专用目录（路径修正）
    cause_dir <- file.path(current_cause) # 修复2：使用file.path代替拼接
    if (!dir.exists(cause_dir)) dir.create(cause_dir, recursive = TRUE)
    
    # 创建临时数据框（添加管道闭合）
    temp_df <- df %>%
      filter(cause == current_cause, measure_name == measure) # 修复3：measure列名修正
    
    # 跳过空数据集
    if (nrow(temp_df) == 5) {
      message("数据不足（<5条），跳过: ", current_cause)
      next
    }
  
    # 检查是否有有效数据
    valid_vals <- temp_df$val[!is.na(temp_df$val)]
    if (length(valid_vals) == 0) {
      message("没有有效数据，跳过: ", current_cause)
      next
    }
    
  # 合并地理数据（使用临时数据框）
  merged_map <- merge(temp_df, map1, by.x = 'location_id', by.y = 'loc_id', all.y = TRUE)
  mapsf <- st_as_sf(merged_map, crs = 4326)
    
  # 动态计算色阶范围
  val_range <- range(temp_df$val, na.rm = TRUE)
  
  # 生成分级色带
  # 生成基于百分位数的断点
  breaks <- quantile(valid_vals, 
                     probs = c(0.05, 0.15, 0.25, 0.5, 0.75, 0.85, 0.95),
                     na.rm = TRUE) %>%  signif(3) %>%
    unique()  # 关键修复：去除重复值
  
  if(length(breaks) < 2) {  # 处理无效断点
    message("有效断点不足，使用等间距方法: ", current_cause)
    breaks <- seq(min(valid_vals, na.rm = TRUE),
                  max(valid_vals, na.rm = TRUE),
                  length.out = 7) %>% signif(3)
  }
  
  # 颜色映射方案
  #color_pal <- c("#0A4290","#0653BE", "#6CACE4", "#9BCBEB", "#FFD200", "#F68D2E", "#E03C31", "#AB2929")
  
  #colors_pal <- rev(brewer.pal(8, 'RdYlBu'))
  
  color_pal <- c("#4575B4", "#74ADD1", "#ABD9E9", "#E0F3F8", "#FEE090", "#FDAE61", "#F46D43","#D73027")
  
  cause <- df$cause
  
  ####  主图  ----------------------------------------------------------
  # 更新全球地图可视化代码
  global <- ggplot(mapsf) +
    geom_sf(
      aes(fill = cut(val, breaks = c(-Inf, breaks, Inf), include.lowest = TRUE)),
      color = "black",
      size = 0.05
    ) +
    geom_sf(data = disputed, linetype = 2, fill = NA, show.legend = F) +
    scale_fill_manual(
      name = paste(measure,"Percentile"),
      values = color_pal,
      na.value = "grey80",
      labels = c(
        paste0("<", format(breaks[1], nsmall = 1)),
        paste0(format(breaks[1], nsmall = 1), "-", format(breaks[2], nsmall = 1)),
        paste0(format(breaks[2], nsmall = 1), "-", format(breaks[3], nsmall = 1)),
        paste0(format(breaks[3], nsmall = 1), "-", format(breaks[4], nsmall = 1)),
        paste0(format(breaks[4], nsmall = 1), "-", format(breaks[5], nsmall = 1)),
        paste0(format(breaks[5], nsmall = 1), "-", format(breaks[6], nsmall = 1)),
        paste0(format(breaks[6], nsmall = 1), "-", format(breaks[7], nsmall = 1)),
        paste0(">", format(breaks[7], nsmall = 1))
      )
    ) +
    guides(fill = guide_legend(
      title.position = "top",
      title.hjust = 0.5,
      nrow = 4,
      keywidth = 0.8, 
      keyheight = 0.6,
      label.position = "right"
    )) +
    theme_void() +
    theme(
      legend.position = c(0.20, 0.20),
      legend.text = element_text(size = 12),
      legend.title = element_text(size = 12, face = "bold"),
      plot.title = element_text(hjust = 0.5, size = 16, face = "bold"),
      legend.spacing.y = unit(0.5, "cm")
    ) +
    labs(title = paste("Global Burden of Disease -", current_cause,"in",max(temp_df$year),"for",measure))
  
  # 保存全球地图
  ggsave(file.path(cause_dir, paste0("global_map_", measure, ".jpg")), global, width = 16, height = 9, dpi = 300)
  ggsave(file.path(cause_dir, paste0("global_map_", measure, ".jpg")), global, width = 16, height = 9)
  
  
  ####  子图  ----------------------------------------------------------
  library(ggplot2)
  library(ggforce)
  
  # 定义区域坐标列表
  region_coords <- list(
    "Southeast Asia" = list(xlim = c(94.9, 119.1), ylim = c(-9.2, 9)),
    "West Africa" = list(xlim = c(-17.8, -7), ylim = c(6.5, 15.8)),
    "Eastern Mediterranean" = list(xlim = c(30.5, 38.5), ylim = c(28.4, 35.9)),
    "Northern Europe" = list(xlim = c(4.7, 27.5), ylim = c(48, 59)),
    "Caribbean and Central America" = list(xlim = c(-92, -59), ylim = c(7, 28)),
    "Persian Gulf" = list(xlim = c(45, 55.8), ylim = c(21, 31.5)),
    "Balkan Peninsula" = list(xlim = c(12.5, 32), ylim = c(34.5, 50))
  )
  
  theme_map_sub <- theme_void() + labs(x="", y="") + theme_bw() +
    #theme(plot.background = element_rect(fill = "lightgrey", colour = NA))
    theme(text = element_text(size = 12),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          legend.position = 'none',
          axis.text = element_blank(),
          axis.ticks = element_blank(),
          plot.background = element_rect(fill = "transparent"),
          plot.title = element_text(vjust = 0.01, hjust = 0.5))
  
  generate_regional_plots <- function(region_name) {
    global +
      #geom_sf(data = disputed, linetype = 2, fill = NA, show.legend = F) +
      coord_sf(
        xlim = region_coords[[region_name]]$xlim,
        ylim = region_coords[[region_name]]$ylim,
        expand = FALSE
      ) +
      labs(title = region_name)+
      theme(legend.position="none",
            plot.margin = margin(0, 0, 0, 0, "cm"),  # 消除子图默认边距
      )+
      theme_map_sub
  }
  
  # 生成所有区域地图
  regional_plots <- lapply(names(region_coords), function(name) {
    generate_regional_plots(name)
  })
  
  # 命名存储列表
  names(regional_plots) <- names(region_coords)
  
  library(patchwork)
  
  # 组合子图模块
  subplot_row1 <- (regional_plots[["Caribbean and Central America"]] +
                     regional_plots[["Persian Gulf"]] +
                     regional_plots[["Balkan Peninsula"]] +
                     regional_plots[["Southeast Asia"]])+plot_layout(nrow = 1)
  
  # 修改东地中海地区标题换行
  regional_plots[["Eastern Mediterranean"]] <- regional_plots[["Eastern Mediterranean"]] +
    labs(title = "Eastern\nMediterranean") +  # 添加换行符
    theme(plot.title = element_text(lineheight = 0.8))  # 调整行高
  
  regional_plots[["Northern Europe"]] <- regional_plots[["Northern Europe"]] + 
    theme(plot.margin = margin(0,0,0,0, "cm"))+ 
    theme(aspect.ratio = 1/2.2)  # 值越小，横向越拉伸（高度压缩）
  
  # 重组子图结构
  subplot_row2 <- (
    # 横向合并西非和东地中海（保持等宽）
    (regional_plots[["West Africa"]] | regional_plots[["Eastern Mediterranean"]])
  ) / 
    # 垂直叠加北欧（继承父容器宽度）
    regional_plots[["Northern Europe"]] + 
    plot_layout(
      heights = c(1, 1),  # 保持原有垂直比例
      widths = 1,             # 强制继承父容器总宽度[6](@ref)
      guides = "collect"      # 统一图例位置[6](@ref)
    )
  
  subplot_row3 <- subplot_row1 + subplot_row2 + plot_layout(nrow = 1)
  
  px <- global/subplot_row3+plot_layout(heights = c(4,1))
  
  # 输出高分辨率图片
  ggsave(file.path(cause_dir,paste0("final_map_", measure, ".jpg")), px, 
         width = 16, height = 14, dpi = 600, bg = "white")
  }
}

# 使用purrr包进行迭代处理
library(purrr)
walk(measures, generate_measure_maps)
