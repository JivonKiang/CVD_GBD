rm(list = ls())
library(tools)


####----------------------------------------------------------------------------
###                有一个内嵌文件夹external_publications是需要载入的，要复制过来
####----------------------------------------------------------------------------


# 设置工作目录
# 设置固定工作目录路径（根据你的需求修改此处）
target_dir <- "E:/20250414 Wenping Gong gaint mission/E"  # 注意使用正斜杠或双反斜杠
setwd(target_dir)  

###----location位置排排坐
# 加载地理位置层次数据
library(dplyr)
locations <- read.csv("external_publications/hierarchies/location_GBD2021.csv") %>% 
  select(location_id, parent_id, location_name, sort_order, region_name, location_type) %>% 
  mutate(parent_id = ifelse(parent_id == 1, NA, parent_id))  # 将Global的parent设为NA

# 自动获取ZIP文件并处理路径
list.files(pattern = "\\.zip$")
zip_file <- list.files(pattern = "\\.zip$")[1]          # 取当前目录第一个ZIP文件
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

# 4. 自动字符转因子
df <- data.frame(lapply(data1, function(x) {
  if (is.character(x)) factor(x) else x
}))

# 5. 动态处理measure列（可选步骤）
library(dplyr)
library(stringr)  # 用于更灵活的字符串操作

# 验证结果
cat("\n最终数据结构：\n")
str(df, max.level = 1)

##   TPC    读取
csv_file <- list.files(pattern = "\\.csv$")[2]        # 取第一个CSV文件
data1 <- read.csv(csv_file)

# 4. 自动字符转因子
df1 <- data.frame(lapply(data1, function(x) {
  if (is.character(x)) factor(x) else x
}))

# 验证结果
cat("\n最终数据结构：\n")
str(df1, max.level = 1)

# 检查 df 的列数
if (ncol(df) == 17) {
  TPC <- df  # 将列数为17的数据框重命名为 TPC
  rm(df)      # 删除原数据框 df
} else if (ncol(df) == 16) {
  AS <- df    # 将列数为16的数据框重命名为 AS
  rm(df)
}

# 对 df1 执行相同操作
if (ncol(df1) == 17) {
  TPC <- df1
  rm(df1)
} else if (ncol(df1) == 16) {
  AS <- df1
  rm(df1)
}

TPC <- TPC[,-14]
TPC[,13] <- c("TPC")

colnames(TPC)[13] <- c("year")

colnames(TPC)
colnames(AS)

#data <- rbind(AS,TPC)

data <- AS

# 5. 动态处理measure列（可选步骤）
library(dplyr)
library(stringr)  # 用于更灵活的字符串操作

# 转换列类型并执行替换（网页2、网页5方法优化）
data <- data %>%
  # 将可能为因子型的列转为字符型（防止因子水平问题）
  mutate(across(c(measure_name, cause_name), as.character)) %>% 
  
  # measure_name列替换（网页1、网页4方法）
  mutate(measure_name = case_when(
    measure_name == "DALYs (Disability-Adjusted Life Years)" ~ "DALYs",
    TRUE ~ measure_name
  )) %>%
  
  # cause_name列替换（网页7、网页8方法优化）
  mutate(cause_name = str_replace(
    cause_name,
    pattern = fixed("Multidrug-resistant tuberculosis without extensive drug resistance"),
    replacement = "MDR-TB without XDR"
  )) %>%
  
  # cause_name列替换（网页7、网页8方法优化）
  mutate(cause_name = str_replace(
    cause_name,
    pattern = fixed("Extensively drug-resistant tuberculosis"),
    replacement = "XDR-TB"
  ))%>%
  
  # cause_name列替换（网页7、网页8方法优化）
  mutate(cause_name = str_replace(
    cause_name,
    pattern = fixed("Respiratory infections and tuberculosis"),
    replacement = "RI and TB"
  ))

# 验证替换结果
cat("\n替换后measure_name取值分布:\n")
table(data$measure_name)

cat("\n替换后cause_name包含MDR的记录:\n")
data %>% filter(str_detect(cause_name, "MDR")) %>% select(cause_name) %>% distinct()


str(data)

library(dplyr)
library(tidyr)

df_formatted <- data %>%
  mutate(
    value_str = sprintf("%.2f (%.2f to %.2f)", val, lower, upper),
    header_key = paste(cause_name, measure_name, year, sep = "|")
  )

df_wide <- df_formatted %>%
  # 按元数据定义的顺序排序
  arrange(cause_name, measure_name, year) %>% 
  pivot_wider(
    id_cols = location_name,
    names_from = header_key,
    values_from = value_str,
    names_sep = "|"
  )

# 提取三级表头元数据
# 提取三级表头元数据
headers <- df_formatted %>%
  distinct(cause_name, measure_name, year) %>%
  arrange(cause_name, measure_name, year)

str(headers)

# 转换所有因子列为字符
headers$cause_name <- as.character(headers$cause_name)
headers$measure_name <- as.character(headers$measure_name)

# 生成带空列的表头矩阵（关键修改）
header_matrix <- rbind(
  c("", headers$cause_name),      # 第一行：留空+疾病名称
  c("", headers$measure_name),    # 第二行：留空+指标类型
  c("", headers$year)             # 第三行：留空+年份
) 

print(header_matrix)

# 构建三级表头矩阵
header_matrix <- rbind(
  headers$cause_name,
  headers$measure_name,
  headers$year
) %>% t()

library(openxlsx)

# 创建Workbook
wb <- createWorkbook()
addWorksheet(wb, "Results")

# 写入三级表头
writeData(wb, sheet = 1, x = t(header_matrix), startRow = 1, , 
          startCol = 2,  # 新增：从第二列开始写入
          colNames = FALSE)

# 写入数据主体
writeData(wb, sheet = 1, x = df_wide, startRow = 4, colNames = TRUE)

# 分层合并逻辑
for(lvl in 1:3) {
  current_level <- na.omit(header_matrix[lvl, ])
  rle_result <- rle(current_level)
  
  # 生成合并区域坐标
  end_cols <- cumsum(rle_result$lengths)
  start_cols <- c(1, end_cols[-length(end_cols)] + 1)
  
  # 动态计算Excel列号
  for(i in seq_along(rle_result$lengths)) {
    if(rle_result$lengths[i] > 1) {
      cols_start <- start_cols[i] + 1  # 补偿地理区域列
      cols_end <- end_cols[i] + 1
      
      mergeCells(
        wb, sheet = 1,
        cols = cols_start:cols_end,
        rows = lvl
      )
    }
  }
}

# 设置样式
header_style <- createStyle(
  textDecoration = "bold",
  halign = "center",
  valign = "center",
  border = "TopBottomLeftRight"
)

addStyle(wb, sheet = 1, style = header_style, rows = 1:3, cols = 1:ncol(header_matrix))

# 保存文件
saveWorkbook(wb, "formatted_results.xlsx", overwrite = TRUE)


####   细的不同cause的分别绘制--------------------------------------------------
library(openxlsx)
library(dplyr)
library(tidyr)

# 生成分页数据
cause_list <- unique(df_formatted$cause_name)

# 创建工作簿
wb <- createWorkbook()

# 遍历每个cause创建sheet
for(cause in cause_list) {
  # 筛选当前cause数据
  df_cause <- df_formatted %>% 
    filter(cause_name == cause) %>%
    arrange(measure_name, year)
  
  str(df_cause)
  str(locations)
  
  # 生成当前分页的宽表
  df_wide_cause <- df_cause %>%# 按元数据定义的顺序排序
    arrange(cause_name, measure_name, year) %>% 
    pivot_wider(
      id_cols = location_name,
      names_from = c(measure_name, year),
      values_from = value_str,
      names_sep = "|"
    )
  
  # ====== 新增排序逻辑 ====== 
  
  df_wide_cause <- df_wide_cause %>%
    # 关联location数据框获取排序号
    inner_join(locations %>% select(location_id,location_name, sort_order, parent_id), 
               by = "location_name") %>%
    # 按sort_order升序排列
    arrange(sort_order) %>%
    # 移除临时排序号列
    select(-sort_order)
  
  df_wide_cause <- df_wide_cause %>%
    left_join(locations %>% 
                select(parent_id = location_id, parent_name = location_name),
              by = "parent_id") %>%
    relocate(parent_name, .after = parent_id)
  
  df_wide_cause <- df_wide_cause %>%
    mutate(parent_name = case_when(
      location_name == "Global" ~ "Global",
      grepl("SDI", location_name) ~ "SDI region",
      TRUE ~ parent_name
    )) %>%
    filter(!is.na(parent_name))
  
  str(df_wide_cause)
  
  # 将 parent_name 移动到第一列
  df_wide_cause <- df_wide_cause %>%
    select(parent_name, everything())
  
  # 对 parent_name 列进行处理：保留第一个字符，其他重复的字符清空为 NA
  df_wide_cause <- df_wide_cause %>%
    mutate(parent_name = case_when(
      row_number() == 1 | parent_name != lag(parent_name) ~ parent_name,
      TRUE ~ NA_character_
    ))
  
  library(dplyr)
  df_wide_cause <- df_wide_cause %>%
    select(-location_id, -parent_id)
  # 查看结果
  print(df_wide_cause)
  
  library(dplyr)
  
  # 生成原始数据顺序标识
  df_wide_cause$orig_order <- seq_len(nrow(df_wide_cause))
  
  # 步骤1：插入新行并标记
  df_processed <- df_wide_cause %>%
    split(1:nrow(.)) %>%
    purrr::map_dfr(function(row) {
      if (all(!is.na(c(row$parent_name, row$location_name))) &&
          all(nzchar(c(row$parent_name, row$location_name)))) {
        # 创建新行
        new_row <- row
        new_row$location_name <- row$parent_name
        new_row[, setdiff(names(new_row), c("location_name", "orig_order"))] <- NA
        new_row$is_inserted <- TRUE
        # 保留原行
        row$is_inserted <- FALSE
        bind_rows(new_row, row)
      } else {
        row$is_inserted <- FALSE
        bind_rows(row)
      }
    }) %>%
    arrange(orig_order, desc(is_inserted))  # 确保插入行在上方
  
  # 步骤2：删除parent_name列
  df_processed <- df_processed %>% select(-parent_name)
  
  # 步骤3：标记需要删除的重复行
  rows_to_remove <- c()
  for (i in 1:(nrow(df_processed) - 1)) {
    if (df_processed$is_inserted[i] && 
        !is.na(df_processed$location_name[i]) &&
        df_processed$location_name[i] == df_processed$location_name[i + 1]) {
      rows_to_remove <- c(rows_to_remove, i)
    }
  }
  
  # 步骤4：执行删除并清理
  df_final <- df_processed %>%
    slice(-rows_to_remove) %>%
    select(-is_inserted, -orig_order)
  
  # 重置行号
  rownames(df_final) <- NULL
  
  # ====== 新增结束 ======
  
  # 提取当前分页的表头元数据
  headers_cause <- df_cause %>%
    distinct(measure_name, year) %>%
    arrange(measure_name, year)
  
  # 构建三级表头矩阵
  header_matrix_cause <- rbind(
    rep(cause, ncol(df_final)-1),    # 疾病名称层
    headers_cause$measure_name,          # 指标类型层
    headers_cause$year                   # 年份层
  )
  
  # 创建分页并写入数据
  addWorksheet(wb, sheetName = substr(cause,1,31))  # Excel限制31字符
  writeData(wb, sheet = cause, x = header_matrix_cause, 
            startRow = 1, startCol = 2, colNames = FALSE)
  writeData(wb, sheet = cause, x = df_final, 
            startRow = 4, colNames = TRUE)
  
  # 三级表头合并逻辑（网页4、网页8方法优化）
  for(lvl in 1:3) {
    current_level <- header_matrix_cause[lvl, ]
    rle_result <- rle(current_level)
    
    end_cols <- cumsum(rle_result$lengths)
    start_cols <- c(1, end_cols[-length(end_cols)] + 1)
    
    for(i in seq_along(rle_result$lengths)) {
      if(rle_result$lengths[i] > 1) {
        cols_range <- (start_cols[i] + 1):(end_cols[i] + 1)  # 补偿地理区域列
        mergeCells(wb, sheet = cause, cols = cols_range, rows = lvl)
      }
    }
  }
  
  # 设置表头样式（网页2、网页6方法）
  header_style <- createStyle(
    textDecoration = "bold",
    halign = "center",
    valign = "center",
    border = "TopBottomLeftRight"
  )
  addStyle(wb, sheet = cause, header_style, 
           rows = 1:3, cols = 1:ncol(header_matrix_cause), gridExpand = TRUE)
}

# 保存文件（网页1、网页3方法）
saveWorkbook(wb, "formatted_by_cause.xlsx", overwrite = TRUE)

