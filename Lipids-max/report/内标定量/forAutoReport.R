library(openxlsx)
library(stringr)
library(dplyr)
library(ggplot2)
library(grid)
library(gridExtra)

python <- "/workspace/.venv/bin/python"

#### ---- Sub functions -----#######

ArgsVerify <- function(inDir, repDir, repLogo) {
  head_file <- grep("^headMessage", list.files(inDir), value = T)[1]
  if (is.na(head_file)) stop("缺少headmessage文件。")


  info_path <- paste0(inDir, "/sampleInfo.xlsx")
  sample_info <- readxl::read_excel(info_path, sheet = 1)
  sample_detail <- NULL
  sheets <- readxl::excel_sheets(info_path)
  if ("details" %in% sheets) {
    sample_detail <- readxl::read_excel(info_path, sheet = "details")
  }

  head_path <- paste0(inDir, "/", head_file)
  headMessage <-
    openxlsx::read.xlsx(
      head_path,
      sheet = 1,
      colNames = F,
      rowNames = T
    )

  product_name <-
    openxlsx::read.xlsx(
      head_path,
      sheet = "项目类型",
      colNames = T,
      rowNames = T
    )
  sample_type <-
    openxlsx::read.xlsx(
      head_path,
      sheet = "样本类型",
      colNames = T,
      rowNames = T
    )


  spl_type <- headMessage["样本类型", 1] # 固体
  proj_name <- headMessage["项目名称", 1] #

  unit_CC <- sample_type[spl_type, "计算浓度"] # 计算浓度
  unit_CM <- sample_type[spl_type, "样品浓度"] # 样品浓度

  pro_type <- headMessage["项目类型", 1] # 报告流程[RPT_FLOW] AQ700
  ddf_flow <- product_name[pro_type, "分析流程[DDF_FLOW]"]
  qtf_type <- product_name[pro_type, "定量类型[QTF_TYPE]"]
  rpt_flow <- product_name[pro_type, "报告流程[RPT_FLOW]"]
  grp_type <- product_name[pro_type, "合并方式[GRP_TYPE]"]

  outDir <-
    glue::glue(
      "{wd}/{code}-Experimental-Results",
      wd = inDir,
      code = headMessage["项目编号", 1],
      name = headMessage["客户姓名", 1]
    )
  outDirWord <-
    glue::glue(
      "{wd}/{code}-Experimental-Word-Report",
      wd = inDir,
      code = headMessage["项目编号", 1],
      name = headMessage["客户姓名", 1]
    )
  outData <-
    glue::glue(
      "{wd}/{code}-RawData",
      wd = inDir,
      code = headMessage["项目编号", 1],
      name = headMessage["客户姓名", 1]
    )
  outPDF <-
    glue::glue(
      "{out}/{code}-{name}-{type}.pdf",
      out = outDir,
      code = headMessage["项目编号", 1],
      name = headMessage["客户姓名", 1],
      type = headMessage["项目名称", 1]
    )
  outWord <-
    glue::glue(
      "{out}/{code}-{name}-{type}.docx",
      out = outDirWord,
      code = headMessage["项目编号", 1],
      name = headMessage["客户姓名", 1],
      type = headMessage["项目名称", 1]
    )

  outTab <-
    glue::glue(
      "{out}/{code}-{name}-{type}-定量结果.xlsx",
      out = outDir,
      code = headMessage["项目编号", 1],
      name = headMessage["客户姓名", 1],
      type = headMessage["项目名称", 1]
    )


  resDir <- paste0(inDir, "/results/")
  repLogo <- ifelse("报告是否含LOGO" %in% rownames(headMessage), switch(headMessage["报告是否含LOGO", 1],
    "是" = TRUE,
    "否" = FALSE
  ), T)

  l.a <- list(
    head_msg = headMessage,
    head_path = head_path,
    molecule_name = "Compound Name",
    spl_name = headMessage["项目名称", 2],
    inDir = inDir,
    outDir = outDir,
    outPDF = outPDF,
    outData = outData,
    outWord = outWord,
    outTab = outTab,
    repDir = repDir,
    resDir = resDir,
    grp_type = grp_type,
    rpt_flow = rpt_flow,
    qtf_type = qtf_type,
    spl_type = spl_type,
    unit_CM = unit_CM,
    unit_CC = unit_CC,
    repLogo = repLogo,
    proj_name = proj_name,
    spl_info = sample_info,
    spl_detail = sample_detail,
    FIG = 0
  )
}

fig_count <- function(ret = TRUE, fig = "") {
  l.a$FIG <<- l.a$FIG + 1
  if (ret) {
    return(glue::glue("{f}{i}", f = fig, i = l.a$FIG))
  }
}


get_table_sheet <- function(f, sheet = "merge", default = 1) {
  sheets <- readxl::excel_sheets(f)
  if (sheet %in% sheets) {
    df_data <- readxl::read_xlsx(f, sheet = sheet, na = "#N/A")
  } else {
    df_data <- readxl::read_xlsx(f, sheet = default, na = "#N/A")
  }
  return(df_data)
}

get_table_quant <- function(l.a = l.a, filename = NA, overwrite = TRUE, sheet = "merge") {
  molecule_name <- l.a$molecule_name

  data_sampleInfo <- openxlsx::read.xlsx(paste0(l.a$inDir, "/sampleInfo.xlsx"), na.strings = "#N/A")
  qc_name_rsd <- data_sampleInfo$sample_name[data_sampleInfo$sample_type == "QC"]
  spl_list <- data_sampleInfo %>%
    filter(sample_type == "SPL") %>%
    select(sample_name) %>%
    unique()


  data_q1 <- get_table_sheet(paste0(l.a$resDir, "/Quantification-Raw.xlsx"), sheet = sheet)
  data_q2 <- get_table_sheet(paste0(l.a$resDir, "/Quantification-Spl.xlsx"), sheet = sheet) %>% select(-c(molecule | contains("$")))


  if (length(qc_name_rsd) > 1) {
    data_qc <- data_q1 %>% select(c(contains("$GRP_TYPE"), !!qc_name_rsd))
  } else {
    data_qc <- NULL
  }

  data_q1 <-
    data_q1 %>%
    select(-c(molecule | contains(c("$", "HQC", "MQC", "LQC")))) %>%
    select(-!!qc_name_rsd)

  cpds <- data_q1[molecule_name] %>% unlist(use.names = FALSE)

  colnames(data_q1)[1] <- sprintf("%s (%s)", molecule_name, l.a$unit_CC)
  colnames(data_q2)[1] <- sprintf("%s (%s)", molecule_name, l.a$unit_CM)


  wb <- createWorkbook()
  modifyBaseFont(wb, fontSize = 10, fontName = "Arial")
  addWorksheet(wb, sheetName = "下机浓度")
  writeDataTable(wb, sheet = 1, data_q1, keepNA = T, na.string = "N/A", withFilter = F)
  addWorksheet(wb, sheetName = "样本浓度")
  writeDataTable(wb, sheet = 2, data_q2, keepNA = T, na.string = "N/A", withFilter = F)
  setColWidths(wb, sheet = 1, cols = 1, "auto")
  setColWidths(wb, sheet = 2, cols = 1, "auto")
  setRowHeights(wb, sheet = 1, rows = 1, 20)
  setRowHeights(wb, sheet = 1, rows = 2:(nrow(data_q1) + 1), 16)
  setRowHeights(wb, sheet = 2, rows = 1, 20)
  setRowHeights(wb, sheet = 2, rows = 2:nrow(data_q2), 16)
  addStyle(
    wb,
    sheet = 1, rows = 1:(nrow(data_q1) + 1), cols = 2:ncol(data_q1),
    gridExpand = T, stack = T, style = createStyle(halign = "center")
  )
  addStyle(
    wb,
    sheet = 2, rows = 1:(nrow(data_q2) + 1), cols = 2:ncol(data_q2),
    gridExpand = T, stack = T, style = createStyle(halign = "center")
  )
  addStyle(
    wb,
    sheet = 1, rows = 1:(nrow(data_q1) + 1), cols = 1:ncol(data_q1),
    gridExpand = T, stack = T, style = createStyle(valign = "center")
  )
  addStyle(
    wb,
    sheet = 2, rows = 1:(nrow(data_q2) + 1), cols = 1:ncol(data_q2),
    gridExpand = T, stack = T, style = createStyle(valign = "center")
  )
  saveWorkbook(wb, filename, overwrite = overwrite)

  return(list(qc = data_qc, cpds = cpds))
}

get_table_calib <- function(l.a = l.a, filename = NA, sheet = "merge", sheetName = "Calibration", overwrite = TRUE) {
  molecule_name <- l.a$molecule_name

  data_cal <- get_table_sheet(paste0(l.a$resDir, "/Calibrators.best.xlsx"), sheet = sheet)
  data_cal <- data_cal[, c(molecule_name, "cc_lod", "cc_loq", "xmax", "score")]
  data_cal[, 2:3] <- round(data_cal[, 2:3], 2)
  data_cal$score <- round(data_cal$score, 4)

  colnames(data_cal) <- c(molecule_name, "LLOD", "LLOQ", "ULOQ", "R^2")

  data_S3 <- data_cal
  colnames(data_S3)[1] <- sprintf("%s (%s)", colnames(data_S3)[1], l.a$unit_CC)
  if (!is.na(filename)) {
    if (overwrite) {
      wb <- createWorkbook()
    } else {
      wb <- loadWorkbook(filename)
    }
    modifyBaseFont(wb, fontSize = 10, fontName = "Arial")
    addWorksheet(wb, sheetName = sheetName)
    writeDataTable(wb, sheet = sheetName, data_S3, keepNA = T, withFilter = F)
    setColWidths(wb, sheet = sheetName, cols = 1:ncol(data_S3), "auto")
    setRowHeights(wb, sheet = sheetName, rows = 1, 20)
    setRowHeights(wb, sheet = sheetName, rows = 2:(nrow(data_S3) + 1), 16)
    addStyle(
      wb,
      sheet = sheetName, rows = 1:(nrow(data_S3) + 1), cols = 1:ncol(data_S3),
      gridExpand = T, stack = T, style = createStyle(halign = "center", valign = "center")
    )
    saveWorkbook(wb, filename, overwrite = TRUE)
  }
  return(data_cal)
}

fill_rsd <- function(x, v = 30) {
  x[is.na(x)] <- 1000
  x0 <- x
  x0[x <= v] <- sprintf("%.1f", x0[x <= v])
  x0[x > v] <- sprintf(">%.1f", v)
  return(x0)
}

get_table_stats <- function(l.a = l.a, filename = NA, fill = NA, sheet = "merge") {
  data_rsd <-
    paste0(l.a$resDir, "/Quantification-Stats.xlsx") %>%
    get_table_sheet(sheet = sheet) %>%
    select(-c(molecule | contains("$")))

  data_rsd$`Accuracy[%]` <- data_rsd$`Accuracy[%]` %>% round(1)
  data_rsd$`RSD[%]` <- data_rsd$`RSD[%]` %>% round(1)
  if (!is.na(fill)) data_rsd$`RSD[%]` <- data_rsd$`RSD[%]` %>% fill_rsd(fill)
  colnames(data_rsd)[1] <- sprintf("%s (%s)", colnames(data_rsd)[1], l.a$unit_CC)

  if (!is.na(filename)) {
    wb <- createWorkbook()
    modifyBaseFont(wb, fontSize = 10, fontName = "Arial")
    addWorksheet(wb, sheetName = "Calibration")
    writeDataTable(wb, sheet = 1, data_rsd, keepNA = T, withFilter = F)
    setColWidths(wb, sheet = 1, cols = 1:ncol(data_rsd), "auto")
    setRowHeights(wb, sheet = 1, rows = 1, 20)
    setRowHeights(wb, sheet = 1, rows = 2:(nrow(data_rsd) + 1), 16)
    addStyle(
      wb,
      sheet = 1, rows = 1:(nrow(data_rsd) + 1), cols = 1:ncol(data_rsd),
      gridExpand = T, stack = T, style = createStyle(halign = "center", valign = "center")
    )
    saveWorkbook(wb, filename, overwrite = T)
  }
  return(data_rsd)
}

corr_plot <- function(df.in, platform = "") {
  # Calculate the correlation matrix of QC
  # Draw the correlation, histogram & regression
  # for each QC & QC-pair

  plotQC0 <- function(i, df.c, axismax, axismin) {
    # Draw a single QC plot

    row.i <- df.c$row[i]
    col.i <- df.c$col[i]
    p <- ggplot(data = data.frame(df.qc))
    if (row.i == col.i) {
      # browser()
      p <- p + geom_histogram(
        aes(log10(df.qc[, col.i])),
        bins = 10, color = "#477aad", fill = "#477aad"
      ) +
        theme_bw() +
        theme(
          # legend.position = "right",
          legend.title = element_blank(),
          legend.key = element_blank(),
          axis.title = element_blank(),
          # axis.ticks = element_blank(),
          # axis.text = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(color = "black", size = 1)
        ) +
        ggplot2::annotate(
          "text",
          label = colnames(df.qc)[row.i],
          x = Inf, y = Inf, hjust = 0.5, vjust = 1.5, size = 0.005
        )
    } else if (row.i < col.i) {
      df.p <- data.frame(
        x = -log10(df.qc[, row.i]), y = -log10(df.qc[, col.i])
      )
      p <- p + geom_blank(
        data = df.p, aes(x, y)
      ) +
        # xlim(axismin, axismax) +
        # ylim(axismin, axismax) +
        theme_bw() +
        theme(
          legend.title = element_blank(),
          legend.key = element_blank(),
          axis.title = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(color = "black", size = 1)
        ) +
        ggplot2::annotate(
          "text",
          label = round(m.cor[row.i, col.i], 3),
          x = 0.5, y = 0.5, size = 1, color = "#477aad"
        )
    } else {
      df.p <- data.frame(
        x = -log10(df.qc[, row.i]), y = -log10(df.qc[, col.i])
      )
      p <- p + geom_point(
        data = df.p, shape = 16, color = "#477aad", size = 8, aes(x, y)
      ) +
        xlim(axismin, axismax) +
        ylim(axismin, axismax) +
        geom_smooth(
          data = df.p, aes(x, y), formula = y ~ x,
          method = "lm", color = "red", fullrange = T
        ) +
        theme_bw() +
        theme(
          legend.title = element_blank(),
          legend.key = element_blank(),
          axis.title = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(color = "black", size = 1)
        )
    }
    return(p)
  }

  plotQC <- function(i, df.c, axismax, axismin) {
    # Draw a single QC plot
    row.i <- df.c$row[i]
    col.i <- df.c$col[i]
    p <- ggplot(data = data.frame(df.qc))
    if (row.i == col.i) {
      # browser()
      p <- p + geom_histogram(
        aes(log10(df.qc[, col.i])),
        bins = 20, color = "#477aad", fill = "#477aad"
      ) +
        theme_bw() +
        theme(
          # legend.position = "right",
          legend.title = element_blank(),
          legend.key = element_blank(),
          axis.title = element_blank(),
          # axis.ticks = element_blank(),
          # axis.text = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(color = "black", size = 4)
        ) +
        ggplot2::annotate(
          "text",
          label = colnames(df.qc)[row.i],
          x = Inf, y = Inf, hjust = 1.5, vjust = 3.5, size = 15
        )
    } else if (row.i < col.i) { # upper
      df.p <- data.frame(
        x = -log10(df.qc[, row.i]), y = -log10(df.qc[, col.i])
      )
      p <- p + geom_blank(
        data = df.p, aes(x, y)
      ) +
        # xlim(axismin, axismax) +
        # ylim(axismin, axismax) +
        theme_bw() +
        theme(
          legend.title = element_blank(),
          legend.key = element_blank(),
          axis.title = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(color = "black", size = 4)
        ) +
        ggplot2::annotate(
          "text",
          label = round(m.cor[row.i, col.i], 3),
          x = 0.5, y = 0.5, size = 16, color = "#477aad"
        )
    } else {
      df.p <- data.frame(
        x = -log10(df.qc[, row.i]), y = -log10(df.qc[, col.i])
      )
      p <- p + geom_point(
        data = df.p, shape = 16, color = "#477aad", size = 8, aes(x, y)
      ) +
        xlim(axismin, axismax) +
        ylim(axismin, axismax) +
        geom_smooth(
          data = df.p, aes(x, y), formula = y ~ x,
          method = "lm", color = "red", fullrange = T
        ) +
        theme_bw() +
        theme(
          legend.title = element_blank(),
          legend.key = element_blank(),
          axis.title = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          panel.border = element_rect(color = "black", size = 4)
        )
    }
    return(p)
  }


  df.data <- data.frame(df.in)
  df.data[, ] <- as.numeric(unlist(df.data[, ]))
  df.data[df.data == 0] <- NA
  exp.min <- min(apply(df.data, MARGIN = 2, min, na.rm = T))
  df.qc <- sapply(df.data, function(column) {
    column[is.na(column)] <- exp.min / 2
    return(column)
  })
  m.qc <- as.matrix(df.qc)
  m.cor <- cor(
    m.qc,
    method = "spearman",
    use = "na.or.complete" # NA methods
  )
  if (min(m.cor) <= 0.8) {
    m.cor <- cor(
      m.qc,
      method = "pearson",
      use = "na.or.complete" # NA methods
    )
  }
  xymin <- -log10(max(m.qc, na.rm = T)) - 1 # limits of the axis
  xymax <- -log10(min(m.qc, na.rm = T)) + 1
  n.cl <- ncol(m.cor)
  df.coord <- data.frame(
    row = rep(1:n.cl, each = n.cl),
    col = rep(1:n.cl, n.cl)
  )
  if (nrow(m.cor) >= 50) {
    l.pics <- lapply(
      1:nrow(df.coord), plotQC0,
      df.c = df.coord, axismax = xymax, axismin = xymin
    )
  } else if (nrow(m.cor) < 50) {
    l.pics <- lapply(
      1:nrow(df.coord), plotQC,
      df.c = df.coord, axismax = xymax, axismin = xymin
    )
  }
  names(l.pics) <- paste0("p", df.coord$row, df.coord$col)
  p.grid <- arrangeGrob(grobs = l.pics, nrow = n.cl)
  # return(m.cor)
  # ggsave(
  #   paste0(l.a$outDir,'/QC corrplot', platform, '.jpg'), p.grid,
  #   width = 35, height = 35, units = 'in', dpi = 300
  # )
  return(p.grid)
}


copy_file <- function(d_base, d_sub, patt, d_dest) {
  d_path <- paste(d_base, d_sub, sep = "/")
  for (f in list.files(d_path, pattern = patt)) {
    d_file <- paste(d_path, f, sep = "/")
    file.copy(d_file, to = d_dest)
  }
}

## escape special chars of latex
latex_escape <- function(x) {
  gsub("_", "\\\\_", x)
}


## ---------------- help functions end ---------------------##


reportAssist <- function(inDir = "", repDir = "", repLogo = T) {
  library(openxlsx)
  library(stringr)
  library(dplyr)
  library(ggplot2)
  library(grid)
  library(gridExtra)
  l.a <<- ArgsVerify(inDir, repDir, repLogo)
  if (!file.exists(l.a$outDir)) dir.create(l.a$outDir)
  if (!file.exists(l.a$outWord)) dir.create(dirname(l.a$outWord), recursive = TRUE)
  sprintf("chmod -R 777 '%s'", l.a$outDir) %>% system()

  report_rmd <- glue::glue("{rep}/{rpt}/rmd/", rep = l.a$repDir, rpt = l.a$rpt_flow)
  file.copy(report_rmd, l.a$outDir, recursive = TRUE, overwrite = TRUE)
  print(l.a)
  ## functions used in RMD
  l.a$get_table_quant <<- get_table_quant
  l.a$get_table_calib <<- get_table_calib
  l.a$get_table_stats <<- get_table_stats
  l.a$get_table_sheet <<- get_table_sheet
  l.a$corr_plot <<- corr_plot
  l.a$copy_file <<- copy_file
  l.a$fig_count <<- fig_count
  l.a$latex_escape <<- latex_escape

  # Sys.setenv(RSTUDIO_PANDOC ="/usr/lib/rstudio-server/bin/quarto/bin/tools/x86_64/pandoc")
  ## for pdf
  report_temp_pdf <- glue::glue("{rmd}/rmd/{tmp}", rmd = l.a$outDir, tmp = "report_template.Rmd")
  rmarkdown::render(report_temp_pdf, output_file = l.a$outPDF, output_format = "bookdown::pdf_document2")

  ## for word
  report_temp_doc <- glue::glue("{rmd}/rmd/{tmp}", rmd = l.a$outDir, tmp = "report_template_word.Rmd")
  if (file.exists(report_temp_doc)) {
    report_temp_ref <- glue::glue("{rmd}/rmd/{tmp}", rmd = l.a$outDir, tmp = "word_template.docx")
    output_format <- rmarkdown::word_document(reference_docx = report_temp_ref)
    rmarkdown::render(report_temp_doc, output_file = l.a$outWord, output_format = output_format)
  }
  glue::glue("{py} {rep}/FeishuRecord.py {ind}", py = python, rep = rep_dir, ind = l.a$inDir) %>% system()
  glue::glue("{py} {rep}/Sample_Info.py {ind} {out}", py = python, rep = rep_dir, ind = l.a$inDir, out = l.a$outDir) %>% system()
  glue::glue("{py} {rep}/ExpDataPack.py {ind}", py = python, rep = rep_dir, ind = l.a$inDir) %>% system()
  paste0(l.a$outDir, "/rmd") %>% unlink(recursive = TRUE)
}


#### Run report -----------
rep_dir <- paste0(scriptloc::script_dir_get(), "/") %>% gsub("\\\\", "/", .)
args <- commandArgs(trailingOnly = T)

# 初始化日志
library(logger)
sink(paste0(args[1], "/log.txt"), append = TRUE, split = TRUE)

tryCatch({
  glue::glue("=========== Report Start ({date}) ===========\n\n", date = Sys.Date()) %>% cat()
  reportAssist(inDir = args[1], repDir = rep_dir)
}, error = function(e) {
  cat("Error: ", e$message, "\n")
}, finally = {
  sink() # 确保结束日志
})
