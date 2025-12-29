##################################################
## R script for GT autoReport
## Author: LiuYulin
## update：20230411
## modify：20250616, hq
##################################################
Sys.setenv(TZ="Asia/Shanghai")
library(rmarkdown)
library(qpdf)
library(xml2)


qtf_type = l.a$qtf_type
rpt_flow = l.a$rpt_flow
spl_type = l.a$spl_type
rep_dir <- l.a$repDir
head_msg <- l.a$head_msg
outPDF <- gsub('amp;','',l.a$outPDF)

# make cover -----
coverTemplate <- readLines(paste0(rep_dir, "cover.html"), encoding = "utf-8",warn = F)
coverTemplate <- paste(coverTemplate, collapse = "")
coverReport <-
  sprintf(
    coverTemplate,
    ifelse(l.a$repLogo, paste0(rep_dir, "report_cover.jpg"), paste0(rep_dir, "report_cover_nologo.jpg")),
    head_msg["项目编号",1],
    head_msg["项目名称",1],
    head_msg["客户单位",1],
    head_msg["客户姓名",1],
    Sys.Date()
  )
writeLines(coverReport, paste0(l.a$outDir,"/cover.html"))

# --enable-local-file-access 新版wkhtmltopdf需要开启本地文件访问权限
glue::glue(
  "wkhtmltopdf --enable-local-file-access --page-size A4 ", 
  "--margin-bottom 0cm --margin-top 0cm ",
  "--margin-left 0cm --margin-right 0cm ",
  " '{outDir}/cover.html'",
  " '{outDir}/cover.pdf'",
  outDir=l.a$outDir
) %>% system()


# make body -----

if(is.null(spl_type)) stop("检查样本类型。")

report_temp <- glue::glue("{rep}/{rpt}/rmd/{tmp}", rep=rep_dir, rpt=rpt_flow, tmp="report_template.Rmd")
file.copy(report_temp, l.a$outDir, overwrite = TRUE)
template_file <- paste0(l.a$outDir, "/report_template.Rmd")


render(input = template_file, output_file =  paste0(l.a$outDir, "/biotree_report_temp.html"))


library('rvest', verbose =F, quietly =T, warn.conflicts =F)

pdfHead <- rvest::read_html(paste0(rep_dir, "biotree_head_pdf.html"), encoding = "UTF-8")
Body <- rvest::read_html(paste0(l.a$outDir, "/biotree_report_temp.html"), encoding = "UTF-8")
reportBody <- Body %>% rvest::html_node("body") 
reportBody <- as.character(reportBody)
reportBody <- gsub("<body>", "<article>",reportBody)
reportBody <- gsub("</body>", "</article>",reportBody)
pdfReport <- paste0("<!DOCTYPE html>\n<html>", pdfHead, reportBody, "</dody>", "</html>") %>% read_html()
write_xml(pdfReport, paste0(l.a$outDir, "/biotree_report_pdf.html"))

# --enable-local-file-access 新版wkhtmltopdf需要开启本地文件访问权限
glue::glue(
  "wkhtmltopdf --enable-local-file-access --page-size A4",  
  if(l.a$repLogo) paste0(" --header-html ", rep_dir, "header.html "), 
  if(l.a$repLogo) paste0(" --footer-html ", rep_dir, "footer.html "),
  " --margin-bottom 2.5cm --margin-top 2.5cm",
  " --margin-left 1.5cm --margin-right 1.5cm",
  " '{outDir}/biotree_report_pdf.html'",
  " '{outDir}/biotree_report_temp.pdf'",
  outDir=l.a$outDir
) %>% system()


## =============== extra scripts ====================
pdf_combine(c(paste0(l.a$outDir, "/cover.pdf"), paste0(l.a$outDir, "/biotree_report_temp.pdf")), output = outPDF)
glue::glue(
  "{py} {rep}/delBlankPage.py -i '{pdf}' -o '{pdf}'", 
  py=python, rep=l.a$repDir, pdf=outPDF
) %>% system()

glue::glue(
  "{py} {rep}/Sample_Info.py {inDir} {outDir} '{pdf}'", 
  py=python, rep=l.a$repDir, inDir=l.a$inDir, 
  outDir=l.a$outDir, pdf=outPDF
) %>% system()



## =============== remove temp-files ================
file.remove(paste0(l.a$outDir,"/report_template.Rmd"))
file.remove(paste0(l.a$outDir,"/cover.html"))
file.remove(paste0(l.a$outDir,"/cover.pdf"))
file.remove(paste0(l.a$outDir,"/biotree_report_pdf.html"))
file.remove(paste0(l.a$outDir,"/biotree_report_temp.pdf"))
file.remove(paste0(l.a$outDir,"/biotree_report_temp.html"))



