library(openxlsx)

#0. Create empty excel spreadsheet ####
#0.1 the workbook ####
excel_wb_formatted <- createWorkbook()

#0.2 the individual tabs ####
addWorksheet(excel_wb_formatted,"fascinating_content")
addWorksheet(excel_wb_formatted,"horrifying_content")
addWorksheet(excel_wb_formatted,"boring_content")
addWorksheet(excel_wb_formatted,"jazzy_plot")


source("./code/styles_excel_format.R")

source("./code/tab1_excel.R")

source("./code/tab2_excel.R")
# 
# source("./code/tab3_excel_format.R")
# 
# source("./code/tab4_excel_format.R")

#4. Save results ####
saveWorkbook(excel_wb_formatted,
             paste0("./outputs/",
                    report_date, 
                    "-",
                    run_by,
                    "_dummy_excel.xlsx"),
             overwrite = TRUE)