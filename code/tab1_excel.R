#1. tab 1 ####
#create, write, and format the table on the first tab of the spreadsheet report

#1.2 Content ####

#1.2.1 Table title/header ####

#1.2.1.1 write the text (in the first cell) ####
writeData(wb = excel_wb_formatted,
          sheet="fascinating_content",
          "Fascinating content on Tab 1")

#1.2.1.2 apply title styling ####
addStyle(wb = excel_wb_formatted,
         sheet="fascinating_content",
         t1_title_style,
         cols=1,
         rows=1)

#1.2.1.3 spread title over 4 rows and the entire column range of the table
mergeCells(wb = excel_wb_formatted,
           sheet="fascinating_content",
           cols = 1:ncol(table_1_data),
           rows = 1:4)

#1.2.2.2 write the table - starting below the title ####
writeDataTable(wb = excel_wb_formatted,
          sheet="fascinating_content",
          xy=c(1,6),
          table_1_data |> dplyr::rename(`AfC Pay Band` = payband),
          bandedRows=TRUE,
          tableStyle = "TableStyleMedium2")

#apply column header styling to table
#apply body styling to table
addStyle(wb = excel_wb_formatted,
         sheet="fascinating_content",
         t1_table_style,
         cols=1:ncol(table_1_data),
         rows=6:(6+nrow(table_1_data)+1),
         gridExpand = TRUE)

#expand the column widths
setColWidths(wb = excel_wb_formatted,
             sheet = "fascinating_content",
             cols=1:ncol(table_1_data),
             widths=rep(20,ncol(table_1_data)))

#add plot
#first you have to show it in the R studio plot window
print(plot_1 + labs(caption="Plot 1: Dummy scores by dummy payband and ethnicity"))
#now you can add it to the sheet
insertPlot(wb= excel_wb_formatted,
           sheet = "fascinating_content",
           xy=c(1,4+1+1+nrow(table_1_data)+4),
           width=30,
           height=8,
           fileType="png",
           units="cm")

dev.off()