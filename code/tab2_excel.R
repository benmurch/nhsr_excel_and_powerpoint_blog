#1.2.1.1 write the text (in the first cell) ####
writeData(wb = excel_wb_formatted,
          sheet="horrifying_content",
          "Horrifying content on Tab 2")

#1.2.1.2 apply title styling ####
addStyle(wb = excel_wb_formatted,
         sheet="horrifying_content",
         t1_title_style,
         cols=1,
         rows=1)

#1.2.1.3 spread title over 4 rows and the entire column range of the table
mergeCells(wb = excel_wb_formatted,
           sheet="horrifying_content",
           cols = 1:ncol(table_2_data),
           rows = 1:4)

writeDataTable(wb = excel_wb_formatted,
               sheet="horrifying_content",
               xy=c(1,6),
               table_2_data |> dplyr::rename(`AfC Pay Band` = payband,
                                             Ethnicity = bame,
                                             `Score 1`=score1_mean,
                                             `Score 2`=score2_mean,
                                             `Flag 1`=flag1_mean,
                                             `Head Count`=headcount),
               bandedRows=TRUE,
               tableStyle = "TableStyleMedium2")

addStyle(wb = excel_wb_formatted,
         sheet = "horrifying_content",
         t2_table_style_header,
         cols = 1:ncol(table_2_data),
         rows = 6)

#identify odd and even table rows
oddrows <- seq(7,(7+nrow(table_2_data)),by=2)
evenrows <- seq(8,(6+nrow(table_2_data)),by=2)

#apply separate syles for odd and even rows
addStyle(wb = excel_wb_formatted,
         sheet = "horrifying_content",
         t2_table_style_odd,
         cols = 1:ncol(table_2_data),
         rows = oddrows,
         gridExpand = TRUE)

addStyle(wb = excel_wb_formatted,
         sheet = "horrifying_content",
         t2_table_style_even,
         cols = 1:ncol(table_2_data),
         rows = evenrows,
         gridExpand = TRUE)


setColWidths(wb = excel_wb_formatted,
             sheet = "horrifying_content",
             cols=1:ncol(table_2_data),
             widths=c(30,8,rep(30,ncol(table_2_data)-2)))