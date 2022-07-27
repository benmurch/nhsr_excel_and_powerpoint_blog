#library(flextable)
#library(dplyr)

#This script uses the flextable library to create formatted tables to put into the powerpoint
#Leting you control column widths, row heights, background colours, font and font size,
#padding, and other things like conditional formatting (which I haven't done here)
#You can just write dataframes straight intointo the slide deck as tables,
#but you won't have any control over the size or formatting. Hence the use of flextable.

#I've done all of this using pipes
#You could, of course, assign it all back to the table after each change
#but that would be pretty verbose

#n.b. in the "create_powerpoint.R" script I've assigned back each time
#rather than use the pipe
#because I wanted to make clear the sequential nature of creating
#(and then adding content to) each slide
#and make it easier to step into the middle of that process
#to add extra slides and/or content

#1. Table 1 - used in Slide 2 ####

#create indexes of rows to colour differently
t1_odd_rows <- seq(1,nrow(table_1_data),by=2)
t1_even_rows <- seq(2,nrow(table_1_data),by=2)

table1_powerpoint <- 
  table_1_data |> 
  rename(`AfC Pay Band` = payband) |> 
  #turn the table into a formatted "flextable"
  qflextable() |> 
  font(fontname = "Arial") |> #set font type for all of the table 
  fontsize(size=18,part="head") |> #set font size just for the header
  fontsize(size=16,part="body") |> #set font size for the rest of the table
  color(color="white",
        part="header") |> #set header font colour
  bold(part="header") |> #make header text bold
  bg(bg="#4472C4",part="header") |> #set header font background colour to the default MS table colour 
  bg(bg="#B8CCE4",i=t1_odd_rows) |> #set background fill colour for odd numbered rows (this is a sort of grey blue)
  bg(bg="#F2F2F2",i=t1_even_rows) |> #set background fill colour for the even numbered rows (this is a grey)
  padding(padding.top=0.2,padding.bottom=0.2,part="body") |> #how much space to leave around the cell contents - you can also set left and right
  vline(border=fp_border(color="white",style="solid",width=1)) |> #put a vertical white border line between each column 
  hline(border=fp_border(color="white",style="solid",width=1), #put a white horizontal border line, just between the header and body of the table (not each row)
        part="header") |>
  #setting the width of the first column to 7.5cm and that of the remaining
    #five columns to 4.5cm will force the table to be 30cm wide
      #of course, you need to eyeball-check that this displays the data in an acceptable way
  width(j=1,width=7.5,unit="cm") |> 
  width(j=2:6,width=4.5,unit="cm") |> 
  #set row heights to 1cm on the page
    #if you want a shorter table, set the row heights lower
      #but you might need to revise the font size too to make it fit
  height(part="body",height=0.8,unit="cm")

#2. Table 2 - used in slide 3 ####

#create indexes of rows to colour differently
t2_odd_rows <- seq(1,nrow(table_2_data),by=2)
t2_even_rows <- seq(2,nrow(table_2_data),by=2)

table2_powerpoint <- 
  table_2_data |> 
  rename(`AfC Pay Band` = payband,
         Ethnicity = bame,
         `Score 1`=score1_mean,
         `Score 2`=score2_mean,
         `Flag 1`=flag1_mean,
         `Head Count`=headcount) |> 
  #turn the table into a formatted "flextable"
  qflextable() |> 
  font(fontname = "Arial") |> #set font type for all of the table 
  fontsize(size=18,part="head") |> #set font size just for the header
  fontsize(size=12,part="body") |> #set font size for the rest of the table
  color(color="white",
        part="header") |> #set header font colour
  bold(part="header") |> #make header text bold
  bg(bg="#4472C4",part="header") |> #set header font background colour to the default MS table colour 
  color(color="#03FCD7",
        part="body") |> #set main table values font colour
  bg(bg="#FC03F0",i=t2_odd_rows) |> #set background fill colour for odd numbered rows
  bg(bg="#C6FC03",i=t2_even_rows) |> #set background fill colour for the even numbered rows
  padding(padding.top=0.2,padding.bottom=0.2,part="body") |> #how much space to leave around the cell contents - you can also set left and right
  vline(border=fp_border(color="black",style="solid",width=2)) |> #put a vertical black border line between each column 
  hline(border=fp_border(color="red",style="solid",width=1), #put a red horizontal border line, just between the header and body of the table (not each row)
        part="header") |>
  #setting the width of the first column to 7cm 
  #the second to 3cm, and that of the remaining
  #four columns to 4.5cm will force the table to be 30cm wide
  #of course, you need to eyeball-check that this displays the data in an acceptable way
  width(j=1,width=7,unit="cm") |> 
  width(j=2,width=3,unit="cm") |> 
  width(j=3:6,width=5,unit="cm") |> 
  #set row heights to 1cm on the page
  #if you want a shorter table, set the row heights lower
  #but you might need to revise the font size too to make it fit
  height(part="body",height=0.8,unit="cm")
