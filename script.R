#1. Make fake data

source("./code/create_dummy_data.R")

#make the content, used in both PowerPoint and Excel (the same content)
source("./code/make_content.R")


#2. Make PowerPoint slides
  #There are several stages:
    #Read in raw data and create content (plots, tables etc.)
    #Format the tables
    #Define a stylesheet
    #Create the workbook (set it up, add slides, place the content on the slides, save down the output)
  #Each of these steps is done here in a separate script,
    #all of the scripts are sourced, in order, from the script below
      #(n.b. you could do it all in one big script, the logical separation here is to make
      #it easier where the individual steps are being done, and to let you pull out and change
      #a relevant bit as needed, while leaving the rest as is)
  #A description of the intended outputs is included at the head of the script below


source("./code/create_powerpoint.R")



#3. Make a formatted Excel worksheet
#Slightly simpler process here, with two stages
  #Define style sheet
  #Create workbook, write content and save


source("./code/create_excel.R")
