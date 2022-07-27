library(officer)
library(flextable)
#guide on using this package is here https://ardata-fr.github.io/officeverse/
#and manuals are here https://davidgohel.github.io/officer/ (on the various tabs)

#you need to read in a template, even if you're not going to use any of the fomratting or content included there
#you set up the template (including naming its slides and placeholders within slides)
#within powerpoint itself, by going to View >> Slide Master
#when you create a slide in the code below, you specify which slide layout to use
#this will be one of the layouts you have created and named in your powerpoint template deck
#in that master layout (created in powerpoint) you can include things you always want to appear on 
  #slides of that type, in the size and location you want them to appear
  #e.g. the NHS logo in the top right corner
  #or the interlinked blue bars at the bottom of a title slide
#in theory you can insert content by specifying the name of a pre-defined location placeholder too
#but there can be problems with this, and if you really want to be specific on content size, location,
  #and fomratting, it is better to create the placeholders directly, as is done in the
  #"slide_styles_powerpoint.R" script, and use those

#1. Read in powerpoint template  ####
slides <- read_pptx(paste0(input_data_folder,"powerpoint_template.pptx"))

#2. Read in style sheet ####
#(element locations and styles)
source("./code/slide_styles_powerpoint.R")

#3. Format data tables (in Microsoft style) ####
source("./code/format_tables_powerpoint.R")

#4. Create slides and add content to them ####

#To see the format and sturcture of the powerpoint master,
  #bring up this table
    #View(layout_properties(slides))
  #use the "name" column to identify available layouts
  #and the "ph_label" column to identify available locations in a given layout
  #the names of both will be exactly as meaningful as the names you gave them
  #when creating the template in powerpoint
  #and if you've duplicated any of the names
  #or copied and pasted, or created text boxes etc instead of using the Slide Master placeholder buttons
    #to create content on the master template, you'll be in trouble

        #((( PARENTHETICAL COMMENT - IGNORE ON FIRST READING
              #if you are using the placeholder locations, you will need to specify the location within the "ph_with" fucntion
              #using a "ph_location" function which uses the name from that "ph_label" column as the value to its
              #"ph_label" argument.
              #if this sounds confusing, don't worry - we'#re not going to use that approach here
              #(but if you want to try it, look it up in the documentation)
        #)))

#In the approach used below, we will just use the "name" column to select slides
#And then use our directly created style sheet (in the "slide_styles_powerpoint.R". script)
  #to specify locations, size, font, colour etc. formats within the layout,
  #because this gives us more control



#4.1 Start with an empty slide deck
  #This means deleting the first slide in the template read in
    #(because it won't let you read in an empty deck)
slides <- remove_slide(slides,index=1)

#4.2 Add slides ####
  #You need to add slides, and add content to them, sequentially
  #i.e. add the first slide, put content into it, then create the next slide etc.


#4.2.1 Create a title slide (slide 1) ####
slides <- add_slide(slides,
                    #choose a layout
                      #in the Powerpoint template we're using, I've created a title slide called "Title Slide"
                      #it has some boilerplate content on it which will be pulled through
                      #such as has the NHS Logo in the top right,
                        #some blue bars at the bottom
                        #some text saying "NHS England and Improvement" in the centre abouve those bars
                      #all of which will be pulled through
                      layout="Title Slide",
                    master="Master")
#specify title, subtitle, and period text (includign format and location)


#4.2.1.2 Add content to the title slide ####
#the non-boilerplate stuff
#there are several steps to this, explained below

#s1 title ####
#we use ph_with to add content
slides <- ph_with(slides, #this is the name of the slide deck we've created. The active slide is the title slide we just created
                  value=fpar( #this function creates a paragraph. Whenever we're adding formatted text, it needs to be wrapped in a paragraph.
                    ftext( #this function creates formatted text.
                      text="Dummy Data Report", #this is the text we're adding
                                   prop=s1_title_format) #and this is the format we're applying - it's created in the "slide_styles_powerpoint.R" script
                    ),
                  location=s1_title_loc) #and this is the exact location (and size) that we want to place it on the slide - again, created in the "slide_styles_powerpoing.R" script

#s1 subtitle ####
#same approach as for the title - but note that we're using different text format
  #(value to the "prop" argument of the "ftext" function)
  #and location/size (value to the "location" argument to the "ph_with" function)
  #specifications - both created in the "slide_styles_powerpoint.R" script
slides <- ph_with(slides,
                  value=fpar(
                    ftext("A report that we've made for some purpose (all of this data etc. is artificial)",
                                   prop=s1_subtitle_format)),
                  location=s1_subtitle_loc)

#s1 period covered ####
#and again similar to above
  #but now we're also using the report_date argument from the user inputs section of the main script
  #then changing the display format before inclujding it within the text string
  #again, we use a specific format and location defined in our "slide_styles_powerpoint.R" script
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      paste0("The date this report was produced was: ",
                                          as.character(report_date,format="%d/%m/%Y")),
                                   s1_period_format)),
                  location=s1_period_loc)

#note that when using formatted text
#with the styles created using fp_text()
#you need to wrap the text string in a ftext() function
#with both the text string, and the format created using fp_text() as arguments
#then wrap the whole thing in an fpar() function
#or else you'll get a mysterious error



#4.2.2 Create a second slide ####
#In our template, we had a "Content Slide" layout - which just has the NHS logo in its top right corner
slides <- add_slide(slides,
                    layout="Content Slide",
                    master="Master")

#4.2.2.1 Add content to second slide ####
#s2 title ####
#same approach as above
#specific formatting, size, and placement (arguments to "prop" and "location"
  #created in the "slide_styles_powerpoint.R" script
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Fascinating content on Slide 2",
                                   prop=sx_title_format)
                    ),
                  location=sx_title_loc)

#s2 table ####
#include one of the tables as content on this slide
#this relies on three things
  #i. the raw content of the table has been created in the "script.R" script
  #ii. that content has been made into a formatted table (in Microsoft style, in this case)
    #in the "format_tables_powerpoint.R" script
  #iii. an appropriate size/location for the table has been specified in the "slide_styles_powerpoint.R"
    #script, and that location is used as the argument here
#note that we just give the name of the fomratted table directly to the "value" argument of "ph_with" here
  #we don't need (or want) to mess around with "fpar" or "ftext" (they were for directly formatting text strings)
slides <- ph_with(slides,
                  value=table1_powerpoint,
                  location=s2_table_loc)

#caption for table - place above table
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Table 1: Dummy demographic data by dummy payband",
                      prop=s2_tcaption_format)
                  ),
                  location=s2_tcaption_loc)

#s2 plot ####
#include the plot on the same slide
slides <- ph_with(slides,
                    value=plot_1,
                    location=s2_plot_loc)

#caption for this plot - place below plot
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Plot 1: Dummy scores by dummy payband and ethnicity",
                      prop=s2_pcaption_format)
                  ),
                  location=s2_pcaption_loc)

#4.2.3 add a third slide ####
#using the same default content layout as for the second slide
slides <- add_slide(slides,
                    layout="Content Slide",
                    master="Master")

#4.2.3.1  add content to third slide ####
#s3 title ####
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Horrifying content on Slide 3",
                      prop=sx_title_format)
                  ),
                  location=sx_title_loc)

#s3 table ####
slides <- ph_with(slides,
                  value=table2_powerpoint,
                  location=s3_table_loc)

#caption for table - place above table
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Table 2: Dummy scores. I'd change the style (including column widths) here if I were you (you need to look at 'table2_powerpoint' in the script 'format_tables-powerpoint.R')",
                      prop=s2_tcaption_format)
                  ),
                  location=s2_tcaption_loc)

#4.2.4 add a fourth slide ####
#using the same default content layout as for the second slide
slides <- add_slide(slides,
                    layout="Content Slide",
                    master="Master")

#4.2.4.1  add content to third slide ####
#s4 title ####
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Boring content on Slide 4",
                      prop=sx_title_format)
                  ),
                  location=sx_title_loc)

#s4 text ####
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text=c("Some text which rambles on and on and on, who knows when it will ever stop, is there anything interesting being said here? I don't think so,   ",
                                  boring_text),
                      prop=sx_text_format)
                  ),
                  location=sx_text_loc)


#4.2.5 add a fifth slide ####
slides <- add_slide(slides,
                    layout="Content Slide",
                    master="Master")
#s5 title ####
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Jazzy plot on Slide 5",
                      prop=sx_title_format)
                  ),
                  location=sx_title_loc)

#s5 plot ####
slides <- ph_with(slides,
                  value=plot_2,
                  location=s5_plot_loc)

#caption for this plot - place below plot
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="Plot 2: Mean dummy scores by team time series. Each grey line is a team. Teams B and H highlighted.",
                      prop=s2_pcaption_format)
                  ),
                  location=s5_pcaption_loc)

#a text box which someone can type their interpretation of/response to the plot in
  #before circulating the slides
slides <- ph_with(slides,
                  value=fpar(
                    ftext(
                      text="This box is just a placeholder for now - but it's where somebody is going to type in their great insights about this plot before circulating the slide deck.",
                      prop=sx_text_format)
                  ),
                  location=s5_textbox_loc)


#and so on


#5. Now save the slide deck ####
#to the outputs folder
#prefixing the date it was created
#and the name of the person who created it
#both taken from the user inputs in the "script.R" script
#to the file name

#WATCH OUT!!!
#If you've got the slide deck open, close it now before writing to it!
#or else your R session will crash
#and you'll lose any work which hasn't been saved :(
print(slides,
      target=paste0("./outputs/",
                    report_date, 
                    "-",
                    run_by,
                    "_dummy_powerpoint.pptx"))
