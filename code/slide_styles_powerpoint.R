
#locations and styles to apply
#these will over-write the placeholders in the powerpoint template

#if copying dimensions and locations from a powerpoint deck
#note that powerpoint reports these in cm
#and in general, living in the 21st century we use metric units
#but officer requires them in inches
#so in the location specifications below
#I'd recommend putting in the distances in cm
#but converting to inches, by multiplying by the cm to inches scale factor
#note that this is 0.3937
#and I've assigned it here to "cm_to_in" to use in the specifications below instead of typing the number in directly each time
  #because I think that makes what's happening more obvious
cm_to_in <- 0.3937


#Here's what you need to do, for each placeholder location:

#i. for everthing:
#specify location on slide using ph_location()
#in terms of distance from top left corner (left and top)
#and height and width of box area

#ii. just for text boxes:
#specify text formatting using fp_text()


#1. elements common to multiple slides ####

#1.1 Title ####
#slide title location (for content slides)
sx_title_loc <- ph_location(left=cm_to_in*1, #place title text box top left corner 1cm from left of slide
                            top=cm_to_in*1.2, #place title text box top left corner 1.2cm from top of slide
                            width=cm_to_in*28, #make title text box 28cm wide
                            height=cm_to_in*1.5) #make title text box 1.5cm tall

#slide title style/format (for content slides)
sx_title_format <- fp_text(font.size = 28, #make title text font size 28 point
                           bold=TRUE, #bold the title text
                           color = "#005EB8", #make the title text be in this specific shade of blue
                           font.family="Arial") #make the title text font Arial

#1.2. Generic text box ####
sx_text_loc <- ph_location(left=cm_to_in*1, #place title text box top left corner 1cm from left of slide
                            top=cm_to_in*5, #place title text box top left corner 5cm from top of slide
                            width=cm_to_in*28, #make title text box 28cm wide
                            height=cm_to_in*15) #make title text box 15cm tall

#slide title style/format (for content slides)
sx_text_format <- fp_text(font.size = 12, #make title text font size 12 point
                           bold=FALSE, #bold the title text
                           color = "black", #make the text be black
                           font.family="Arial") #make the title text font Arial



#2. Elements for specific slides ####

#2.1. Title Slide ####

#2.1.1 Title ####
s1_title_loc <- ph_location(left=cm_to_in*1.5,
                         top=cm_to_in*7,
                         width=cm_to_in*20,
                         height=cm_to_in*1.5)

s1_title_format <- fp_text(font.size = 40,
                        bold=TRUE,
                        color = "#005EB8",
                        font.family="Arial")

#2.1.2 Subtitle ####
s1_subtitle_loc <- ph_location(left=cm_to_in*1.5,
                            top=cm_to_in*10,
                            width=cm_to_in*28,
                            height=cm_to_in*1.5)

s1_subtitle_format <- fp_text(font.size = 18,
                              bold=FALSE,
                              color = "#005EB8",
                              font.family="Arial")

#2.1.3 Report date ####
s1_period_loc <- ph_location(left=cm_to_in*1.5,
                             top=cm_to_in*12,
                             width=cm_to_in*28,
                             height=cm_to_in*1.5)

s1_period_format <- fp_text(font.size = 18,
                         bold=FALSE,
                         color = "#005EB8",
                         font.family="Arial")


#2.2 Second slide ####
#2.2.1. Table ####
#You can set the top left corner location,
  #and to some extent the width here
  #but the height, and in general the table size
  #will be determined by the padding, font size etc
    #(which will over-ride the width here too)
    #which you need to set in the "format_tables_powerpoint.R" script
    #based on visual inspection of the output

#2.2.1.1 table location ####
s2_table_loc <- ph_location(left=cm_to_in*1,
                            top=cm_to_in*3.2,
                            width=cm_to_in*30)

#2.2.1.2 table caption location ####
s2_tcaption_loc <- ph_location(left=cm_to_in*0.7,
                               top=cm_to_in*2.5,
                               width=cm_to_in*28,
                               height=cm_to_in*1)

#2.2.1.3 table caption format ####
s2_tcaption_format <- fp_text(font.size = 11,
                              bold=TRUE,
                              color = "black",
                              font.family="Arial")


#2.2.2.1 Plot ####
s2_plot_loc <- ph_location(left=cm_to_in*1,
                            top=cm_to_in*10.2,
                            width=cm_to_in*30,
                            height=cm_to_in*8)

#2.2.2.2 plot caption location ####
s2_pcaption_loc <- ph_location(left=cm_to_in*0.7,
                               top=cm_to_in*18.2,
                               width=cm_to_in*28,
                               height=cm_to_in*1)

#2.2.2.3 plot caption format ####
s2_pcaption_format <- fp_text(font.size = 11,
                              bold=TRUE,
                              color = "black",
                              font.family="Arial")


#2.3 Third slide ####

#2.3.3 Table ####
#as per above, table size depends on content formatting,
  #set in the "format_tables_powerpoint.R" script
s3_table_loc <- ph_location(left=cm_to_in*1,
                            top=cm_to_in*3.8,
                            width=cm_to_in*30)


#2.5 Fifth slide ####
#where to put the time series plot
s5_plot_loc <- ph_location(left=cm_to_in*1,
                            top=cm_to_in*3.5,
                            width=cm_to_in*22,
                            height=cm_to_in*14)
#position for the caption
s5_pcaption_loc <- ph_location(left=cm_to_in*0.7,
                               top=cm_to_in*17.4,
                               width=cm_to_in*28,
                               height=cm_to_in*1)

#where to put the placeholder textbox for some gripping commentary
s5_textbox_loc <- ph_location(left = cm_to_in*24,
                              top=cm_to_in*3.5,
                              width=cm_to_in*8,
                              height=cm_to_in*14)
