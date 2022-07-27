#0. Description - what this report is for

#This is a dummy template script
#Intended to show how you can write results directly to powerpoint
  #and precisely specify the formatting of that powerpoint from a script
  #so it is reporduced in the same way each time, without the need for manual manipulation


#This dummy script is intended to create:
#i. A table which summarises some staff demographics
#ii. A table which summarises some staff scores by payband and ethnicity
#iii. A bar plot of some of the data from the second table
#It reads in two tables
#And an empty powerpoint "master" template
#Then writes out a formatted powerpoint slide deck


#1. set up r session - load packages etc
library(dplyr)
library(ggplot2)


#2. user inputs - the bits somebody may need to change when running the report

#Two things you really should check and update every time
#In this script, they are simply appended to the file names of the saved outputs
report_date <- Sys.Date()
run_by <- "Senor Mysterioso"

#folders where data is read from/written to
input_data_folder <- "./inputs/"
outputs_folder <- "./outputs/"

#file names in the input folders
#to run an updated report with new data, you would put some new data in the inputs folder,
  #then change the names here to the names of those updated files
  #or else you'd archive the old data in the input file and simply add new data with these names
  #and then leave the inputs below unchanged
input_data_file1 <- "dummy_raw_data_1.csv"
input_data_file2 <- "dummy_raw_data_2.csv"
#and so on, for however many raw source files you have



#3. Read source data ####
raw_dat1 <- read.csv(paste0(input_data_folder,
                       input_data_file1))

raw_dat2 <- read.csv(paste0(input_data_folder,
                        input_data_file2))

#4. Data processing ####

#join the two tables, by employee number
joined_data <- raw_dat2 |> 
  left_join(raw_dat1,
            by=c("empnum2" = "empnum"))

#5. Produce output tables and plots ####

#make a pivot table of counts of gender and ethnicity by payband
table_1_data <- 
  raw_dat1 |> 
  group_by(payband) |>
  summarise(Males=sum(gender=="M"),
            Females=sum(gender=="F"),
            BAME=sum(bame=="BAME"),
            White=sum(bame=="White"),
            Total = n()) |> 
  ungroup()

#make a table showing average scores per month, split by payband and ethnicity
#this requires the joined data
#this data is then put into the right format for the table
  #in the format_tables_powerpoint script, using the flextable package
table_2_data <- 
  joined_data |> 
  group_by(payband,bame) |> 
  summarise(score1_mean=mean(score1),
            score2_mean=mean(score2),
            flag1_mean=mean(flag1),
            headcount=n()) |> 
  ungroup()

#because of how the data is generated, you might not have
  #values for every payband/ethnicity combination
  #and so there will not be rows for these combinations in table_2_data
  #this will not look nice when we make a dodged barplot from it
  #(we will want two columns to compare for each payband - but will only get one big fat column
    #in cases when there is zero data and hence no row for a given ethnicity category)
  #we deal with this by pivoting, replacuing the resulting NAs with zeros, and then unpivoting
  #before making the plot
  #note that since the data is randomly generated, sometimes there will be missing rows that need pivoting
    #and sometimes there won't - but this approach keeps it generalisable, rather than requiring ad hoc
plot_1_data <- 
  table_2_data |> 
  select(payband,
         bame,
         score1_mean) |> 
  tidyr::pivot_wider(names_from=bame,
              values_from=score1_mean) |> 
  tidyr::replace_na(list(White=0,BAME=0)) |> 
  tidyr::pivot_longer(-payband,
                      names_to="bame",
                      values_to="score1_mean") |> 
  mutate(textlabel=ifelse(score1_mean==0,"Nobody in this group",NA))

#plot average scores by payband
plot_1 <- 
  plot_1_data |> 
  ggplot(aes(x=payband,
             y=score1_mean,
             fill=bame)) +
  geom_col(position="dodge") +
  #when there is no data, add a label saying so
    #the NA values are automatically ignored, so text only appears where there is no bar
    #using a value of y which centres the rotated text at the middle of the plot
    #and using the dodge argument so it doesn't appear in the middle of the two bars (0.9 is the default bar width)
  geom_text(aes(label=textlabel,
                y=0.5*max(score1_mean)),
            angle=90,
            position=position_dodge(width=0.9)) +
  labs(x="",
       y="Mean of Score 1") +
  scale_fill_manual(values=c("green","purple")) +
  theme_bw() +
  theme(legend.title=element_blank(),
        #box the entire plot area
        plot.background = element_rect(colour="black",
                                       fill=NA,
                                       size=2))

#plot scores by team by week
#aggregate the data
plotdat2 <- 
  raw_dat2 |> 
  mutate(date=as.Date(date)) |>
  group_by(date,team2) |> 
  summarise(score1=mean(score1)) |> 
  ungroup()

#and plot, with a supposed policy date change marked and a couple of teams highlighted (the rest in grey)
plot_2 <- 
  plotdat2 |> 
  ggplot(aes(x=date,
             y=score1,
             group=team2)) +
  geom_line(colour="grey",alpha=0.6) +
  geom_line(data=filter(plotdat2,team2=="Team B"),
            aes(colour="Team B"),
            size=2,alpha=0.5) +
  geom_line(data=filter(plotdat2,team2=="Team H"),
            aes(colour="Team H"),
            size=2,alpha=0.5) +
  geom_vline(xintercept=as.Date("2021-09-13"),
             linetype="longdash",
             size=1) +
  annotate(geom="text",
           x=as.Date("2021-09-10"),
           y=min(plotdat2$score1),
           hjust=1,
           vjust=-1,
           label="This is when we made \nour brilliant policy change") +
  scale_colour_manual(breaks=c("Team B","Team H"),
                      values=c("#D81B60","#FFC107")) +
  scale_x_date(date_labels = "%b %Y") +
  labs(x="Date of measurement",
       y="Arbitrary Score 1") +
  theme_bw() +
  theme(panel.grid=element_blank(),
        legend.title=element_blank(),
        legend.position="bottom")

#boring text to put on slide 4/tab 3 ####
boring_text <- c(rep(c("blah, blah, blah, "), 185)," ...")


#6. Save results ####
#take the outputs produced in section 5
#put them into a formatted powerpoint slide deck
#and save that in the outputs folder
#automatically naming it in a consistent format
#which includes the date of the report, and the name of the person who ran it,
  #as entered in the user inputs in section 2
#the reason for doing this is so that we have a automatic audit trail of previous reprots
  #we can find them all
  #we known who ran them and when
  #their names and fomratting are all consistent
    #which is useful for identifying them, and if we subsequently need to read them in to another script

