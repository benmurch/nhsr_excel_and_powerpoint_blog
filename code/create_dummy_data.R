
#create dummy datasets
#a six month period
#one staff characteristics dataset
#one scores/measures per month per employee dataset

#1. dataset 1 ####
#first, a sort of employee characteristics list
#with 100 rows each
#create some fixed characteristics for each employee
#realistically, things like payband should be in the monthly update dataframe
#and people would come and go
#but to keep this simple, we're assuming these are fixed characteristics
#and that there is no change in staff in these six months

#I've made each of the columns as individal vectors here
#and just used base R rather than tidyverse
#not because it's better, but just to make it easeir to see what's going on

how_many_employees <- 100

#the columns
empnum <- 10000 + (1:how_many_employees)
team <- paste0("Team ",sample(LETTERS,size=how_many_employees,replace=TRUE))
gender <- sample(c("M","F"),
                 size = how_many_employees,
                 replace=TRUE)
bame <- sample(c("BAME","White"),
               size = how_many_employees,
               prob = c(0.19,0.81),
               replace=TRUE)
payband <- sample(c("Band 6","Band 7", "Band 8a", "Band 8b", "Band 8c", "Band 8d", "Band 9"),
                  size=how_many_employees,
                  prob = c(0.05,
                           0.1,
                           0.2,
                           0.3,
                           0.2,
                           0.1,
                           0.05),
                  replace=TRUE)


#put in dataframe
rawdat1 <- data.frame(empnum=empnum,
                      team,
                      gender=gender,
                      bame=bame,
                      payband=payband)


#2. Dataset 2 ####
#(similar approach to Dataset 1 above)
#now create some measurements on those employees
#e.g. some sort of employee scores and data over time
#say once per month, over six months

measurement_interval <- "week"

measurement_dates <- 
  seq(from=as.Date("2021-06-01"),
    to=as.Date("2021-11-01"),
    by=measurement_interval)

dates_length <- length(measurement_dates)

#make the columns
date <- rep(measurement_dates,
            how_many_employees)

empnum2 <- unlist(lapply(empnum,function(x){rep(x,dates_length)})) #we need a copy of each employee number, for each time interval
team2 <- unlist(lapply(team,function(x){rep(x,dates_length)})) #same for the teams

score1 <- unlist(lapply(empnum,function(x){
  #this is just a silly function to give a wobbly set of time series trends
  rnorm(n=dates_length,
        mean=100+10*sin(seq(from=pi/2,to=3*pi/2,length.out=dates_length))
          + seq(1,30,length.out=dates_length),
        sd=1)
}))

#systematically bias score 1 at some paybands by ethnicity
#put make the choice of ethnicity and paybands to bias arbitrary, since this is dummy data
#vector of employee numbers to change the score for
change_score1_for <- 
  rawdat1$empnum[rawdat1$bame == sample(unique(rawdat1$bame),1) & 
                 rawdat1$payband %in% sample(unique(rawdat1$payband),3,replace=FALSE)]

#for those individuals,
  #shift their scores downwards by an average of 10 points at each date - but with random variation
#the main point of doing this here is just so there is some difference to see in the plot on slide 2
#and to make it very clear that the difference here is entirely synthetic, nothing to do with
  #actual data or anyone's ideas about it
score1[empnum2 %in% change_score1_for] <- score1[empnum2 %in% change_score1_for] +
  rnorm(length(score1[empnum2 %in% change_score1_for]),
        mean=-10,
        sd=1)


score2 <- rpois(n=how_many_employees*dates_length,lambda = 5)
flag1 <- sample(c(0,1),
                 size=how_many_employees*dates_length,
                 prob=c(0.95,0.05),
                replace=TRUE)
flag2 <- sample(c("Yes","No"),
                size=how_many_employees*dates_length,
                prob=c(0.7,0.3),
                replace=TRUE)

#put in dataframe
rawdat2 <- data.frame(empnum2,
                      team2,
                      date,
                      score1,
                      score2,
                      flag1,
                      flag2)


#Now write into the inputs folder
write.csv(rawdat1,
          "./inputs/dummy_raw_data_1.csv",
          row.names=FALSE)

write.csv(rawdat2,
          "./inputs/dummy_raw_data_2.csv",
          row.names=FALSE)