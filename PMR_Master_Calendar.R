library(lubridate)
library(dplyr)
library(sqldf)
library(readxl)
library(writexl)

#Set library to "S:/RWC Master Calendar/Physical Medicine/2019"
setwd(gsub('\\\\', '/', readClipboard()))
getwd()

#Set calendar date for current year
Y = 2019
M = 1
D = 1
strt = ymd(paste(Y,M,D, sep = '-'))
n = days_in_month(strt)

dt <- strt+c(0: 364)
dow <- weekdays(dt, abbreviate = T)
month <- month(dt)
#wknum <- as.integer( format(dt, format="%U") ) +1

#Set 'weeknum' acc. order of day of week shown up in the month
df <- data.frame(dt, dow, month) %>% group_by(month) %>% arrange(dow,dt) %>%
                 group_by(month, dow) %>% mutate(wknum = 1:n()) %>%
                 arrange(dt)

#Get week num for each month
#minwknum <- df %>% group_by(month) %>% summarise(minwknum = min(wknum))
#df <- inner_join(df, minwknum) %>% 
#  mutate(wknum = wknum - minwknum +1) %>%
#  select(-minwknum)

#Adjust week number: if month starts on Fri or Sat, set weeknum = 1 through second Sat of the month
#because 1st week always starts on Mon, Tue, Wed, Thur
#Fri schedule is always the same regardless of the week
# firstday <- df %>% group_by(month) %>% filter(row_number()==1 & dow =='Fri') %>%
#                    rename(firstday = dow) %>%
#                    select(month, firstday)
# 
# df1 <- df %>% left_join(select(firstday, month, firstday), by = 'month') %>%
#               mutate(wknum = ifelse(!is.na(firstday) & firstday == 'Fri' & wknum >1, wknum-1, wknum)) %>%
#               select(-firstday)

df1 <-df

#Now read in calendar template from .xlsx
templ <- read_excel('PMR_Master_Calendar_Template.xlsx', sheet = 1, col_names =FALSE, skip=2)

#Set column nam33
MD <- rep(c('ANGUYEN' ,'FIRTCH', 'HEILMAN' ,'LAIRAY', 'MARATUK', 'QUESADA', 'TREINEN' ,'AHN' ,'TOY', c(10:41)) ,each = 2)
      
x <-paste(MD, '_',rep(c('AM', 'PM'), 41), sep='')
x <- c('Day', x, 'Add_New', 'Notes')

colnames(templ)<- x

#Enumerate week number in template
templ$wknum <- c(rep(1, each=6), rep(2:5, each = 7), 6)

#Set month to the month master calendar is created for
m <- 12

#Join calendar date and template
cur_month <- df1[df1$month == m, -3]
this_month <- cur_month %>% left_join(templ, by = c('wknum', 'dow' = 'Day')) 

this_month <- this_month[c(3, 1, 2, 4:87)]


#Export to .xlsx and copy to master calendar
writexl::write_xlsx(this_month, 'this_month.xlsx')
