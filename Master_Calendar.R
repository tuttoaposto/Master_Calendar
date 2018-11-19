library(lubridate)
library(dplyr)
library(sqldf)
library(writexl)

gsub('\\\\', '/', readClipboard())
setwd(readClipboard())
getwd()

#Create master calendar with just dates now, then add profiles to it

## Set start date = first day of the year
Y = 2019
M = 1
D = 1
strt = ymd(paste(Y,M,D, sep = '-'))
#n = days_in_month(strt)

#Add to calendar 365 dates and corresponding week numbers in the month
dow <- weekdays(dt, abbreviate = T)
month <- month(dt)
wknum <- as.integer( format(dt, format="%U") ) +1  #week number in the year
df <- data.frame(dt, dow, month, wknum, stringsAsFactors = F)

minwknum <- df %>% group_by(month) %>% summarise(minwknum = min(wknum))
df <- inner_join(df, minwknum, by = c("month" = "month")) %>% 
      mutate(wknum = wknum - minwknum +1) %>%
      select(-minwknum)



#Get physician names in template file = (all fields - first two fields) -> pick odd numbered fields
n <- count.fields(file='Master_Calendar.csv',
             sep = ',')[1] - 2
tmpl2 <- read.csv(file='Master_Calendar.csv',
                 stringsAsFactors = F,
                 #skip = 1,
                 nrows = 1,
                 colClasses = c(rep("NULL", 2), rep("character", n)),
                 header = T)

MD_name <- names(tmpl2)[(seq(1,n)%%2 == 1)]

##Add AM/PM suffix to names
MD_name_apm <- paste0(rep(MD_name, each = 2), c('_AM', '_PM'))
#MD_name_apm <- paste(rep(MD_name, each = 2), c('AM', 'PM'), sep = '_')
   
                                               
#Now read in template file, skip first 2 rows = merged fields and AM/PM fields
tmpl <- read.csv(file='Master_Calendar.csv',
                 stringsAsFactors = F,
                 skip = 2,
                 header = F,
                 col.names = c('week', 'DayofWeek', MD_name_apm)
                )
#names(tmpl)[1] <- 'week'

##Fill in week number as integer
tmpl$wknum <- rep(c(1:5), each = 7)
tmpl <- tmpl[c(11, 1:10)]

#Now fill in new calendar with MD's profiles by matching weeknum's
MC <- df %>% left_join(tmpl, by = c("wknum" = "wknum", "dow" = "DayofWeek")) %>%
  select(-c(month, week))
head(MC); tail(MC)

write_xlsx(MC, path = 'MC.xlsx')
