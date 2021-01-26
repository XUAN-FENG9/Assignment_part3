library(tidyverse)
library(dplyr)
library(readxl)
library(writexl)
library(lubridate)


info<-read_xlsx("D:/Event Study.xlsx",sheet="Event Data")
stock<-read_xlsx("D:/Event Study.xlsx",sheet = "Stock Data",col_types = c("text","text","numeric","text","guess","numeric"))
stock$Returns<-ifelse(stock$Returns%in%c("A","B","C"),NA,stock$Returns)
stock<-na.omit(stock)
stock$Returns<-as.numeric(stock$Returns)
stock$`Names Date`<-ymd(stock$`Names Date`)
stock<-stock%>%filter(stock$PERMNO%in%info$PERMNO)
stock<-stock[,-c(1,2)]
stock<-stock%>%spread(key = `PERMNO`,value = "Returns")
stock<-stock%>%filter(stock$`Names Date`>=ymd(20050101))

write_xlsx(stock,"D:/Event Study1.xlsx")


reg<-read_xlsx("D:/Event Study.xlsx",sheet = "Regression")
R<-lm(reg$CAR~reg$D_Overlev,data = reg)
summary(R)