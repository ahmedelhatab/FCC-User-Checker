setwd(".")
library(tidyverse)
library(xlsx)
rm(list=ls())
#Cairo

## PC
cairo_pc <- read.xlsx(file="./cairo.xlsx",sheetName = "pc",encoding = 'UTF-8')
names(cairo_pc) <- c("name","national_id","gender","department","general_department","sector","sectors","presidency","user_name","email","direct_manager","direct_manager_email","mobile","employee_id","user_ip","remarks","comments")
cairo_pc <- cairo_pc[,(1:17)]
wrong_email_cairo_pc <- cairo_pc %>% filter(grepl("@",email,fixed = FALSE)==FALSE & is.na(email)==FALSE)
wrong_manager_email_cairo_pc <- cairo_pc %>% filter(grepl("@",direct_manager_email,fixed=FALSE)==FALSE & is.na(direct_manager_email)==FALSE)
wrong_national_id_length_cairo_pc<- cairo_pc %>% mutate(national_id_length=str_length(national_id)) %>% filter(national_id_length!=14)
empty_national_id_cairo_pc <- cairo_pc %>% filter(is.na(national_id))
empty_employee_id_cairo_pc <- cairo_pc %>% filter(is.na(employee_id))
empty_gender_cairo_pc <- cairo_pc %>% filter(is.na(gender))
wrong_mobile_number_employee_cairo_pc <- cairo_pc %>% filter(str_length(mobile)<10)


write.xlsx(wrong_national_id_length_cairo_pc,file="./cairo_pc_issues.xlsx",sheetName = "wrong_national_id")
write.xlsx(empty_national_id_cairo_pc,file="./cairo_pc_issues.xlsx",sheetName = "empty_national_id",append = TRUE)
write.xlsx(empty_gender_cairo_pc,file="./cairo_pc_issues.xlsx",sheetName="empty_gender",append = TRUE)
write.xlsx(wrong_email_cairo_pc,file="./cairo_pc_issues.xlsx",sheetName = "wrong_email",append = TRUE)
write.xlsx(wrong_mobile_number_employee_cairo_pc,file="./cairo_pc_issues.xlsx",sheetName = "wrong_mobile",append = TRUE)


rm(list=c("wrong_email_cairo_pc"
          ,"wrong_manager_email_cairo_pc","wrong_national_id_length_cairo_pc",
          "empty_employee_id_cairo_pc"
          ,"wrong_mobile_number_employee_cairo_pc",
          "wrong_mobile_number_employee_cairo_pc"
          ,"empty_gender_cairo_pc"
))

## Tablet

cairo_tablet <- read.xlsx(file="./cairo.xlsx",sheetName = "tablet",encoding = 'UTF-8')
names(cairo_tablet) <- c("name","national_id","gender","department","general_department","sector","sectors","presidency","user_name","email","direct_manager","direct_manager_email","mobile","employee_id","user_ip","remarks","comments")
cairo_tablet <- cairo_tablet[,(1:17)]
wrong_email_cairo_tablet <- cairo_tablet %>% filter(grepl("@",email,fixed = FALSE)==FALSE & is.na(email)==FALSE)
wrong_manager_email_cairo_tablet <- cairo_tablet %>% filter(grepl("@",direct_manager_email,fixed=FALSE)==FALSE & is.na(direct_manager_email)==FALSE)
wrong_national_id_length_cairo_tablet<- cairo_tablet %>% mutate(national_id_length=str_length(str_trim(national_id))) %>% filter(national_id_length!=14)
empty_national_id_cairo_tablet <- cairo_tablet %>% filter(is.na(national_id))
empty_employee_id_cairo_tablet <- cairo_tablet %>% filter(is.na(employee_id))
empty_gender_cairo_tablet <- cairo_tablet %>% filter(is.na(gender))
wrong_mobile_number_employee_cairo_tablet <- cairo_tablet %>% filter(str_length(mobile)<10)

write.xlsx(wrong_national_id_length_cairo_tablet,file="./cairo_tablet_issues.xlsx",sheetName = "wrong_national_id")
write.xlsx(empty_national_id_cairo_tablet,file="./cairo_tablet_issues.xlsx",sheetName = "empty_national_id",append = TRUE)
write.xlsx(empty_gender_cairo_tablet,file="./cairo_tablet_issues.xlsx",sheetName="empty_gender",append = TRUE)
write.xlsx(wrong_email_cairo_tablet,file="./cairo_tablet_issues.xlsx",sheetName = "wrong_email",append = TRUE)
write.xlsx(wrong_mobile_number_employee_cairo_tablet,file="./cairo_tablet_issues.xlsx",sheetName = "wrong_mobile",append = TRUE)

rm(list=c("wrong_email_cairo_tablet"
          ,"wrong_manager_email_cairo_tablet",
          "empty_national_id_cairo_tablet",
          "empty_employee_id_cairo_tablet"
          ,"wrong_national_id_length_cairo_tablet",
          "wrong_mobile_number_employee_cairo_tablet"
          ,"empty_gender_cairo_tablet"
))
#Delta
## tablet
delta <- read.xlsx(file="./delta.xlsx",sheetName = "tablet",encoding = 'UTF-8')
names(delta) <- c("name","national_id","gender","department","general_department","sector","sectors","presidency","user_name","email","direct_manager","direct_manager_email","mobile","employee_id","user_ip","remarks","comments")
delta <- delta[,(1:17)]
wrong_email_delta <- delta %>% filter(grepl("@",email,fixed = FALSE)==FALSE & is.na(email)==FALSE)
wrong_manager_email_delta <- delta %>% filter(grepl("@",direct_manager_email,fixed=FALSE)==FALSE & is.na(direct_manager_email)==FALSE)
wrong_national_id_length_delta <- delta %>% mutate(national_id_length=str_length(str_trim(national_id))) %>% filter(national_id_length!=14)
empty_national_id_delta <- delta %>% filter(is.na(national_id))
empty_employee_id_delta <- delta %>% filter(is.na(employee_id))
empty_gender_delta <- delta %>% filter(is.na(gender))
wrong_mobile_number_employee_delta <- delta %>% filter(str_length(mobile)<10)
#wrong_gender_delta <- delta %>% filter(gender not %in%("")))

write.xlsx(wrong_manager_email_delta,file="./delta_issues.xlsx",sheetName = "wrong_manager_email",append=TRUE)
write.xlsx(wrong_email_delta,file="./delta_issues.xlsx",sheetName = "wrong_email",append=TRUE)
write.xlsx(wrong_national_id_length_delta,file="./delta_issues.xlsx",sheetName = "wrong_national_id")
write.xlsx(empty_national_id_delta,file="./delta_issues.xlsx",sheetName = "empty_national_id",append = TRUE)
write.xlsx(empty_gender_delta,file="./delta_issues.xlsx",sheetName="empty_gender",append = TRUE)
write.xlsx(wrong_email_delta,file="./delta_issues.xlsx",sheetName = "wrong_email",append = TRUE)
write.xlsx(wrong_mobile_number_employee_delta,file="./delta_issues.xlsx",sheetName = "wrong_mobile",append = TRUE)

rm(list=c("wrong_manager_email_delta","wrong_mobile_number_employee_delta"
          ,"wrong_email_delta","empty_national_id_delta",
          "wrong_national_id_length_delta"
          ,"empty_national_id_length_delta","empty_employee_id_delta"
          ,"empty_gender_delta"
          ))
#Canal
## Tablet
canal_tablet <- read.xlsx(file="./canal.xlsx",sheetName = "tablet",encoding = 'UTF-8')
names(canal_tablet) <- c("name","national_id","gender","department","general_department","sector","sectors","presidency","user_name","email","direct_manager","direct_manager_email","mobile","employee_id","user_ip","remarks","comments")
canal_tablet <- canal_tablet[,(1:17)]
wrong_email_canal_tablet <- canal_tablet %>% filter(grepl("@",email,fixed = FALSE)==FALSE & is.na(email)==FALSE)
wrong_manager_email_canal_tablet <- canal_tablet %>% filter(grepl("@",direct_manager_email,fixed=FALSE)==FALSE & is.na(direct_manager_email)==FALSE)
wrong_national_id_length_canal_tablet<- canal_tablet %>% mutate(national_id_length=str_length(national_id)) %>% filter(national_id_length!=14)
empty_national_id_cancal_tablet <- canal_tablet %>% filter(is.na(national_id))
empty_employee_id_canal_tablet <- canal_tablet %>% filter(is.na(employee_id))
empty_gender_canal_tablet <- canal_tablet %>% filter(is.na(gender))
wrong_mobile_number_employee_canal_tablet <- canal_tablet %>% filter(str_length(mobile)<10)

write.xlsx(wrong_email_canal_tablet ,file="./canal_tablet_issues.xlsx",sheetName = "wrong_email")
write.xlsx(wrong_manager_email_canal_tablet ,file="./canal_tablet_issues.xlsx",sheetName = "wrong_manager_email")
write.xlsx(wrong_national_id_length_canal_tablet,file="./canal_tablet_issues.xlsx",sheetName = "wrong_national_id")
write.xlsx(empty_national_id_cancal_tablet,file="./canal_tablet_issues.xlsx",sheetName = "empty_national_id",append = TRUE)
write.xlsx(empty_gender_canal_tablet,file="./canal_tablet_issues.xlsx",sheetName="empty_gender",append = TRUE)
write.xlsx(wrong_email_canal_tablet,file="./canal_tablet_issues.xlsx",sheetName = "wrong_email",append = TRUE)
write.xlsx(wrong_mobile_number_employee_canal_tablet,file="./canal_tablet_issues.xlsx",sheetName = "wrong_mobile",append = TRUE)


rm(list=c("wrong_mobile_number_employee_canal_tablet"
          ,"empty_gender_canal_tablet",
          "empty_employee_id_canal_tablet"
          ,"empty_national_id_cancal_tablet"
          ,"wrong_national_id_length_canal_tablet"
          ,"wrong_email_canal_tablet"))


## PC
canal_pc <- read.xlsx(file="./canal.xlsx",sheetName = "pc",encoding = 'UTF-8')
names(canal_pc) <- c("name","national_id","gender","department","general_department","sector","sectors","presidency","user_name","email","direct_manager","direct_manager_email","mobile","employee_id","user_ip","remarks","comments")
canal_pc <- canal_pc[,(1:17)]
wrong_email_canal_pc <- canal_pc %>% filter(grepl("@",email,fixed = FALSE)==FALSE & is.na(email)==FALSE)
wrong_manager_email_canal_pc <- canal_pc %>% filter(grepl("@",direct_manager_email,fixed=FALSE)==FALSE & is.na(direct_manager_email)==FALSE)
wrong_national_id_length_canal_pc<- canal_pc %>% mutate(national_id_length=str_length(national_id)) %>% filter(national_id_length!=14)
empty_national_id_cancal_pc <- canal_pc %>% filter(is.na(national_id))
empty_employee_id_canal_pc <- canal_pc %>% filter(is.na(employee_id))
empty_gender_canal_pc <- canal_pc %>% filter(is.na(gender))
wrong_mobile_number_employee_canal_pc <- canal_pc %>% filter(str_length(mobile)<10)


write.xlsx(wrong_national_id_length_canal_pc,file="./canal_pc_issues.xlsx",sheetName = "wrong_national_id")
write.xlsx(wrong_email_canal_pc,file="./canal_pc_issues.xlsx",sheetName = "wrong_email",append = TRUE)
write.xlsx(wrong_manager_email_canal_pc,file="./canal_pc_issues.xlsx",sheetName = "wrong_manager_email",append = TRUE)
write.xlsx(empty_national_id_cancal_pc,file="./canal_pc_issues.xlsx",sheetName = "empty_national_id",append = TRUE)
write.xlsx(empty_gender_canal_pc,file="./canal_pc_issues.xlsx",sheetName="empty_gender",append = TRUE)
write.xlsx(wrong_email_canal_pc,file="./canal_pc_issues.xlsx",sheetName = "wrong_email",append = TRUE)
write.xlsx(wrong_mobile_number_employee_canal_pc,file="./canal_pc_issues.xlsx",sheetName = "wrong_mobile",append = TRUE)

rm(list=c("wrong_mobile_number_employee_canal_pc",
          "empty_gender_canal_pc",
          "empty_employee_id_canal_pc",
          "empty_national_id_cancal_pc",
          "wrong_national_id_length_canal_pc"
          ,"wrong_email_canal_pc",
          "wrong_manager_email_canal_pc"))
