#Training Dataset Prep

## Loading libraries
library(tidyverse)

## Get the cleaned dataset

#source("./src/aimm_text_fig_final.R")
load("./ALL_AIMM_text.Rda")

#Summary of the data received
glimpse(AIMM_text)

#Creating Project narrative prompt
AIMM_text <- AIMM_text[which((!is.na(AIMM_text$project_description))|(!(str_squish(AIMM_text$project_description)==""))),]
AIMM_text$project_description <- gsub( pattern = "&nbsp;",replacement = "",x = AIMM_text$project_description)
AIMM_text$project_description <- stringi::stri_enc_toascii(AIMM_text$project_description)
AIMM_text$project_description <- gsub( pattern = "\032;",replacement = "",x = AIMM_text$project_description)

AIMM_text$prompt <- paste0(str_squish(AIMM_text$project_description),"\n\n###\n\n")
AIMM_text$completion <- paste0(" ",str_squish(AIMM_text$text.dev_impct_p),"\n[END]")

temp <- AIMM_text[,c("prompt","completion")]
temp <- temp[which(!is.na(temp$prompt)),]
temp <- na.omit(temp[!duplicated(temp),])

#Checking for token
#Checking that no data is more than token limit
if(temp %>% mutate(id=row_number()) %>% tidytext::unnest_tokens(prompt,output = "token") %>% count(id) %>% summarise(highest = max(n)) > 2000){warning("Token limit exceeded")}
#if(AIMM_text %>% mutate(id=row_number()) %>% tidytext::unnest_tokens(completion,output = "token") %>% count(id) %>% summarise(highest = max(n)) > 2000){warning("Token limit exceeded")}

jsonlite::stream_out(x = temp[,c("prompt","completion")],file("./output/ALL-BP-PRO-0001.json"))

gc()
