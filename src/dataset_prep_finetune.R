#Training Dataset Prep

## Loading libraries
library(tidyverse)

## Get the cleaned dataset

#source("./src/aimm_text_fig_final.R")
load("./FIG_AIMM_text.Rda")

#Summary of the data received
glimpse(AIMM_text)

#Creating Project narrative prompt

AIMM_text$prompt <- ifelse(any(is.na(AIMM_text$text.dev_impct_p),is.na(AIMM_text$text.summary)),NA,
                                   paste0("Given the following summary for development impact, please write a note on the project outcomes under the IFC Anticipated Impact Measurement and Monitoring (AIMM) framework.\n",
                                                     str_squish(AIMM_text$text.summary),"Please provide an assessment of the project outcomes in paragraphs.\n###\n"
                                                     ))
AIMM_text$completion <- ifelse(any(is.na(AIMM_text$text.dev_impct_p),is.na(AIMM_text$text.summary)),NA,
                               paste0(str_squish(AIMM_text$text.dev_impct_p))
                               )

#Checking for token
#Checking that no data is more than token limit
#if(AIMM_text %>% mutate(id=row_number()) %>% tidytext::unnest_tokens(prompt,output = "token") %>% count(id) %>% summarise(highest = max(n)) > 2000){warning("Token limit exceeded")}
#if(AIMM_text %>% mutate(id=row_number()) %>% tidytext::unnest_tokens(completion,output = "token") %>% count(id) %>% summarise(highest = max(n)) > 2000){warning("Token limit exceeded")}

jsonlite::stream_out(x = AIMM_text[,c("prompt","completion")],file("./output/FIG-BP-PRO-0001.json"))

gc()
