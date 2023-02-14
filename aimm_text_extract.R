#' Title: aimm_text_extract.R
#' Summary:
#' This script is aimed to scrap text narratives from a nested folder structure containing this narrative across multiple word files. 

#Load necessary files
library(tidytext)
library(tidyverse)
library(tidyr)
library(purrr)
library(readxl)
library(janitor)
library(xlsx)
library(readtext)


# Read folder location address from file
folder_paths <- read.csv(file = "./file_location.csv") #The file has location for both delegated and panel approved projects

#' Considering the variation in how different sector store their information, it might be wise to split this into sections.
#' Currently doing conditions based on sector only but possible to nuance further based on approval type. 
#' Converting most of the code below into a function

read_mas <- function(fldr_path){
  # identify the files and needed docs. (write a function )
  flds <- list.files(path = fldr_path ,recursive = TRUE) %>%
    as.data.frame() %>%
    rename(file_name = '.') %>%
    mutate(dr = file_name) %>% #use this as directory
    mutate(file_name = sub(".*/", "", dr)) %>% #get the file name when there are multiple layers of files
    mutate(fld_name = sapply(dr, function(x) {
      parts <- strsplit(x, "/")[[1]]
      part <- parts[grep("\\b[0-9]{5}\\b", parts)][1]
      part
    })) %>%
    mutate(fld_name = case_when( grepl("Amaggi Cotton", file_name) ~ "Amaggi Cotton(43740)",  #correct cases where the project number is not specified in the folder name. 
                                 grepl( "Rupshi Foods final document", file_name) ~ "Rupshi Foods final document (46329)",
                                 TRUE ~ fld_name)) %>%
    mutate(id = ifelse(grepl("\\b\\d{5}\\b", fld_name), 
                       as.numeric(gsub("[^[:digit:]]", "", regmatches(fld_name, regexpr("\\b\\d{5}\\b", fld_name)))), 
                       NA)) %>%
    filter(!grepl(".pptx|.xlsx|irm|assistant|is report|is final report|pds|esap|esrs|model",file_name,ignore.case = TRUE)) %>% #Maybe we need IS report later (content), concept note (abstract)
    filter(grepl("board paper|narrative|aimm",file_name,ignore.case = TRUE)) %>% #filter out AIMM narrative documents.
    mutate(full_dir = paste0(fldr_path,"/",dr))%>%
    filter(grepl(".docx$",dr)) #there's pdf potentially can be used. 
  
  # use the directory to extract the text. 
  aimm_mas <- flds$full_dir[which(!(flds$id %in% c(45637,42506)))] %>% map(readtext::readtext) %>% 
    bind_rows() %>%
    mutate(word_c = str_count(text, "\\b\\w+\\b")) #count the number of words
  #changed the method to remove the cases based on project id - as the warning came up at different locations
  #there's warning on 55: "IBS III (45637)/IBS_Sonagrin_Additionality_AIMM_Assessment_draft.docx" (because it's encrypted)
  #there's warning on 91: "Sunshine (43150)/Board paper Sunshine (#43150) 19Nov20 v2[2305843009214489141].docx" - The project ID for this is noted as 42506
  
  # read in the sector and country data 
  ops_full <- read_csv("./IFC_disclosure_investment_11.2022.csv") %>%
    clean_names() 
  
  ops <- ops_full %>%
    select(project_number,sector,country_description,region_description) %>%
    distinct() 
  
  test <- ops_full %>%
    select(project_number,sector,country_description,region_description,impact) %>%
    filter(!is.na(impact))
  
  #merge the text data with features. 
  mas_delegate <- flds %>%
    select(id,file_name,full_dir) %>%
    right_join(aimm_mas,c("file_name" = "doc_id")) %>%
    left_join(ops,c("id" = "project_number")) %>%
    select(id, file_name, region_description, country_description, sector, word_c, full_dir, text)
  
  save(mas_delegate, file ="mas_delegate.rda")
  
  return(mas_delegate)
}


