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
library(officer)


# Read folder location address from file
folder_paths <- read.csv(file = "./file_location.csv") %>%
  mutate(path = gsub("\\\\", "/", path)) #The file has location for both delegated and panel approved projects

wd <-getwd()

if (grepl("xweng", getwd(),ignore.case = TRUE)) {
  folder_paths <- folder_paths %>%
    mutate(path = gsub("gjain5","xweng",path))
} # If "xweng" is using this file, replace path "gjain5" with "xweng"

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
  file_list <- flds$full_dir[which(!(flds$id %in% c(45637,42506)))]
  
  aimm_mas <- map2(file_list, seq_along(file_list), function(file, file_index) {
    doc <- read_docx(file)
    summary <- docx_summary(doc)
    summary$file_path <- file
    summary$file_index <- file_index
    summary
  }) %>% 
    bind_rows()
  
  mas_delegate <- flds %>%
    select(id,file_name,full_dir) %>%
    right_join(aimm_mas,c("full_dir" = "file_path")) 
  
  mas_delegate
  #there's warning on 55: "IBS III (45637)/IBS_Sonagrin_Additionality_AIMM_Assessment_draft.docx" (because it's encrypted)
}

mas_del_path <- folder_paths %>%
  filter(approval_type == "manager" & sector == "MAS") %>%
  pull(path)

mas_pan_path <- folder_paths %>%
  filter(approval_type == "panel" & sector == "MAS") %>%
  pull(path)

mas_delegate <- read_mas(mas_del_path)
mas_panel <- read_mas(mas_pan_path) #many folder without project number specified. Need data cleaning help/ or only get file from FY20 (redefine another function)

#merge the text data with features. 
save(mas_delegate, file ="mas_delegate.rda")




#### ----- read in the sector and country data, and the aimm database (later merge to master)
#ops_full <- read_csv("./IFC_disclosure_investment_11.2022.csv") %>%
#  clean_names() 

#ops <- ops_full %>%
#  select(project_number,sector,country_description,region_description) %>%
#  distinct() 

#test <- ops_full %>%
#  select(project_number,sector,country_description,region_description,impact) %>%
#  filter(!is.na(impact))
