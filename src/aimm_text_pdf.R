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
library(docxtractr)
library(readr)
library(stringr)
library(pdftools)

# Read folder location address from file -----
folder_paths <- read.csv(file = "./file_location.csv") %>%
  mutate(path = gsub("\\\\", "/", path)) %>%
  mutate(path = gsub("C:/Users/gjain5/WBG/Ahmad Famm Alkhuzam - CSE AIMM/Projects Documents/","C:/Users/XWeng/WBG/AIMM Repository - Projects Documents/",path)) #' The file has location for both delegated and panel approved projects

if (grepl("gjain5", getwd(),ignore.case = TRUE)) {
  folder_paths <- folder_paths %>%
    mutate(path = gsub("xweng","gjain5", path))
} # If "xweng" is using this file, replace path "gjain5" with "xweng"

#' Considering the variation in how different sector store their information, it might be wise to split this into sections.
#' Currently doing conditions based on sector only but possible to nuance further based on approval type. 

sector_select <- "FIG"

# Extract the narrative -----

# Define the folder paths
del_path <- folder_paths %>%
  filter(approval_type == "manager" & sector == sector_select) %>%
  pull(path)

pan_path <- folder_paths %>%
  filter(approval_type == "panel" & sector == sector_select) %>%
  pull(path)

fldr_paths <- c(del_path,pan_path)


# identify cases with pdf as project doc. 
load(file = paste0(sector_select,"_section.rda"))
get_path_pdf <- function(fldr_path){list.files(path = fldr_path ,recursive = TRUE) %>%
    as.data.frame() %>%
    rename(file_name = '.') %>%
    mutate(dr = file_name) %>% #use this as directory
    mutate(file_name = sub(".*/", "", dr)) %>% #get the file name when there are multiple layers of files
    mutate(fld_name = sapply(dr, function(x) {
      parts <- strsplit(x, "/")[[1]]
      part <- parts[grep("\\b[0-9]{5}\\b", parts)][1]
      part
    })) %>%
    mutate(fld_name = case_when( grepl("Amaggi Cotton", file_name) ~ "Amaggi Cotton(43740)",  
                                 grepl( "Rupshi Foods final document", file_name) ~ "Rupshi Foods final document (46329)",
                                 #correct cases where the project number is not specified in the folder name.(above are just fixation for delegated_mas files)
                                 TRUE ~ fld_name)) %>%
    mutate(id = ifelse(grepl("\\b\\d{5}\\b", fld_name), 
                       as.numeric(gsub("[^[:digit:]]", "", regmatches(fld_name, regexpr("\\b\\d{5}\\b", fld_name)))), 
                       NA)) %>%
    filter(!grepl(".pptx|.xlsx|irm|assistant|Assisstant|assistance|is report|is final report|pds|esap|esrs|model|minutes|questions",file_name,ignore.case = TRUE)) %>% #Maybe we need IS report later (content), concept note (abstract)
    filter(grepl("board paper|narrative|aimm",file_name,ignore.case = TRUE)) %>% #filter out AIMM narrative documents.
    mutate(full_dir = paste0(fldr_path,"/",dr))%>%
    filter(!grepl(".docx$",dr)) #there's pdf potentially can be used.
}

flds_pdf <- lapply(fldr_paths, get_path_pdf) %>%
  do.call(rbind,.) %>%
  group_by(id)

flds_pdf %>% count_id # 29 projects uses pdf board paper

flds_pdf <- anti_join(flds_pdf,final_section,"id")  
flds_pdf %>% count_id # 16 projects only have pdf board paper/AIMM narrative


save(flds_pdf, file = paste0(sector_select,"_pdf_check.rda")) 
write_csv(flds_pdf, file = paste0(sector_select,"_pdf_check.csv"))


# convert pdf to docx
convert_pdf_to_docx <- function(file_path) {
  doc_path <- gsub(".pdf$", ".docx", file_path)
  pdf_convert(file_path, doc_path)
  return(doc_path)
}

pdf_file <- flds_pdf$full_dir[2]
docx_file <- "C:/Users/XWeng/OneDrive - WBG/Aline_Project/GPT_AIMM/aimm_text_extract/Git_folder/GPT-Pilot/Coverfox Approved Board Paper.docx"

test <- purrr::map_dfr(docx_file, function(file) {
  tryCatch({
    doc <- officer::read_docx(file)
    summary <- officer::docx_summary(doc)
    tibble(summary = summary, folder_path = file)
  }, error = function(e) {
    message(sprintf("Error reading file %s: %s", file, e$message))
    capture.output(e$message, file = "List_reading_error.txt",append = T)
    tibble(summary = NA, folder_path = file)
  })
}) %>%
  mutate(summary_n = n()) %>%
  mutate(summary_text = summary$text,
         summary_style = summary$content_type)%>%
  mutate(id = 1) %>%
  identify_section() %>%
  tagging_section()




