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


# Read folder location address from file -----
folder_paths <- read.csv(file = "./file_location.csv") %>%
  mutate(path = gsub("\\\\", "/", path)) #' The file has location for both delegated and panel approved projects

if (grepl("xweng", getwd(),ignore.case = TRUE)) {
  folder_paths <- folder_paths %>%
    mutate(path = gsub("gjain5","xweng",path))
} # If "xweng" is using this file, replace path "gjain5" with "xweng"

#' Considering the variation in how different sector store their information, it might be wise to split this into sections.
#' Currently doing conditions based on sector only but possible to nuance further based on approval type. 

# Extract the narrative -----
read_paths <- function(fldr_paths){
  # identify the files and needed docs. (write a function )
  get_path <- function(fldr_path){list.files(path = fldr_path ,recursive = TRUE) %>%
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
      filter(!grepl(".pptx|.xlsx|irm|assistant|is report|is final report|pds|esap|esrs|model",file_name,ignore.case = TRUE)) %>% #Maybe we need IS report later (content), concept note (abstract)
      filter(grepl("board paper|narrative|aimm",file_name,ignore.case = TRUE)) %>% #filter out AIMM narrative documents.
      mutate(full_dir = paste0(fldr_path,"/",dr))%>%
      filter(grepl(".docx$",dr)) #there's pdf potentially can be used.
  }
  flds <- lapply(fldr_paths, get_path) %>%
    do.call(rbind,.)
  
  # use the directory to extract the text. 
  file_list <- flds$full_dir
  
  aimm <- purrr::map_dfr(flds$full_dir, function(file) {
    tryCatch({
      doc <- officer::read_docx(file)
      summary <- officer::docx_summary(doc)
      tibble(summary = summary, folder_path = file)
    }, error = function(e) {
      message(sprintf("Error reading file %s: %s", file, e$message))
      tibble(summary = NA, folder_path = file)
    })
  })
  
  final <- flds %>%
    select(id,file_name,full_dir) %>%
    right_join(aimm,c("full_dir" = "folder_path")) 
  
  final
}

sector_select <- "CDF"

# Define the folder paths
del_path <- folder_paths %>%
  filter(approval_type == "manager" & sector == sector_select) %>%
  pull(path)

pan_path <- folder_paths %>%
  filter(approval_type == "panel" & sector == sector_select) %>%
  pull(path)

fldr_paths <- c(del_path,pan_path)

# Extract AIMM narrative
final <- read_paths(fldr_paths)

# Split the narrative into sections
section_split <-function(file){
  sections <- file %>%
    group_by(id) %>%     
    mutate(summary_n = n()) %>%
    filter(summary_n > 100) %>% # earlier assessments are too short. 
    filter(!grepl("board|questions",file_name,ignore.case = TRUE)) %>% #filter out not relevant doc
    mutate(summary_text = summary$text)  %>%
    group_by(id) %>% 
    mutate(start_idx = which(str_detect(summary_text, regex("^(?i)development impact|(?i)development impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes |PROJECT IMPACTS", ignore_case = TRUE)))[1]) %>%
    mutate(end_idx = which(str_detect(summary_text, regex("(The )?Assessment of Contribution to Market Creation|Market (creation|outcome) | ^(?i)Assessment of Market Outcomes | Market Creation – |market impact", ignore_case = TRUE)))[1]) %>%
    filter(!is.na(start_idx)) %>% 
    mutate(section = case_when( row_number() %in% 1:start_idx ~ "addi",
                                row_number() %in% start_idx:(end_idx-1) ~ "dev_impct_p",
                                row_number() %in% end_idx:n() ~ "dev_impct_m")
    )
  sections
}

final_section <- section_split(final)

# Export file
save(final_section, file = paste0(sector_select,".rda"))