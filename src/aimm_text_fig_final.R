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

#Removing any files from previous run 
file.remove("./List_reading_error.txt")

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

# Extract AIMM narrative
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
      filter(!grepl(".pptx|.xlsx|irm|assistant|is report|is final report|pds|esap|esrs|model|minutes|questions",file_name,ignore.case = TRUE)) %>% #Maybe we need IS report later (content), concept note (abstract)
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
      capture.output(paste("Doc Reading error:",e$message), file = "List_reading_error.txt",append = T)
      tibble(summary = NA, folder_path = file)
    })
  })
  
  final <- flds %>%
    select(id,file_name,full_dir) %>%
    right_join(aimm,c("full_dir" = "folder_path")) 
  
  return(final)
}

final <- read_paths(fldr_paths)

save(final, file = paste0(sector_select,".rda"))

# Extract AIMM indicators -----
## This approach uses docxtractr 
# Extracting the Tables 
indicator <- lapply(unique(final$full_dir), function(file) {
  output <- data.frame(matrix(data=NA,nrow=1,ncol=4)) 
  names(output) <- c("file","project_indicator","market_indicator","reporting_indicator")
  output$file <- file
  tryCatch({
    table_list <- docxtractr::docx_extract_all_tbls(docx = docxtractr::read_docx(file),guess_header = F,preserve = T) #List of all tables in the doc
    #Filtering tables based on dimensions 
    table_list <- table_list %>% 
      map_lgl(~(ncol(.) > 4 & ncol(.) < 9)) %>% 
      magrittr::extract(table_list, .)
    
    for(tbl in table_list){ #Looping through all tables in the doc
      #Remvoing cases of Additionality tables
      if(any(grepl(x = c(unlist(tbl[1,]),unlist(tbl[,1])),pattern = ".*[Aa]dditionality.*"))) next
      if(any(grepl(x = unlist(tbl[1,2:ncol(tbl)]),pattern = ".*Indicator.*|.*Target.*|.*Baseline.*"))){
      #Checking if new format
        if(all(any(grepl(x = unlist(tbl[1:3,4:ncol(tbl)]),pattern = ".*Project.*|.*Market.*|.*Reporting.*")),
           ncol(tbl)>=7)){
          #Extracting the indicators
          indicator_col <- which(grepl(x = unlist(tbl[1,2:4]),pattern = ".*Indicator.*")) %>% ifelse(is_empty( .), 2 , . )
          
          output$project_indicator <- format_delim(tbl[which(grepl(x = unlist(tbl[,5]),pattern = ".*X.*")),indicator_col],delim="; ")
          output$market_indicator <- format_delim(tbl[which(grepl(x = unlist(tbl[,6]),pattern = ".*X.*")),indicator_col],delim="; ")
          output$reporting_indicator <- format_delim(tbl[which(grepl(x = unlist(tbl[,7]),pattern = ".*X.*")),indicator_col],delim="; ")
        }else if(any(grepl(x = unlist(tbl[,1]),pattern = ".*Financial.*|.*Economic.*|.*Private Sector Development.*|.*Project Outcomes.*|.*Market Creation.*|.*Coporate Reporting.*|.*Mandatory Reach.*|.*Strategic Indicators.*"))){
          #check for old format and DOTS
          project_start <- first(which(grepl(x = unlist(tbl[,1]),pattern = ".*Project.*|.*Financial.*")))
          market_start <- first(which(grepl(x = unlist(tbl[,1]),pattern = ".*Market.*|.*Economic.*|.*Private Sector Development.*")))
          reporting_start <- first(which(grepl(x = unlist(tbl[,1]),pattern = ".*Reach.*|.*Reporting.*|.*Environment.*|.*Social.*"))) %>% ifelse(is_empty( .), nrow(tbl)+1 , . )
          indicator_col <- which(grepl(x = unlist(tbl[1,]),pattern = ".*Indicator.*")) %>% ifelse(is_empty( .), 3 , . )
          
          output$project_indicator <- format_delim(tbl[project_start:(market_start-1),indicator_col],delim="; ")
          output$market_indicator <- format_delim(tbl[market_start:(reporting_start-1),indicator_col],delim="; ")
          output$reporting_indicator <- if_else(reporting_start > nrow(tbl),NA,format_delim(tbl[reporting_start:nrow(tbl),indicator_col],delim="; "))
          }
        
      #Exiting the loop
      }
    }
    rm(tbl,table_list)
    return(output)
  },error = function(e) {
    message(sprintf("Error reading file %s: %s", file, e$message))
    capture.output(paste("Indicator reading error:",e$message), file = "List_reading_error.txt",append = T)
    return(output)
  })

}) %>% do.call(rbind.data.frame, . )


# Split the narrative into sections ----
load(file = paste0(sector_select,".rda"))

# data cleaning
final <- final %>%
  group_by(id) %>%     
  mutate(summary_n = n()) %>%
  mutate(summary_text = summary$text,
         summary_style = summary$content_type)%>%
  select(id, full_dir, file_name, summary_n, summary_text, summary_style) %>%
  filter(!grepl("table",summary_style)) %>% # filter out table content or footnote
  filter(!grepl("^table|^\\*", summary_text, ignore.case = TRUE)) %>%
  filter(!is.na(summary_text)) 

final %>% count_id() #228 fig projects

# specify the regular expression for each section
dev_impct <- "^^Development Impact|^Project’s Expected Development Impact|^(?!.*Additionality).*(?i)Development Impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes|^PROJECT IMPACTS"
prjct_outcome <- "^Assessment of Project Outcome(s)?|^Project Outcome|Assessment of Project Outcomes:"
market_creation <- "^(Assessment of )?contribution to market creation|^(?i)(Assessment of Market Creation|^(The )?contribution to market creation|market creation:|^market creation)| Assessment of Contribution to Market Creation –|^Assessment of Market Outcomes| Assessment of Contribution to Market Creation: "

# the functions to help identify the sections
identify_section <- function(file) {
  file %>%
    mutate(summary_text = str_trim(summary_text))%>%
    group_by(id) %>% 
    mutate(aimm_idx = which(str_detect(summary_text,regex("^ADDITIONALITY|^( )?ADDITIONALITY|IFC’s Expected Additionality",ignore_case = TRUE)))[1],
           end_idx = which(str_detect(summary_text,regex("^APPENDIX|^( )?Results Measurement( -)?|^Results Measurement", ignore_case = TRUE)))[1]) %>%
    mutate(section = if_else(row_number() >= end_idx & !is.na(end_idx),
                             "other", "na"),
           section = if_else(row_number() < aimm_idx & !is.na(aimm_idx),
                             "other", section)) %>%
    filter(section != "other") %>%
    mutate(di_idx = which(str_detect(summary_text, regex(dev_impct, ignore_case = TRUE)))[1],
           prjct_idx = which(str_detect(summary_text, regex(prjct_outcome, ignore_case = TRUE)))[1],
           mrkt_idx = which(str_detect(summary_text, regex(market_creation, ignore_case = TRUE)))[1]
    ) 
}

tagging_section <- function(file) {
  file %>%
    filter(di_idx <= prjct_idx,
           prjct_idx < mrkt_idx) %>%
    mutate(section = case_when( row_number() %in% 1:(di_idx-1) ~ "addi",
                                row_number() %in% di_idx:(prjct_idx-1) ~ "summary",
                                row_number() %in% prjct_idx:(mrkt_idx-1) ~ "dev_impct_p",
                                row_number() %in% mrkt_idx:n() ~ "dev_impct_m"
    )
    ) 
}

count_id <- function(file){
  file %>% 
    ungroup() %>%
    select(id) %>% 
    unique() %>% 
    count()
}


# Prioritize the BP, because they followed same structure
bp <- final %>%
  filter(grepl("board|bp",file_name,ignore.case = TRUE)) # select board paper

sections_bp <- bp %>% 
  identify_section()%>%
  mutate(row_n = row_number()) %>%  
  tagging_section()

sections_bp$section %>% table()

bp %>% count_id() #123 projects with board paper
sections_bp %>% count_id() # 120 projects split successfully, 97%

check_bp <- anti_join(bp,sections_bp,"id")%>% 
  identify_section()%>%
  mutate(n_number = row_number()) 

check_bp %>%
  select(aimm_idx,di_idx,prjct_idx,mrkt_idx, end_idx,id) %>%
  unique() # 4 projects have issue splitting different sections

# After splitting the BP, work on projects only have aimm narratives
aimm <- anti_join(final,sections_bp,"id")

sections_aimm <- aimm %>%
  identify_section() %>%
  mutate(row_n = row_number()) %>%  
  tagging_section()

aimm %>% count_id() #110 projects
sections_aimm %>% count_id() #101 projects split successfully

check_aimm <- anti_join(aimm,sections_aimm,"id")%>% 
  identify_section()%>%
  mutate(n_number = row_number()) %>%
  select(aimm_idx,di_idx,prjct_idx,mrkt_idx, end_idx,id) %>%
  unique() # 14 projects need to be checked

# Export file --- 

# the project split by section
final_section <- sections_bp %>%
  rbind(sections_aimm) %>%
  group_by(id,full_dir,file_name,section) %>%
  summarize(text = paste(summary_text,collapse = "\n\n")) 

final_section %>% count_id() # 221 projects

save(final_section, file = paste0(sector_select,"_section.rda"))

AIMM_text <- final_section %>% as.data.frame %>% 
  reshape(data = . , idvar = "full_dir", timevar = "section",v.names  = "text", direction = "wide") %>% #Only 211 unique full_dir path
  merge(x = . ,y = indicator,by.x="full_dir",by.y="file",all.x=T)

save(AIMM_text, file = paste0(sector_select,"_AIMM_text.rda"))

# projects need further checkup 
final_check <- final %>%
  anti_join(final_section, "id") %>%
  identify_section()%>%
  mutate(n_number = row_number())  %>%
  select(aimm_idx,di_idx,prjct_idx,mrkt_idx, end_idx,id, full_dir, file_name) %>%
  unique() 

save(final_check, file = paste0(sector_select,"_section_check.rda")) # 5 projects
write_csv(final_check, file = paste0(sector_select,"_section_check.csv"))

accuracy <- 1- final_check %>% count_id()/final %>% count_id()
accuracy #97.8%
