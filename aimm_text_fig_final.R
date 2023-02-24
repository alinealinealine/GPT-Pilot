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

# Read folder location address from file -----
folder_paths <- read.csv(file = "./file_location.csv") %>%
  mutate(path = gsub("\\\\", "/", path)) #' The file has location for both delegated and panel approved projects

if (grepl("xweng", getwd(),ignore.case = TRUE)) {
  folder_paths <- folder_paths %>%
    mutate(path = gsub("gjain5","xweng",path))
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
      capture.output(e$message, file = "List_reading_error.txt",append = T)
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
indicator <- lapply(unique(flds$full_dir), function(file) {
  output <- data.frame(matrix(data=NA,nrow=1,ncol=4)) 
  names(output) <- c("file","project_indicator","market_indicator","reporting_indicator")
  output$file <- file
  tryCatch({
    table_list <- docxtractr::docx_extract_all_tbls(docx = docxtractr::read_docx(file),guess_header = F,preserve = T) #List of all tables in the doc
    for(tbl in table_list){ #Looping through all tables in the doc
      #Check for DOTS template
      if(any(grepl(x = unlist(tbl[,1]),pattern = ".*Financial.*|.*Economic.*|.*Private Sector Development.*|.*Performance.*|.*Environment.*|.*Social.*"))){ 
        output$project_indicator <- "DOTS TEMPLATE"
        output$market_indicator <- "DOTS Template"
        output$reporting_indicator <- "DOTS Tempalte" 
     }else{ #check for old format
      print("identified old format table")
      if(any(grepl(x = unlist(tbl[,1]),pattern = ".*Project Outcomes.*|.*Market Creation.*|.*Coporate Reporting.*|.*Mandatory Reach.*|.*Strategic Indicators.*"))){
        output$project_indicator <- format_delim(tbl[which(grepl(x = unlist(tbl[,1]),pattern = ".*Project.*")):(which(grepl(x = unlist(tbl[,1]),pattern = ".*Market.*"))-1),3],delim="; ")
        output$market_indicator <- format_delim(tbl[which(grepl(x = unlist(tbl[,1]),pattern = ".*Market.*")):(which(grepl(x = unlist(tbl[,1]),pattern = ".*Reach.*|.*Reporting.*"))-1),3],delim=", ")
        output$reporting_indicator <- format_delim(tbl[which(grepl(x = unlist(tbl[,1]),pattern = ".*Reach.*|.*Reporting.*")):nrow(tbl),3],delim=", ")
      }else{
        #Checking if new format
      if(ncol(tbl)>=7){ #First checking for a table with more than 7 columns - as AIMM table should have 7/8 columns
        print("identified possible format table")
        # if(any(all(tbl[1,1:4] == c("Description of Indicator", "Indicator", "Baseline", "Target")), #Then checking first level of headers - these might match with other indicator tables as well (if they are in annex etc.)
        #     all(tbl[2,5:7]==c("Project", "Market", "Reporting")))){ #Final check with second level of headers - these are unlikely to repeat in this order and location across any other table
        if(any(unlist(tbl[1:2,]) %in% c("Project", "Market", "Reporting"))){
          print("confirmed new format table")
          #Extracting the indicators
          output$project_indicator <- format_delim(tbl[which(tbl[,5]=="X"),2],delim="; ")
          output$market_indicator <- format_delim(tbl[which(tbl[,6]=="X"),2],delim=", ")
          output$reporting_indicator <- format_delim(tbl[which(tbl[,7]=="X"),2],delim=", ")
        }
      }
      }
      #Exiting the loop
    }}
    rm(tbl,table_list)
    return(output)
  },error = function(e) {
    message(sprintf("Error reading file %s: %s", file, e$message))
    capture.output(e$message, file = "List_reading_error.txt",append = T)
    return(output)
  })

})
indicator <- do.call(rbind.data.frame,indicator)
temp <- final %>%
  group_by(id) %>%     
  mutate(summary_n = n()) %>%
  mutate(summary_text = summary$text,
         summary_style = summary$content_type)



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

final %>% count_id() #227 fig projects

# specify the regular expression for each section
 dev_impct <- "^^Development Impact|^Project’s Expected Development Impact|^(?!.*Additionality).*(?i)Development Impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes |^PROJECT IMPACTS"
 prjct_outcome <- "^Assessment of Project Outcomes"
 market_creation <- "(Assessment of )?contribution to market creation|^(?i)(Assessment of Market Creation|(The )?contribution to market creation|market creation:|market creation)"

# the functions to help identify the sections
identify_section <- function(file) {
  file %>%
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
    filter(di_idx < prjct_idx,
           prjct_idx < mrkt_idx) %>%
    mutate(section = case_when( row_number() %in% 1:(di_idx-1) ~ "addi",
                                row_number() %in% di_idx:(prjct_idx-1) ~ "summary",
                                row_number() %in% prjct_idx:mrkt_idx ~ "dev_impct_p",
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
sections_bp %>% count_id() # 117 projects split successfully, 95%

check_bp <- anti_join(bp,sections_bp,"id")%>% 
  identify_section()%>%
  mutate(n_number = row_number()) %>%
  select(aimm_idx,di_idx,prjct_idx,mrkt_idx, end_idx,id) %>%
  unique() # 6 projects have issue splitting different sections

# After splitting the BP, work on projects only have aimm narratives
aimm <- anti_join(final,sections_bp,"id")

sections_aimm <- aimm %>%
  identify_section() %>%
  mutate(row_n = row_number()) %>%  
  tagging_section()

aimm %>% count_id() #150 projects
sections_aimm %>% count_id() #135 projects split successfully, 90%

check_aimm <- anti_join(aimm,sections_aimm,"id")%>% 
  identify_section()%>%
  mutate(n_number = row_number()) %>%
  select(aimm_idx,di_idx,prjct_idx,mrkt_idx, end_idx,id) %>%
  unique() # 14 projects need to be checked

# Export file -----

# the project split by section
final_section <- sections_bp %>%
  rbind(sections_aimm) %>%
  group_by(id,full_dir,file_name,section) %>%
  summarize(text = paste(summary_text,collapse = "\n\n")) 

final_section %>% count_id() # 213 projects

save(final_section, file = paste0(sector_select,"_section.rda"))

# projects need further checkup 
final_check <- final %>%
  anti_join(final_section, "id") %>%
  identify_section()%>%
  mutate(n_number = row_number()) %>%
  select(aimm_idx,di_idx,prjct_idx,mrkt_idx, end_idx,id, full_dir, file_name) %>%
  unique() 

save(final_check, file = paste0(sector_select,"_section_check.rda"))

accuracy <- 1- final_check %>% count_id()/final %>% count_id()
accuracy #94.3%
