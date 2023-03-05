
#' Summary:
#' This script is aimed to convert pdf board paper to docx.
#' This file should be firstly run through before the aimm_text_fig_final, becaue it will identify the pdf files and convert them to docx, out of simplicity concern, this one only identifies the projects that only contains pdf board paper.
#' So the order was: run aimm_text_fig_final.r first -> get all the posible results -> in this file identify those project only has pdf board paper -> conver them to docx -> run the aimm_text_fig_final.r again. 

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


# find out bp in pdf ----
get_paths_pdf <- function(fldr_paths){
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
      filter(!grepl(".pptx|.xlsx|irm|assistant|Assistance|Assisstant|Assitant|Assist|approach|Decision|AIMM Rating|memo|report|is final report|pds|esap|esrs|model|minutes|questions",file_name,ignore.case = TRUE)) %>% 
      mutate(id = str_extract(dr, "\\d{5}")) %>%
      mutate(full_dir = paste0(fldr_path,"/",dr))%>%
      filter(grepl(".pdf$",dr))
    #there's pdf potentially can be used.
  }
  
  flds <- lapply(fldr_paths, get_path) %>%
    do.call(rbind,.)
}

flds_pdf <- get_paths_pdf(fldr_paths) %>%
  mutate(id = as.double(id))
  
load(file = paste0(sector_select,"_AIMM_text.rda")) # this is the file firstly run by aimm_text_fig_final.r got the projects with docx board paper. 

pdf_check <- anti_join(flds_pdf,AIMM_text,"id") %>%
  filter(!grepl("FY18",full_dir)) %>% # board paper in FY18 is not applicable
  mutate(full_dir_docx = sub("\\.pdf$", ".docx", full_dir))

write.csv(pdf_check, file = paste0(sector_select,"_pdf.csv")) # get the list of files converted to pdf. 

# read the pdf paths --- (this is just text, should be deleted)
pdf_read <- pdf_check %>%
  select(-full_dir) %>%
  rename(full_dir = full_dir_docx)

read_paths_pdf <- function(flds){
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

final_pdf <- read_paths_pdf(pdf_read)

save(final_pdf,file = paste0(sector_select,"final_pdf.rda"))