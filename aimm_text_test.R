
#Load necessary files
library(tidytext)
library(tidyverse)
library(tidyr)
library(purrr)
library(readxl)
library(janitor)
library(readtext)
library(officer)

# this script to identify the sections: additionality, development impact (project, market), impact measurement table. 

load(file = 'fig.rda')

select_id <- sample(unique(final$id),5)


# desribe the content
final %>% # the length of the document
  group_by(id) %>%     
  mutate(summary_n = n()) %>%
  ungroup() %>%
  summary(summary_n)

final %>%
  group_by(id) %>%     
  mutate(summary_n = n()) %>%
  filter(grepl("board paper|bp",file_name,ignore.case = TRUE)) %>%
  summary(summary_n) # size of board paper: min: 616, 3rd: 1436

final %>%
  group_by(id) %>%     
  mutate(summary_n = n()) %>%
  filter(!grepl("board paper|question",file_name,ignore.case = TRUE)) %>%
  summary(summary_n) # size of non- board paper min: 109, 3rd: 745

final %>% 
  group_by(id) %>%     
  mutate(summary_n = n()) %>%
  filter(!grepl("board paper|question|minutes",file_name,ignore.case = TRUE)) %>%
  filter(summary_n > 1400) %>%
  ungroup() %>%
  select(file_name) %>%
  unique() %>%
  View()  # the project larger than 1,400 are still aimm

final %>%
  filter(grepl("board|question|minute",file_name,ignore.case = TRUE)) %>%
  select(file_name) %>%
  unique() %>% count() #122 projects with board papers,question,or minute in doc. 

final %>% count_id() #228 projects total scripted


# identify sections

  relevant <- final %>%
    #group_by(id) %>%     
    mutate(summary_n = n()) %>%
    mutate(summary_text = summary$text,
           summary_style = summary$content_type)%>%
    select(id, full_dir, file_name, summary_n, summary_text, summary_style) %>%
    # filter out table content or footnote
    filter(!grepl("table",summary_style)) %>%
    filter(!grepl("^table|^\\*", summary_text, ignore.case = TRUE)) %>%
    filter(!is.na(summary_text)) %>%
    filter(!grepl("board|questions|minute|bp",file_name,ignore.case = TRUE)) #filter out board paper or other not relevant documents
  
  relevant %>% count_id() #133 projects with relevant doc
  
  sections <- relevant %>% # split the relevant doc
    group_by(id) %>% 
    mutate(start_idx = which(str_detect(summary_text, regex("^(?!.*Additionality).*(?i)Development Impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes |PROJECT IMPACTS", ignore_case = TRUE)))[1]) %>%
    mutate(end_idx = which(str_detect(summary_text, regex("(Assessment of )?contribution to market creation|^(?i)(Assessment of Market Creation|(The )?contribution to market creation|market creation:|market creation)", ignore_case = TRUE)))[1]) %>%
    mutate(drop_idx = which(str_detect(summary_text,regex("^APPENDIX", ignore_case = TRUE)))[1]) %>%
    filter(!is.na(start_idx) & !is.na(end_idx)) %>% 
    filter(start_idx < end_idx) %>%
    mutate(section = case_when( row_number() %in% 1:start_idx ~ "addi",
                                row_number() %in% start_idx:(end_idx-1) ~ "dev_impct_p",
                                row_number() %in% end_idx:n() ~ "dev_impct_m")
    )  %>%
    mutate(section = if_else(row_number() >= drop_idx & !is.na(drop_idx),
                             "appendix",
                             section))%>%
    mutate(row_n = row_number()) 
  
  sections %>% count_id() #127 projects can be split
  sections$section %>% table()
  
# Check 1: Identify doc that is relevant but can't be split into sections
  check <- anti_join(relevant,sections,by ="id") %>% 
    group_by(id) %>% 
    mutate(start_idx = which(str_detect(summary_text, regex("^(?!.*Additionality).*(?i)Development Impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes |PROJECT IMPACTS", ignore_case = TRUE)))[1]) %>%
    mutate(end_idx = which(str_detect(summary_text, regex("(Assessment of )?contribution to market creation|^(?i)(Assessment of Market Creation|(The )?contribution to market creation|market creation:|market creation)", ignore_case = TRUE)))[1]) %>%
    mutate(drop_idx = which(str_detect(summary_text,regex("^APPENDIX", ignore_case = TRUE)))[1]) 
  
  check %>% count_id() # 6 projects
  
  check_summ <- check %>% #see which are the 6 projects and the issue
    ungroup() %>%
    select(start_idx,end_idx,drop_idx,file_name) %>% 
    unique() %>%
    mutate(end_start = end_idx - start_idx )
  
# Check 2: For those identified as not relevant but these doc are the only available doc for these projects. 
  
  other_doc <- anti_join(final,relevant,by= "id") %>%
    select(id,file_name)%>%
    unique()
  View(other_doc)
  count_id(other_doc)# 95 projects
  
# Design the method to identify sections in these board paper
  bp <- other_doc %>%
    select(id) %>%
    left_join(final) %>%
    #group_by(id) %>%     
    mutate(summary_n = n()) %>%
    mutate(summary_text = summary$text,
           summary_style = summary$content_type)%>%
    select(id, full_dir, file_name, summary_n, summary_text, summary_style) %>%
    # filter out table content or footnote
    filter(!grepl("table",summary_style)) %>%
    filter(!grepl("^table|^\\*", summary_text, ignore.case = TRUE)) %>%
    filter(!is.na(summary_text)) %>%
    filter(grepl("board|bp",file_name,ignore.case = TRUE)) %>%
    group_by(id)
  
  bp %>% count_id()  #92 projects
  
  sections_bp <- bp %>% 
    group_by(id) %>% 
    mutate(aimm_idx = which(str_detect(summary_text,regex("^ADDITIONALITY",ignore_case = TRUE)))[1]) %>% 
    mutate(start_idx = which(str_detect(summary_text, regex("^(?!.*Additionality).*(?i)Development Impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes |PROJECT IMPACTS", ignore_case = TRUE)))[1]) %>%
    mutate(end_idx = which(str_detect(summary_text, regex("(Assessment of )?contribution to market creation|^(?i)(Assessment of Market Creation|(The )?contribution to market creation|market creation:|market creation)", ignore_case = TRUE)))[1]) %>%
    mutate(drop_idx = which(str_detect(summary_text,regex("^APPENDIX|^Results Measurement", ignore_case = TRUE)))[1]) %>%
    mutate(row_n = row_number()) %>%
    filter(!is.na(start_idx) & !is.na(end_idx)) %>% 
    filter(start_idx < end_idx) %>%
    mutate(section = if_else(row_number() >= drop_idx & !is.na(drop_idx),
                             "other",
                             "na"),
           section = if_else(row_number() < aimm_idx & !is.na(aimm_idx),
                             "other",
                             "na"))%>%
    filter(section != "other") %>%
    mutate(section = case_when( row_number() %in% 1:start_idx ~ "addi",
                                row_number() %in% start_idx:(end_idx-1) ~ "dev_impct_p",
                                row_number() %in% end_idx:n() ~ "dev_impct_m")
    )

  sections_bp %>% count_id() #120 projects. 
  
  #the only 2 bp that can't be identified:
  check_bp <- anti_join(bp,sections_bp, by= "id")%>% 
    mutate(aimm_idx = which(str_detect(summary_text,regex("^ADDITIONALITY",ignore_case = TRUE)))[1]) %>% 
    mutate(start_idx = which(str_detect(summary_text, regex("^(?!.*Additionality).*(?i)Development Impact$|The Project has an Anticipated Impact Measurement|Assessment of Project Outcomes |PROJECT IMPACTS", ignore_case = TRUE)))[1]) %>%
    mutate(end_idx = which(str_detect(summary_text, regex("(Assessment of )?contribution to market creation|^(?i)(Assessment of Market Creation|(The )?contribution to market creation|market creation:|market creation)", ignore_case = TRUE)))[1]) %>%
    mutate(drop_idx = which(str_detect(summary_text,regex("^APPENDIX|^Results Measurement", ignore_case = TRUE)))[1]) 
  
  bp %>% count_id()
  sections_bp %>% count_id()
  check_bp %>% count_id() # 2 projects
  
  check_summ_bp <- check_bp %>% #see which are the projects and the issue
    ungroup() %>%
    select(start_idx,end_idx,drop_idx,file_name,id) %>% 
    unique() %>%
    mutate(end_start = end_idx - start_idx )
