library(tidytext)
library(tidyverse)
library(tidyr)
library(purrr)
library(readxl)
library(janitor)
library(xlsx)


# This file is to extract the aimm narrative into text. 
cd_raw <- "C:/Users/XWeng/OneDrive - WBG/Aline_Project/GPT_AIMM/Narrative Data/Raw"
cd_raw_mas <- paste0(cd_raw,"/MAS_delegate")
cd_raw_fig <- paste0(cd_raw,"/FIG_delegate")

# delegated projects ----

# identify the files and needed docs. (write a function )
flds <- list.files(cd_raw_mas,recursive = TRUE) %>%
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
  mutate(full_dir = paste0(cd_raw,"/",dr))%>%
  filter(grepl(".docx$",dr)) #there's pdf potentially can be used. 

# use the directory to extract the text. 
aimm_mas <- map(flds$full_dir[c(1:54,56:90,92:101)],readtext::readtext) %>% 
  bind_rows() %>%
  mutate(word_c = str_count(text, "\\b\\w+\\b")) #count the number of words

 #there's warning on 55: IBS III (45637)/IBS_Sonagrin_Additionality_AIMM_Assessment_draft.docx (because it's encrypted)
 #there's warning on 91: Sunshine (43150)/Board paper Sunshine (#43150) 19Nov20 v2[2305843009214489141].docx

# read in the sector and country data 
ops_full <- read_csv("IFC_disclosure_investment_11.2022.csv") %>%
  clean_names() 

ops <- ops_full %>%
  select(project_number,sector,country_description,region_description) %>%
  distinct() 

test <- ops_full %>%
  select(project_number,sector,country_description,region_description,impact) %>%
  filter(!is.na(impact))

#merge the text data with features. 
mas_delegate <- flds %>%
  select(id,file,full_dir) %>%
  right_join(aimm_mas,c("file" = "doc_id")) %>%
  left_join(ops,c("id" = "project_number")) %>%
  select(id, file, region_description, country_description, sector, word_c, full_dir, text)

save(mas_delegate, file ="mas_delegate.rda")
