library(readxl)
library(dplyr)
library(stringr)

stock <- read_excel("stock 02.09.2022.xlsx")

stock <- stock %>%
  select(`Bar Code`, Description, `Desc 2`, `Item Group`, AUTHOR, PUBDATE) %>%
  rename(barCode = `Bar Code`)

LawBooks <- read_excel("lazada.xlsx", sheet = "LawBooks", skip = 1)

LawBooks <- LawBooks %>%
  select(-c(`Product ID`, Brand, Color, Name_CN, catId, 'tr(s-wb-product@md5key)')) %>% 
  mutate('*ISBN/ISSN' = gsub("\\D", "", Model)) %>%
  left_join(stock, by = c('*ISBN/ISSN' = "barCode")) %>%
  mutate(Language = case_when(
    `Item Group` == "MB" ~ "Malay",
    `Item Group` == "EB" ~ "English")) %>%
  mutate(Publisher = `Desc 2`,
         Year = str_extract(PUBDATE, "\\d{4}"),
         Edition = "Regular edition",
         Author = AUTHOR)

write.csv(LawBooks, "LawBooks.csv")

#repeats for all book types
#manually check those with missing or weird ISBN
#key in color