library(openxlsx)
library(dplyr)
library(tidyr)
library(tibble)

options(scipen = 999) #display numbers in non-scientific form

b2b <- read.csv("CreditNoteBatch_Popular_20221009125452.csv")

autocount <- read.xlsx("autocountCN.xlsx") %>%
  {.[-1]} #remove first column

b2b$Return.Purchase.Order.No <- as.character(b2b$Return.Purchase.Order.No)

merged <- b2b %>%
  left_join(autocount, by = c("Return.Purchase.Order.No" = "Remark.1")) %>%
  separate(Branch.Code, into = c("Popular", "Branch", "Number"), sep = " ") %>%
  mutate(CN.Date = POD.Upload.Date,
         Doc..No. = gsub("-", "", Doc..No.))
    
merged$CN.No <- apply(merged[, c("Branch", "Number", "Doc..No.")], 1, function(i){ paste(na.omit(i), collapse = "")}) #merge columns to create CN No. (ignoring NA)

merged2 <- merged %>% 
  mutate(CN.No = gsub("-", "", CN.No,)) %>% #remove "-"
  select(Vendor.No, CN.No, Return.Purchase.Order.No, CN.Date, Store.Code) %>%
  add_column("Import Format Indicator" = "HEADER", .before = 1) %>%
  rename("Vendor Code" = Vendor.No,
         "Credit Note Number" = CN.No,
         "RPO Number" = Return.Purchase.Order.No,
         "Credit Note Date" = CN.Date,
         "Store Code" = Store.Code)
  
write.csv(merged2, "b2bImportCN.csv", row.names = FALSE)

#remove empty rows and duplicates in Excel