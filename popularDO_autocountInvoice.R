library(readr)
library(dplyr)
library(openxlsx)
library(tibble)

options(scipen = 999) #display numbers in non-scientific form

autocountItems <- read.xlsx("autocountItems.xlsx") %>%
  select(2,3) #to match item code and ISBN

b2bPO <- list.files(path="b2bPO", full.names = TRUE) %>% 
  lapply(read_csv, col_types = cols(`ENTRY DATE` = col_date(format = "%Y%m%d"))) %>% 
  bind_rows() %>%
  select(`ENTRY DATE`, PO_NUMBER, `VENDOR DISCOUNT % RATE`, SITE_NAME, SITE_CODE) %>%
  arrange(SITE_CODE) %>%
  rowid_to_column()
  
b2bDO <- list.files(path="b2bDO", full.names = TRUE) %>% 
  lapply(read.csv, encoding="UTF-8") %>% 
  bind_rows()

b2bDO$BARCODE <- as.character(b2bDO$BARCODE)

autocountInvoice <- b2bDO %>%
  left_join(autocountItems, by = c("BARCODE" = "Bar.Code")) %>%
  arrange(SITE_CODE) %>%
  rowid_to_column()
  
autocountPopInvoice <- autocountInvoice %>%
  left_join(b2bPO, by = "PO_NUMBER") %>%
  select(SITE_NAME, Item.Code, DELIVERY_QTY, DO_NUMBER, `ENTRY DATE`, `VENDOR DISCOUNT % RATE`) %>%
  filter(DELIVERY_QTY != "0") %>%
  mutate(`VENDOR DISCOUNT % RATE` = gsub("-", "", `VENDOR DISCOUNT % RATE`),
         `VENDOR DISCOUNT % RATE` = paste(`VENDOR DISCOUNT % RATE`, "%", sep = "")) %>%
  rename("ItemCode" = Item.Code,
         "Qty" = DELIVERY_QTY,
         "Discount" = `VENDOR DISCOUNT % RATE`,
         "YourPONo" = DO_NUMBER,
         "YourPODate" = `ENTRY DATE`)

write.csv(autocountPopInvoice, "autocountPopInvoice.csv", row.names=FALSE)

#Excel turns last digit of PO/DO Number to 0
#use Excel's import data function instead of straight opening the output file