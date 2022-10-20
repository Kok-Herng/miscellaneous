library(openxlsx)
library(tidyr)
library(dplyr)
library(tibble)

options(scipen = 999) #display numbers in non-scientific form

#required to connect GR No. and DO No. (DO No. required to find corresponding invoice No. in autocount)
b2bGRN <- list.files(path="b2bGRNs", full.names = TRUE) %>% 
  lapply(read.csv, encoding="UTF-8") %>% 
  bind_rows %>% #merge GRN files
  rename(GR.No = GR_NUMBER)

#invoices to be worked on  
b2bInvoice <- list.files(path="b2bInvoices", full.names = TRUE) %>% 
  lapply(read.csv, encoding="UTF-8") %>% 
  bind_rows %>% #merge Invoice files
  mutate(Invoice.Date = GR.Date) %>%
  group_split(GR.No) #group according to GR.No

b2b <- lapply(b2bInvoice, left_join, b2bGRN)

#invoices extracted from autocount
autocount <- read.xlsx("autocountInvoice.xlsx", detectDates = T) %>%
  separate(Branch.Code, into = c("Popular", "Invoice.No", "Number"), sep = " ") %>%
  rename(DO_NUMBER = Ref..Doc..No.)

autocount$Invoice.No <- apply(autocount[, c("Invoice.No", "Number")], 1, function(i){ paste(na.omit(i), collapse = "-")}) #merge columns to create branch code

autocount$DO_NUMBER <- as.numeric(autocount$DO_NUMBER)

mergedList <- lapply(b2b, left_join, autocount, by = "DO_NUMBER") #connect b2b and autocount invoices using DO No. 

merged <- mergedList %>%
  lapply(function(x) x[!duplicated(x$BARCODE),]) %>% #select only unique barcode
  lapply(mutate, 
         Doc..No. = gsub("-", "", Doc..No.,), #remove "-"
         Invoice.No.y = paste(Invoice.No.y, Doc..No., sep = "-"), #branch name followed by invoice number
         Tally = Total.Amount == Net.Total) #to check the total amount

mergedDf <- merged %>%
  bind_rows() %>%
  select(Vendor.No, Invoice.No.y, GR.No, Invoice.Date, Store.Code, Tally) %>%
  add_column("Import Format Indicator" = "HEADER", .before = 1) %>%
  rename("Vendor Code" = Vendor.No,
         "Invoice Number" = Invoice.No.y,
         "GR Number" = GR.No,
         "Invoice Date" = Invoice.Date,
         "Store Code" = Store.Code)

write.csv(mergedDf, "b2bInvoiceImport.csv", row.names = FALSE)

#IKC renamed to IKC-II
#remove rows with NA, and duplicates in Excel