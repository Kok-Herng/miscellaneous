library(readxl)
library(dplyr)
library(janitor)
library(openxlsx)

autocountItems <- read_excel("autocountItems.xlsx") %>%
  select(2,3) #to match item code and ISBN

#main data to work on
shopeeOrderSummary <- read_excel("Income.released.20220713_20221013.xls",col_types = c("numeric", "text", "text", "text", "text", "text", "text", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "text", "numeric", "numeric", "skip", "numeric", "text", "text"), skip = 5) %>%
  select(2,5,8:21) #does not contain individual items for each order, only shows  total amount

shopeeOrderDetails <- list.files(path="allOrders", full.names = TRUE) %>% 
  lapply(read_excel) %>% 
  bind_rows %>%
  select(1,2,12:16, 18) %>% #contains individual items for each order
  mutate(`SKU Reference No.` = gsub("978967594225-9", "9789675942259", `SKU Reference No.`),
         `SKU Reference No.` = gsub("978967311944", "9789670311944", `SKU Reference No.`)) #rectify wrongly formatted ISBN
 
shopeeOrder <- shopeeOrderSummary %>%
  left_join(shopeeOrderDetails) %>%
  filter(`Order Status` == "Completed") %>% #only work on completed orders (ie payment released)
  left_join(autocountItems, by = c("SKU Reference No." = "Bar Code")) %>%
  mutate("Discount" = `Your Seller product promotion` + `Refund Amount to Buyer (RM)` + `Product Discount Rebate from Shopee` + Voucher + `Seller Absorbed Coin Cashback` + `Shipping Fee Paid by Buyer` + `Shipping Rebate From Shopee` + `Actual Shipping Fee` + `Reverse Shipping Fee` + `Commission Fee (incl. SST)` + `Service Fee (Incl. SST)` + `Transaction Fee (Incl. SST)`, #calculate total discount
         "Fees & Charges" = `Shipping Fee Paid by Buyer` + `Shipping Rebate From Shopee` +`Actual Shipping Fee` + `Commission Fee (incl. SST)` + `Service Fee (Incl. SST)` + `Transaction Fee (Incl. SST)`,
         "Nett" = `Original product price` + Discount, #calculate nett income (add because discount is minus)
         "Tally" = near(`Original product price` + Discount, `Total Released Amount (RM)`)) #check if amount is tally

shopeeInvoice <- shopeeOrder %>%
  select(`Order Creation Date`, `Order ID`, `Product Name`, Quantity, `Original Price`, `Item Code`, `SKU Reference No.`, `Variation Name`, `Original product price`, Discount, `Fees & Charges`, Nett, `Total Released Amount (RM)`, Tally) %>%
  group_by(`Order ID`)

shopeeInvoiceNoFreeBook <- shopeeInvoice %>% #orders that do not contain "New Road Map To A Developed Nation by Koon Yew Yin" (free book)
  filter(all(`SKU Reference No.` != "9789675832437")) %>%
  mutate("Order Item Quantity" = n(),
         #Discount and `Vouchers & Discounts` & `Fees & Charges` sections cannot be run simultaneously 
         Discount = Discount/`Order Item Quantity`,
         Discount = format(round(Discount, 2), nsmall = 2),
         Discount = as.numeric(Discount),
         
         `Vouchers & Discounts` = Discount - `Fees & Charges`,
         `Vouchers & Discounts` = `Vouchers & Discounts`/`Order Item Quantity`, #split amount depends on no. of items in order
         `Vouchers & Discounts` = format(round(`Vouchers & Discounts`, 2), nsmall = 2), #format as 2 decimal places
         `Fees & Charges` = `Fees & Charges`/`Order Item Quantity`,
         `Fees & Charges` = format(round(`Fees & Charges`, 2), nsmall = 2)) 

shopeeInvoiceFreebook <- shopeeInvoice %>% #orders that contain road map book
  filter(any(`SKU Reference No.` == "9789675832437")) %>%
  mutate(Discount = Discount + 10, #overwrite road map book's price since its free (+ because the discount is negative)
         "Order Item Quantity" = n() - 1, #overwrite road map book
         #Discount and `Vouchers & Discounts` & `Fees & Charges` sections cannot be run simultaneously 
         Discount = Discount/`Order Item Quantity`,
         Discount = format(round(Discount, 2), nsmall = 2),
         Discount = as.numeric(Discount),
         
         `Vouchers & Discounts` = Discount - `Fees & Charges`,
         `Vouchers & Discounts` = `Vouchers & Discounts`/`Order Item Quantity`,
         `Vouchers & Discounts` = format(round(`Vouchers & Discounts`, 2), nsmall = 2),
         `Fees & Charges` = `Fees & Charges`/`Order Item Quantity`,
         `Fees & Charges` = format(round(`Fees & Charges`, 2), nsmall = 2))

autocountInvoice <- shopeeInvoiceNoFreeBook %>%
  rbind(shopeeInvoiceFreebook) %>% #combine both type of orders
  ungroup() %>%
  mutate(`Item Code` = ifelse(grepl("bookmark", `Product Name`, ignore.case = TRUE), "BMARK", `Item Code`), #input item code for bookmark
         `Item Code` = ifelse(grepl("POLITIKO", `Product Name`, ignore.case = TRUE), "PO34", `Item Code`)) %>% #input item code for politiko
  filter(`SKU Reference No.` != "9789675832437") %>% #remove road map book
  group_split(`Order Creation Date`) %>% #group by order date
  lapply(mutate, 
         Discount = gsub("-", "", Discount),
         `Vouchers & Discounts` = gsub("-", "", `Vouchers & Discounts`),
         `Fees & Charges` = gsub("-", "", `Fees & Charges`)) %>% #remove "-"
  lapply(select, c(`Item Code`, Quantity, `Original Price`, Discount, `Order ID`, `Order Creation Date`, `Product Name`, `SKU Reference No.`, `Vouchers & Discounts`, `Fees & Charges`)) %>%
  lapply(rename, 
         "ItemCode" = `Item Code`,
         "Qty" = Quantity,
         "UnitPrice" = `Original Price`,
         "YourPONo" = `Order ID`,
         "YourPODate" = `Order Creation Date`,
         "Description" = `Product Name`,
         "ISBN" = `SKU Reference No.`)

write.xlsx(autocountInvoice, paste("shopeeInvoiceAsOf", format(Sys.Date(), "%Y-%m-%d"), ".xlsx", sep = "")) #output to Excel file including output date

#for checking total amount----
orderSummary <- read_excel("Income.released.20220713_20221013.xls",col_types = c("numeric", "text", "text", "text", "text", "text", "text", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "text", "numeric", "numeric", "skip", "numeric", "text", "text"), skip = 5) %>%
  select(`Order ID`, `Order Creation Date`, `Payout Completed Date`, `Total Released Amount (RM)`) %>%
  filter(`Order Creation Date` > "2022-08-25") %>%
  group_split(`Order Creation Date`) %>%
  lapply(adorn_totals) %>% #add sum for total released amount
  write.xlsx(paste("shopeeOrderSummary", format(Sys.Date(), "%Y-%m-%d"), ".xlsx", sep = ""))

#daily orders summary----
dailyOrders <- autocountInvoice %>%
  lapply(mutate, 
         across(c(UnitPrice, `Vouchers & Discounts`, `Fees & Charges`), as.numeric),
         "Merchandise Subtotal" = UnitPrice * Qty,
         "Total" = `Merchandise Subtotal` - `Vouchers & Discounts` - `Fees & Charges`) %>%
  lapply(select, c(YourPONo, ISBN, ItemCode, Description, UnitPrice, Qty, `Merchandise Subtotal`, `Vouchers & Discounts`, `Fees & Charges`, Total)) %>%
  lapply(rename, 
         "Your P/O No." = YourPONo,
         "Item Code" = ItemCode,
         "Unit Price" = UnitPrice
         ) %>%
  lapply(adorn_totals) #add "Total" row

write.xlsx(dailyOrders, paste("dailyOrdersAsOf", format(Sys.Date(), "%Y-%m-%d"), ".xlsx", sep = "")) 
