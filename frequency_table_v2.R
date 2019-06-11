#### Setting Up ####
if(!require("data.table")){install.packages("data.table", dependencies = T) ; library("data.table")} # Install the package if not already installed, and load it
if(!require("dplyr"))     {install.packages("dplyr",      dependencies = T) ; library("dplyr")     }
if(!require("magrittr"))  {install.packages("magrittr",   dependencies = T) ; library("magrittr")  }
if(!require("openxlsx"))  {install.packages("openxlsx",   dependencies = T) ; library("openxlsx")  }
if(!require("rstudioapi")){install.packages("rstudioapi", dependencies = T) ; library("rstudioapi")}
gc();                                                       # Garbage collection
setwd(dirname(rstudioapi::getActiveDocumentContext()$path)) # Set working directory to the folder code is in 
rm(list=ls()); cat("\014");

#### Read input dataset and drop redundant columns ####
dataset = fread("C:/Users/Ashrith Reddy/Desktop/justice/02_data/01_pos/updated_point_of_sale_01242018.csv",nrows = -1) %>% 
  select(-HOUSEHOLD_ID, -INDIVIDUAL_ID,-TRANSACTION_ID)

dataset = dataset %>% mutate_if(is.factor, as.character)

#### Frequency table generation ####
wb = openxlsx::createWorkbook("frequency_table.xlsx")     # Creating and mounting an empty excel
for(i in 1:ncol(dataset)){                                # Loop over each column
  message(names(dataset)[i])
  gc()
  
  sheetname = names(dataset)[i]
  openxlsx::addWorksheet(wb, ifelse(nchar(sheetname)>31, paste0(substr(sheetname,1,15),substr(sheetname,nchar(sheetname)-15,nchar(sheetname)))  , sheetname)) # Add a new sheet
  
  one_columns = dataset[,i,with = FALSE]                  # Retain column of interest
  one_columns = one_columns[,.(Count=.N),by = c(names(dataset)[i])] %>% as.data.frame() # Aggregate to get frequency
  one_columns$Percentage = 100*one_columns$Count/sum(one_columns$Count)
  one_columns = one_columns %>% arrange(desc(Count)) %>% mutate(sl_no = 1:nrow(.), string_length = nchar(.[[names(dataset)[i]]]) ) %>% select(sl_no,everything(),string_length)
  
  openxlsx::writeData(wb, i, one_columns)                 # Write frequency table to excel
}
openxlsx::saveWorkbook(wb, file = "frequency_table.xlsx") # Write and unmount final excel
rm(i,one_columns,wb)
