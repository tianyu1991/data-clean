
library(dplyr)

library("openxlsx")
mydf <- read.xlsx("All Activity Query.xlsx", sheet = 1, startRow = 1, colNames = TRUE)

# This statement prints out the number of unique values in each column - helpful for considering which columns
# need to be cleaned.
for (i in seq_along(names(mydf))){
  print(paste0("The column ", names(mydf[i]), " contains ",length(unique(mydf[,i])), " unique values."))
}

# Cleaning the data
# 1. Assuming that the addresses in the Clients table have already been standardized - do some basic formatting
# on FName, MiddleInitial, and LName columns

clientsTable<-mydf
colnames(clientsTable)[1] <- "ClientGUID"

clientsTable$First.Name <- tolower(clientsTable$First.Name)
clientsTable$Last.Name <- tolower(clientsTable$Last.Name)


# 2. Eliminating leading and trailing whitespace from all columns (may take a minute or two)
for (i in seq_along(names(clientsTable))){
  trimws(as.character(clientsTable[i]), which = c("both"))
}

# 3. Removing all formatting (spaces, parentheses, dashes) from phone numbers
clientsTable$HomePhone <- gsub("[^0-9.]", "", clientsTable$HomePhone)
clientsTable$WorkPhone <- gsub("[^0-9.]", "", clientsTable$WorkPhone)
clientsTable$CellPhone <- gsub("[^0-9.]", "", clientsTable$CellPhone)

# 4. Formatting the Address2 column (convert to all uppercase, remove 'Apt' string, etc.)
clientsTable$Address.2 <- tolower(clientsTable$Address.2)
clientsTable$Address.2 <- gsub("(apt|aptment)*(\\s)", "",clientsTable$Address.2)
clientsTable$Address.2 <- gsub("#", "",clientsTable$Address.2)

# Attempting to find duplicates based on the following column names:
dupeParameters <- c("First.Name", 
                    "Last.Name", 
                    "DOB", 
                    "Address(Final)", 
                    "Address.2", 
                    "Zip")


# See the following StackOverflow answer for more information about this function:
# https://stackoverflow.com/a/34294144
find_dupes = function(.table = clientsTable, ...) {
  require(dplyr)
  require(tidyr)
  # get column names of primary key
  pk <- .table %>% select(...) %>% names
  other <- names(.table)[!(names(.table) %in% pk)]
  # group by primary key,
  # get number of rows per unique combo,
  # filter for duplicates,
  # get number of distinct values in each column,
  # gather to get df of 1 row per primary key, other column,
  # filter for where a columns have more than 1 unique value,
  # order table by primary key
  
  .table %>%
    group_by(...) %>%
    mutate(cnt = n()) %>%
    filter(cnt > 1) %>%
    select(-cnt) %>%
    summarise_all(funs(n_distinct)) %>%
    gather_('column', 'unique_vals', other) %>%
    filter(unique_vals > 1) %>%
    arrange(...) %>%
    return
  # Final dataframe:
  ## One row per primary key and column that creates duplicates.
  ## Last column indicates how many unique values of
  ## the given column exist for each primary key.
}

# To edit this function: 
# 1. To save the data under another name, change 'example' to something else (no spaces)
# 2. To change the criteria to de-dupe by, edit the columns inside find_dupes()
example <- clientsTable %>% find_dupes(FName, LName, DOB)
# This just saves a few rows of example data 
write.csv(example[1:50,], "find_dupes_example_rows.csv", row.names = FALSE)


# This function takes a list of columns to check for dupes, like above. It returns a list of GUIDs
# matching the duplicate columns. Change .return_values to TRUE if you want to see the GUIDs.
find_ClientGUIDs <- function(.table, .return_values = FALSE, ...){
  require(dplyr)
  require(tidyr)
  
  if(.return_values){
  # group by primary key,
  # get number of rows per unique combo,
  # filter for duplicates,
  # get number of distinct values in each column,
  # gather to get df of 1 row per primary key, other column,
  # filter for where a columns have more than 1 unique value,
  # order table by primary key
  .table %>%
    select(... , ClientGUID) %>%
    distinct() %>%
    group_by(...) %>%
    summarise(guid_count = n_distinct(ClientGUID),
              guid_list = paste(ClientGUID, collapse = ",")) %>%
    filter(guid_count > 1) %>%
    return
  # Final dataframe:
  ## One row per primary key and column that creates duplicates.
  ## Last column indicates how many unique values of
  ## ClientGUID exist for each primary key. 
  } else {
    
    .table %>%
      group_by(...) %>%
      summarise(guid_count = n_distinct(ClientGUID)) %>%
      filter(guid_count > 1) %>%
      arrange(...) %>%
      return

  }
  
}
colnames(clientsTable)[1] <- "ClientGUID"

example.2 <- clientsTable %>% find_ClientGUIDs(.return_values = FALSE, LName, DOB)
example.3 <- clientsTable %>% find_ClientGUIDs(.return_values = TRUE, First.Name, Last.Name, DOB)
example.4 <- clientsTable %>% find_ClientGUIDs(.return_values = TRUE, LName, DOB, Address1)

write.csv(example.3, "find_ClientGUIDs_FName_LName_DOB.csv", row.names = FALSE)
write.csv(example.4, "find_ClientGUIDs_LName_DOB_Address1_dupes.csv", row.names = FALSE)

newdata <- subset(mydata, age >= 20 | age < 10, 
select=c(ID, Weight))
