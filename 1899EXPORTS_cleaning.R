## ---------------------------
##
## Script name: 1899EXPORTS_cleaning.R
##
## Purpose of script: sample cleaning for 1899 EXPORTS.xlsx in Canadian data
##
## Author: Yuxin Cai
##
## Date Created: 2022-06-20
##
## Email: yuxin.cai@mail.utoronto.ca
##
## ---------------------------

## ---------------------------read input file---------------------------------
library("tidyverse")
library("readxl")
library("writexl")
library("janitor")
# read in excel
my_data <- read_excel("~/Desktop/1899 EXPORTS.xlsx", 
                      sheet = 1, guess_max = 10000)
# set variable year
year <- 1899
# select columns 
my_data %>% select(1:8) -> my_data
# rename the cols
my_data %>% 
  rename("Good" = 1) %>%
  rename("Country" = 2) %>%
  rename("POC_Quantity" = 3) %>%
  rename("POC_Value" = 4) %>%
  rename("NPOC_Q" = 5) %>%
  rename("NPOC_V"  = 6) %>%
  rename("Total_Quantity" = 7) %>%
  rename("Total_Value" = 8) -> my_data
# remove empty rows
my_data %>% remove_empty("rows") -> my_data
## ---------------------------make unit table ---------------------------------
# move units down one cell
my_data %>% 
  mutate(`POC_Quantity` = lag(`POC_Quantity`, 1)) -> my_copy
# extract Good and POC Quantity
my_copy %>% select(`Good`, `POC_Quantity`) -> my_copy

# remove title rows
my_copy %>% 
  filter(!grepl("Quantity", `POC_Quantity`)) %>% 
  filter(!grepl("GOO", `POC_Quantity`)) %>% 
  filter(!grepl("Goo",`POC_Quantity`)) %>% 
  filter(!grepl("CANADA",`POC_Quantity`))%>% 
  filter(!grepl("THE",`POC_Quantity`))-> my_copy
# remove numbers
my_copy %>% 
  filter(!grepl("\\d", `POC_Quantity`))-> my_copy
# remove the nas
my_copy %>% 
  filter(!is.na(`POC_Quantity`)) %>% 
  filter(!is.na(`Good`)) -> my_copy

# rename the unit col
my_copy %>% 
  rename ("Measure" = "POC_Quantity") -> my_copy
# remove duplicates rows
units <- unique(my_copy)

##---------------------------------cleaning---------------------------------
# replace spaces by period
my_data$`POC_Quantity` <- gsub("\\s+", ".", my_data$`POC_Quantity`)
my_data$`POC_Value` <- gsub("\\s+", ".", my_data$`POC_Value`)
my_data$`NPOC_Q` <- gsub("\\s+", ".", my_data$`NPOC_Q`)
my_data$`NPOC_V` <- gsub("\\s+", ".", my_data$`NPOC_V`)
my_data$`Total_Quantity` <- gsub("\\s+", ".", my_data$`Total_Quantity`)
my_data$`Total_Value` <- gsub("\\s+", ".", my_data$`Total_Value`)
# remove comma
my_data$`POC_Quantity` <- gsub(",", "", my_data$`POC_Quantity`)
my_data$`POC_Value` <- gsub(",", "", my_data$`POC_Value`)
my_data$`NPOC_Q` <- gsub(",", "", my_data$`NPOC_Q`)
my_data$`NPOC_V` <- gsub(",", "", my_data$`NPOC_V`)
my_data$`Total_Quantity` <- gsub(",", "", my_data$`Total_Quantity`)
my_data$`Total_Value` <- gsub(",", "", my_data$`Total_Value`)
# change data type from chr to num
# if keeping other data, use text to cols
my_data$`POC_Quantity` <- as.numeric(my_data$`POC_Quantity`)
my_data$`POC_Value` <- as.numeric(my_data$`POC_Value`)
my_data$`NPOC_Q` <- as.numeric(my_data$`NPOC_Q`)
my_data$`NPOC_V` <- as.numeric(my_data$`NPOC_V`)
my_data$`Total_Quantity` <- as.numeric(my_data$`Total_Quantity`)
my_data$`Total_Value` <- as.numeric(my_data$`Total_Value`)

# remove extra rows by checking country
my_data %>% 
  filter(!is.na(`Country`)) %>%
  filter(!grepl("AND",`Country`)) %>%
  filter(!grepl("Countr",`Country`)) %>%
  filter(!grepl("Column",`Country`)) %>%
  filter(!grepl("PROVINCE",`Country`)) %>%
  filter(!grepl("COUNT",`Country`))%>%
  filter(!grepl("Total",`Country`))%>%
  filter(!grepl("ARTICLES",`Good`))-> my_data
# set cells including "Total" to NA in the col Good
my_data$Good <- replace(my_data$Good, grepl("Total",my_data$Good), NA)
# fill the Goods col 
my_data %>% fill(Good) -> my_data

##---------------------------------separation---------------------------------
# read in data set
table_before1955 <- data.frame(province = c("Ontario",
                                            "Quebec",
                                            "Nova Scotia",
                                            "P. E. Island",
                                            "N. Brunswick",
                                            "B. Columbia",
                                            "Alberta",
                                            "Manitoba",
                                            "Saskatchewan",
                                            "N. W. Ter"))
table_after1955 <- data.frame(province = c("Ontario",
                                           "Quebec",
                                           "Nova Scotia",
                                           "P. E. Island",
                                           "N. Brunswick",
                                           "B. Columbia",
                                           "Alberta",
                                           "Manitoba",
                                           "Saskatchewan",
                                           "Newfoundland",
                                           "N. W. Ter"))
# separate into two tables
my_data %>% filter(Country %in% table_before1955$province) -> my_province
my_data %>% filter(!Country %in% table_before1955$province) -> my_country
# rename cols in my_province 
my_province %>% 
  rename("Province" = "Country") -> my_province
# join two tables
# create result data frame
result <- data.frame(Good = character(),
                     Country = character(),
                     Province = character(),
                     POC_Quantity = numeric(),
                     POC_Value = numeric(),
                     NPOC_Quantity = numeric(),
                     NPOC_Value = numeric(),
                     Total_Quantity = numeric(),
                     Total_Value = numeric())

country1 <- data.frame(Good = character(),
                       Country = character(),
                       `POC_Quantity` = numeric(),
                       `POC_Value` = numeric(),
                       `NPOC_Quantity` = numeric(),
                       `NPOC_Value` = numeric(),
                       `Total_Quantity` = numeric(),
                       `Total_Value` = numeric())
province1 <- data.frame(Good = character(),
                        Province = character(),
                        `POC_Quantity` = numeric(),
                        `POC_Value` = numeric(),
                        `NPOC_Quantity` = numeric(),
                        `NPOC_Value` = numeric(),
                        `Total_Quantity` = numeric(),
                        `Total_Value` = numeric())
# the list of goods
goods <- unique(my_data$Good)
# fill in country1 and province1 by looping goods
for(item in goods){
  # find rows with corresponding good
  my_country %>% filter(Good == item) -> country
  # add those row to country1
  country1 <- rbind(country1, country)
  # add summation for each section
  country1[1+nrow(country1),] <- list(item, "Total", 
                                      sum(country$`POC_Quantity`),
                                      sum(country$`POC_Value`), 
                                      sum(country$`NPOC_Q`),
                                      sum(country$`NPOC_V`), 
                                      sum(country$`Total_Quantity`),
                                      sum(country$`Total_Value`))
  # same pocedures for province1
  my_province %>% filter(Good == item) -> province
  province1 <- rbind(province1, province)
  province1[1+nrow(province1),] <- list(item, "Total", 
                                        sum(province$`POC_Quantity`),
                                        sum(province$`POC_Value`), 
                                        sum(province$`NPOC_Q`),
                                        sum(province$`NPOC_V`),
                                        sum(province$`Total_Quantity`),
                                        sum(province$`Total_Value`))
}
# fill in the empty data frame
for(i in 1:nrow(country1)) {
  result[i,]$Good <- country1[i,]$Good
  result[i,]$Country <- country1[i,]$Country
  result[i,]$POC_Quantity <- country1[i,]$`POC_Quantity`
  result[i,]$POC_Value <- country1[i,]$`POC_Value`
  result[i,]$NPOC_Quantity <- country1[i,]$`NPOC_Q`
  result[i,]$NPOC_Value <- country1[i,]$`NPOC_V`
  result[i,]$Total_Quantity <- country1[i,]$`Total_Quantity`
  result[i,]$Total_Value <- country1[i,]$`Total_Value`
}

for(i in 1:nrow(province1)) {
  result[i+nrow(country1),]$Good <- province1[i,]$Good
  result[i+nrow(country1),]$Province <- province1[i,]$Province
  result[i+nrow(country1),]$`POC_Quantity` <- province1[i,]$`POC_Quantity`
  result[i+nrow(country1),]$`POC_Value` <- province1[i,]$`POC_Value`
  result[i+nrow(country1),]$`NPOC_Quantity` <- province1[i,]$`NPOC_Q`
  result[i+nrow(country1),]$`NPOC_Value` <- province1[i,]$`NPOC_V`
  result[i+nrow(country1),]$`Total_Quantity` <- province1[i,]$`Total_Quantity`
  result[i+nrow(country1),]$`Total_Value` <- province1[i,]$`Total_Value`
}
##--------------merge unit table and the data table----------------------------
# join the space table with the data file
result %>% 
  left_join(units, by = "Good" ) -> join1
# reorder the columns
join1 <- join1[, c(1,10,2,3,4,5,6,7,8,9)]
# add the year column
join1 %>% mutate(Year = year) -> join1
# reorder the columns
join1 <- join1[, c(11,1,2,3,4,5,6,7,8,9,10)]
# fill in good and measure
join1 %>% 
  fill(Good) %>%
  fill(Measure) -> join1

##---------------------------------output---------------------------------
write_xlsx(join1,"~/Desktop/1899result.xlsx")



