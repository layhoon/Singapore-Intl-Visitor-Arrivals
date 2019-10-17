library(xlsx)
library(dplyr)
library(tidyr)
library(zoo)


# Retrieve raw data

t1 <- read.xlsx("Data/International Visitor Arrivals.xlsx", sheetName = "T1", rowIndex = c(6:68), stringsAsFactors = FALSE)
t2 <- read.xlsx("Data/International Visitor Arrivals.xlsx", sheetName = "T2", rowIndex = c(6:18), stringsAsFactors = FALSE)
t3 <- read.xlsx("Data/International Visitor Arrivals.xlsx", sheetName = "T3", rowIndex = c(6:22), stringsAsFactors = FALSE)


# Extract global arrivals from worksheet T1

t11 <- t1[1, -1]
t11 <- gather(t11, Month, Arrivals)
t11$Month <- gsub("X", "", t11$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t11$Arrivals <- gsub(",", "", t11$Arrivals) %>% as.numeric()
t11 <- arrange(t11, Month)


# Separate column "Variables" from worksheet T1 into "Region" and "Country"

t1 <- t1[-1,]
colnames(t1)[1] <- "Country"
t1 <- mutate(t1, Region = "")

x <- grep("^[ ]{8}[A-Z]", t1$Country)
t1$Country <- trimws(t1$Country)
t1$Region[x] <- t1$Country[x]
t1$Country[x] <- ""

len <- nrow(t1)
for (i in 2:len)
  if (t1$Region[i] == "")
    t1$Region[i] <- t1$Region[i-1]

t1$Country[which(t1$Country == "USA")] <- "United States"


# Duplicate arrivals from "Others" region at country level

others <- t1[t1$Region == "Others",]
others$Country <- "Other Markets in Others"
t1 <- rbind(t1, others)


# Extract arrivals by region from worksheet T1

t12 <- t1[x, -1]
t12 <- gather(t12, Month, Arrivals, -Region)
t12$Month <- gsub("X", "", t12$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t12$Arrivals <- gsub(",", "", t12$Arrivals) %>% as.numeric()
t12 <- arrange(t12, Region, Month)


# Extract arrivals by country from worksheet T1

t13 <- t1[-x,]
t13 <- gather(t13, Month, Arrivals, -Region, -Country)
t13$Month <- gsub("X", "", t13$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t13$Arrivals <- gsub(",", "", t13$Arrivals) %>% as.numeric()
t13 <- arrange(t13, Region, Country, Month)


# Update any NA arrivals in the regional data with the country data

t12Agg <- group_by(t13, Region, Month) %>% summarise(Arrivals = sum(Arrivals, na.rm = TRUE)) %>% arrange(Region, Month)
x <-which(is.na(t12$Arrivals))
t12$Arrivals[x] <- t12Agg$Arrivals[x]


# Update any NA arrivals in the global data with the regional data

t11Agg <- group_by(t12, Month) %>% summarise(Arrivals = sum(Arrivals, na.rm = TRUE)) %>% arrange(Month)
x <- which(is.na(t11$Arrivals))
t11$Arrivals[x] <- t11Agg$Arrivals[x]


# Define country abbreviations

country <- c(
  "Australia", "AUS",
  "Bangladesh", "BGD",
  "Belgium & Luxembourg", "BEL & LUX",
  "Brunei Darussalam", "BRN",
  "Canada", "CAN",
  "China", "CHN",
  "Denmark", "DNK",
  "Egypt", "EGY",
  "Finland", "FIN",
  "France", "FRA",
  "Germany", "DEU",
  "Hong Kong SAR", "HKG",
  "India", "IND",
  "Indonesia", "IDN",
  "Iran", "IRN",
  "Israel", "ISR",
  "Italy", "ITA",
  "Japan", "JPN",
  "Kuwait", "KWT",
  "Malaysia", "MYS",
  "Mauritius", "MUS",
  "Myanmar", "MMR",
  "Netherlands", "NLD",
  "New Zealand", "NZL",
  "Norway", "NOR",
  "Pakistan", "PAK",
  "Philippines", "PHL",
  "Rep Of Ireland", "IRL",
  "Russian Federation", "RUS",
  "Saudi Arabia", "SAU",
  "South Africa (Rep Of)", "ZAF",
  "South Korea", "KOR",
  "Spain", "ESP",
  "Sri Lanka", "LKA",
  "Sweden", "SWE",
  "Switzerland", "CHE",
  "Taiwan", "TWN",
  "Thailand", "THA",
  "United Arab Emirates", "ARE",
  "United Kingdom", "GBR",
  "United States", "USA",
  "Vietnam", "VNM",
  "Other Markets In Africa", "Other Markets In Africa",
  "Other Markets In Americas", "Other Markets In Americas",
  "Other Markets In Europe", "Other Markets In Europe",
  "Other Markets In Greater China", "Other Markets In Greater China",
  "Other Markets In North Asia", "Other Markets In North Asia",
  "Other Markets In Oceania", "Other Markets In Oceania",
  "Other Markets In South Asia", "Other Markets In South Asia",
  "Other Markets In Southeast Asia", "Other Markets In Southeast Asia",
  "Other Markets In West Asia", "Other Markets In West Asia",
  "Other Markets In Others", "Other Markets In Others")

country <- data.frame(matrix(country, ncol = 2, byrow = TRUE))
colnames(country) <- c("Name", "Code")


# Extract global arrivals from worksheet T2

t21 <- t2[1, -1]
t21 <- gather(t21, Month, Arrivals)
t21$Month <- gsub("X", "", t21$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t21$Arrivals <- gsub(",", "", t21$Arrivals) %>% as.numeric()
t21 <- arrange(t21, Month)


# Extract arrivals by sex from worksheet T2

t22 <- filter(t2, Variables %in% c("Males", "Females"))
colnames(t22)[1] <- "Sex"
t22$Sex[which(t22$Sex == "Males")] <- "Male"
t22$Sex[which(t22$Sex == "Females")] <- "Female"
t22 <- gather(t22, Month, Arrivals, -Sex)
t22$Month <- gsub("X", "", t22$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t22$Arrivals <- gsub(",", "", t22$Arrivals) %>% as.numeric()


# Extract arrivals by age group from worksheet T2

t23 <- t2[grep("[1-9]", t2$Variables),]
colnames(t23)[1] <- "AgeGroup"
t23$AgeGroup <- trimws(t23$AgeGroup)
t23 <- gather(t23, Month, Arrivals, -AgeGroup)
t23$Month <- gsub("X", "", t23$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t23$Arrivals <- gsub(",", "", t23$Arrivals) %>% as.numeric()


# Extract global arrivals from worksheet T3

t31 <- t3[1, -1]
t31 <- gather(t31, Month, Arrivals)
t31$Month <- gsub("X", "", t31$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t31$Arrivals <- gsub(",", "", t31$Arrivals) %>% as.numeric()
t31 <- arrange(t31, Month)

# Extract arrivals by length of stay from worksheet T3

t32 <- t3[-c(1, nrow(t3)),]
colnames(t32)[1] <- "LengthOfStay"
t32$LengthOfStay <- trimws(t32$LengthOfStay)
t32 <- gather(t32, Month, Arrivals, -LengthOfStay)
t32$Month <- gsub("X", "", t32$Month) %>% as.yearmon("%Y.%b") %>% as.Date() %>% as.character() %>% substr(start = 1, stop = 7)
t32$Arrivals <- gsub(",", "", t32$Arrivals) %>% as.numeric()


# Write processed arrivals into xlsx files

write.xlsx(t11, "Data/International Visitor Arrivals T1.xlsx", sheetName = "Global Arrivals", row.names = FALSE)
write.xlsx(t12, "Data/International Visitor Arrivals T1.xlsx", sheetName = "Arrivals By Region", row.names = FALSE, append = TRUE)
write.xlsx(t13, "Data/International Visitor Arrivals T1.xlsx", sheetName = "Arrivals By Country", row.names = FALSE, append = TRUE)
write.xlsx(country, "Data/International Visitor Arrivals T1.xlsx", sheetName = "Country", row.names = FALSE, append = TRUE)

write.xlsx(t21, "Data/International Visitor Arrivals T2.xlsx", sheetName = "Global Arrivals", row.names = FALSE)
write.xlsx(t22, "Data/International Visitor Arrivals T2.xlsx", sheetName = "Arrivals By Sex", row.names = FALSE, append = TRUE)
write.xlsx(t23, "Data/International Visitor Arrivals T2.xlsx", sheetName = "Arrivals By Age Group", row.names = FALSE, append = TRUE)

write.xlsx(t31, "Data/International Visitor Arrivals T3.xlsx", sheetName = "Global Arrivals", row.names = FALSE)
write.xlsx(t32, "Data/International Visitor Arrivals T3.xlsx", sheetName = "Arrivals By Length Of Stay", row.names = FALSE, append = TRUE)