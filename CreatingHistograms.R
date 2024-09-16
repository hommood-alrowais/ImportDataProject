
library("ggplot2")
library(readxl)
library(tidyverse)

library(dplyr)
library(lubridate)
library(readr)

# Load the data
HQTest <- read_excel("DemographicsHQTEST.xlsx")



HQTest <- HQTest%>%
  # If I wanted to Add 5+ on EmrID column:
  mutate(EmrID = EmrID + 5, 
         ##Changing the Surname column to uppercase
         Surname = toupper(Surname), 
         #Adding a column for age using their birth date until today
         Age = floor(interval(`Birth date`, today()) / years(1)))  
  


# Creating the Age Histogram

pdf("thisPlot.pdf")
##  aes(x = age) maps the age variable to the x-axis.
ggplot(HQTest, aes(x = HQTest$Age)) + 
  
##  geom_histogram(binwidth = 5) creates the histogram, with a bin width of 5 years 
geom_histogram(binwidth = 5, 

##  fill = "skyblue" and color = "black" set the bar color and outline.
fill = "skyblue", color = "black") +
  
##  labs() adds a title and labels for the x and y axes.
labs(title = "Histogram of Ages", x = "Age", y = "Frequency") +
  
##  theme_minimal() gives a clean, minimal background to the plot.  
theme_minimal()

dev.off()


# Creating a Histogram of the number of first visits per month

# Load the data
HQTest <- read_excel("DemographicsHQTEST.xlsx")


# Ensure 'date' column is in Date format
HQTest$Date <- as.Date(HQTest$Date)

#Renaming "Date" to "VisitDate"
HQTest <- rename(HQTest, VisitDate = Date)

# Create a new column for year and month
HQTest <- HQTest %>%
  mutate(year_month = floor_date(VisitDate, "month"))

# Summarize the number of visits per year/month
data_summary <- HQTest %>%
  group_by(year_month) %>%
  summarise(visits = n(), .groups = "drop")



# Create the histogram plot
ggplot(data_summary, aes(x = year_month, y = visits)) +
  geom_bar (stat = "identity", fill = "skyblue", color = "black") + 
  labs(title = "Number of Visits Per Month", y = "Number of Visists") +
  theme_minimal() +
  scale_x_date(date_labels = "%Y-%m", date_breaks = "1 month") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
  

