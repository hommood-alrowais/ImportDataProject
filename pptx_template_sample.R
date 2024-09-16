# Load libraries ----
library(tidyverse)
library(readxl)
library(lubridate)
library(dplyr)
library(gdata)
library(stringr)
library(ggplot2)
library(xts)
library(tableHTML)
library(kableExtra)
library(tableone)
library(gdata)
library(scales)
library(officer)

# Load Current Data ----
load("../../Data/currentData.Rdata")
# Download a PowerPoint template file from STHDA website
#download.file(url="http://www.sthda.com/sthda/RDoc/example-files/r-reporters-powerpoint-template.pptx",
#              destfile="r-reporters-powerpoint-template.pptx", quiet=TRUE)
# options('ReporteRs-fontsize'= 18, 'ReporteRs-default-font'='Arial')
doc <- read_pptx(path ="Jan 2024 SH Board Meeting Slide Deck.pptx" )
doc <- on_slide( doc, index = 2)
doc <- ph_with(x = doc, location = ph_location(left = 5,top = 5,width = 4,height = 4), value = "123456")
## Slide 1 : Title slide
#+++++++++++++++++++++++
doc <-add_slide(doc, "Title Slide")
print(doc,target = "test1.pptx")
