#install.packages("officer") 

library(officer) 

ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\ImportDataProject\\FINALOUTPUT.pptx")


# Select the first slide
ppt <- on_slide(ppt, index = 1)

slide_summary(ppt)

#Add new date at a specific location (coordinates: left = 2, top = 3, width = 5, height = 1)

ppt <- ph_with(ppt, value = "September 2024", location = ph_location(left = 15, top = 9, width = 5, height = 1))
              
formatted_date <- format(Sys.Date(), "%B %Y") 

# Define text formatting (e.g., font size 20, Arial, bold, blue color)
text_style <- fp_text(font.size = 35, font.family = "Garnett 1", bold = FALSE, color = "black")


# Add the formatted date with the defined style to the slide
ppt <- ph_with(ppt, 
               value = fpar(ftext(formatted_date, prop = text_style)), 
               location = ph_location(left = 14, top = 7, width = 5, height = 1))



# Select the first slide
ppt <- on_slide(ppt, index = 1)


# Add the text with the defined style to the slide
ppt <- ph_with(ppt, 
               value = fpar(ftext("September 2024", prop = text_style)), 
               location = ph_location(left = 14, top = 7, width = 5, height = 1))




# Save the modified presentation
print(ppt, target = "updated_presentation_with_date.pptx")
             


               

