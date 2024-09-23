#install.packages("officer") 
#install.packages ("ggplot2")
#install.packages ("dplyr")
#install.packages ("lubridate")

library(officer) 
library(ggplot2)
library(dplyr)
library(lubridate)


ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\ImportDataProject\\FINALOUTPUT.pptx")


# Select the first slide
ppt <- on_slide(ppt, index = 1)

# Formatting the date to today's date
formatted_date <- format(Sys.Date(), "%B %Y")

# Define text formatting for the date
date_style <- fp_text(font.size = 35, font.family = "Garnett 2", bold = FALSE, color = "black")

# Create a formatted text object using fpar() and ftext()
formatted_text <- fpar(
  ftext(formatted_date, prop = date_style)
)

# Add the formatted date text to slide 1 at a specific location
ppt <- ph_with(ppt, value = formatted_date, 
               location = ph_location(left = 12.8, top = 9, width = 5, height = 1))

# Save the presentation after adding the date to slide 1
print(ppt, target = "updated_presentation_with_date.pptx")


#ppt <- ph_with(ppt, value = formatted_text, 
# location = ph_location(left = 12.8, top = 9, width = 5, height = 1))
# Get the current month and year dynamically
#formatted_date <- format(Sys.Date(), "%B %Y")



# Select the second slide
ppt <- on_slide(ppt, index = 2)


# Add the new dynamic date in "TextBox 6" with the dynamic date (formatted_date)
ppt <- ph_with(ppt, value = formatted_date, 
               location = ph_location(left = 4.4471861, top = 3.253519, width = 3.662061461, height = 0.4251717))

# Save the modified presentation
print(ppt, target = "updated_presentation_with_date.pptx")


# Now move to the second slide
ppt <- on_slide(ppt, index = 2)

# Check slide summary (optional)
# slide_summary(ppt)


# Define multi-line formatted text with font size adjustments
formatted_text <- fpar(
  ftext("• 85% of people report improved mood after regular exercise, follow their example and take care of yourself.", 
        prop = fp_text(font.size = 28, bold = FALSE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n• 70% of individuals aged 30-50 engage in physical activity at least 3 times per week.", 
        prop = fp_text(font.size = 28, bold = FALSE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n• Regular physical activity can help reduce stress and anxiety, leading to a more balanced life.", 
        prop = fp_text(font.size = 28, bold = FALSE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n• Studies show that people who exercise regularly sleep better and have higher energy levels.", 
        prop = fp_text(font.size = 28, bold = FALSE, color = "black", font.family = "Garnett 1"))
)


# Add the formatted multi-line text to slide 2 at the specified location
ppt <- ph_with(ppt, value = formatted_text, 
               location = ph_location(left = 10, top = 4, width = 9.5, height = 15))

# Save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")



# Now select the second slide
ppt <- on_slide(ppt, index = 2)

# Define the title using fpar() and ftext() for consistent formatting
formatted_title <- fpar(
  ftext("Health Benefits of Physical Activity", 
        prop = fp_text(font.size = 32, bold = TRUE, color = "black", font.family = "Garnett 1"))
)

# Add the formatted title to slide 2 at the specified location
ppt <- ph_with(ppt, value = formatted_title, 
               location = ph_location(left = 4.5, top = 2, width = 6, height = 1))


# Save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")




# Now select the second slide
ppt <- on_slide(ppt, index = 5)


# Define the title using fpar() and ftext() for consistent formatting
formatted_title <- fpar(
  ftext("Look at this picture of Hommood fixing the Ipad", 
        prop = fp_text(font.size = 32, bold = FALSE, color = "black", font.family = "Garnett 2"))
)

# Add the formatted title to slide 2 at the specified location
ppt <- ph_with(ppt, value = formatted_title, 
               location = ph_location(left = 13, top = 6.3, width = 6, height = 1))

# Save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")




# Now select the second slide
ppt <- on_slide(ppt, index = 5)


# Define the title using fpar() and ftext() for consistent formatting
formatted_title <- fpar(
  ftext("100 IPADS", 
        prop = fp_text(font.size = 41.7, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n 100 COMPUTERS", 
        prop = fp_text(font.size = 41.7, bold = TRUE, color = "black", font.family = "Garnett 1"))
  
)

# Add the formatted title to slide 2 at the specified location
ppt <- ph_with(ppt, value = formatted_title, 
               location = ph_location(left = 13, top = 8, width = 6, height = 1))

# Save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")



### Next level


#install.packages("pdftools")
#install.packages("magick")

library(pdftools)
library(magick)


#Converting the pdf to png 

# Path to your PDF file
pdf_file <- "thisPlot.pdf"

# Convert PDF to PNG
pdf_convert(pdf_file, format = "png", dpi = 300)  # This should give "thisPlot_1.png"

# Make the background of the PNG transparent (if needed)
png_image <- image_read("thisPlot_1.png")
transparent_image <- image_transparent(png_image, "white")

# Save the image with transparency
image_write(transparent_image, path = "thisPlot_transparent.png", format = "png")


#Now saving the image in the ppt


slide_summary(ppt)

# Go to slide 6
ppt <- on_slide(ppt, index = 6)


# Path to the PNG image
png_image <- "thisPlot_transparent.png"


# Add the PNG image to slide 6
ppt <- ph_with(ppt, external_img(png_image, width = 10, height = 13), 
               location = ph_location(left = 10, top = 4, width = 9, height = 6))

# Save the updated PowerPoint presentation
print(ppt, target = "updated_presentation_with_date.pptx")

# Go to slide 6
ppt <- on_slide(ppt, index = 6)

# Define the title using fpar() and ftext() for consistent formatting
formatted_title <- fpar(
  ftext("Histogram of Ages", 
        prop = fp_text(font.size = 30, bold = TRUE, color = "black", font.family = "Garnett 1"))
)

# Add the formatted title to slide 6 at the specified location
ppt <- ph_with(ppt, value = formatted_title, 
               location = ph_location(left = 10, top = 2, width = 6, height = 1))

# Now save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")



#Testing another way to to a plot in ppt 

# Install and load required packages
# install.packages("officer")
install.packages("rvg")
library(officer)
library(rvg)
library(ggplot2)

# Create a ggplot plot
data <- data.frame(age = rnorm(100, mean = 30, sd = 10))
p <- ggplot(data, aes(x = age)) +
  geom_histogram(binwidth = 5, fill = "blue", color = "black") +
  labs(title = "Histogram of Ages", x = "Age", y = "Frequency")


# Go to slide 7
ppt <- on_slide(ppt, index = 7)


# Add a slide and include the ggplot as an editable vector graphic
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with_vg(ppt, code = print(p), type = "body")

# Save the PowerPoint file
print(ppt, target = "presentation_with_ggplot.pptx")

# Now save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")