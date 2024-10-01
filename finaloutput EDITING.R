library(officer)
library(ggplot2)
library(dplyr)
library(lubridate)

# Open the PowerPoint presentation
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

# Add the formatted date text without a bullet point using `ph_with()`
# Add `level` argument set to 0 to avoid bullets
ppt <- ph_with(ppt, value = formatted_text, 
               location = ph_location(left = 12.8, top = 9, width = 5, height = 1),
               level = 0)

# Save the presentation after adding the date to slide 1
print(ppt, target = "updated_presentation_with_date.pptx")


####Moving to the second slide

# Select the second slide
ppt <- on_slide(ppt, index = 2)

# Formatting the date to today's date
formatted_date <- format(Sys.Date(), "%B %Y")

# Define text formatting for the date
date_style <- fp_text(font.size = 35, font.family = "Garnett 2", bold = FALSE, color = "black")

# Create a formatted text object using fpar() and ftext()
formatted_text <- fpar(
  ftext(formatted_date, prop = date_style)
)

# Add the formatted date text without a bullet point using `ph_with()`
# Add `level` argument set to 0 to avoid bullets
ppt <- ph_with(ppt, value = formatted_text, 
               location = ph_location(left = 4.4471861, top = 3.253519, width = 3.662061461, height = 0.4251717),
               level = 0)

# Save the modified presentation
print(ppt, target = "updated_presentation_with_date.pptx")



# Going to the second slide
ppt <- on_slide(ppt, index = 2)

#Check slide summary (optional)
#slide_summary(ppt)


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




# Now select slide 5
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




# Now select slide 5
ppt <- on_slide(ppt, index = 5)


# Define the title using fpar() and ftext() for consistent formatting
formatted_title <- fpar(
  ftext("100 IPADS", 
        prop = fp_text(font.size = 41.7, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n 100 COMPUTERS", 
        prop = fp_text(font.size = 41.7, bold = TRUE, color = "black", font.family = "Garnett 1"))
  
)

# Add the formatted title to slide 5 at the specified location
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




# Now select slide 7
ppt <- on_slide(ppt, index = 7)


# Define paragraph properties with custom line spacing (e.g., 1.5)
line_spacing_prop <- fp_par(line_spacing = 1.58)

# Define the title using fpar() and ftext() for consistent formatting and spacing
formatted_title <- fpar(
  ftext("100 IPADS", 
        prop = fp_text(font.size = 35, bold = FALSE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n100 COMPUTERS", 
        prop = fp_text(font.size = 35, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n100 COMPUTERS", 
        prop = fp_text(font.size = 35, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n100 COMPUTERS", 
        prop = fp_text(font.size = 35, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n100 COMPUTERS", 
        prop = fp_text(font.size = 35, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  ftext("\n100 COMPUTERS", 
        prop = fp_text(font.size = 35, bold = TRUE, color = "black", font.family = "Garnett 1")),
  
  fp_p = line_spacing_prop  # Apply line spacing to the whole paragraph
)


# Add the formatted title to slide 2 at the specified location
ppt <- ph_with(ppt, value = formatted_title, 
               location = ph_location(left = 13, top = 3.5, width = 6, height = 1))

# Save the final version of the presentation
print(ppt, target = "updated_presentation_with_date.pptx")




# Load necessary libraries
library(officer)
library(flextable)
library(dplyr)


# Select slide 9 in the existing presentation
ppt <- on_slide(ppt, index = 9)

# Create fake health data for demonstration purposes
set.seed(123)  # Set seed for reproducibility

# Generate a data frame with health-related columns
health_data <- data.frame(
  Name = paste("Person", 1:10),                       # Person names
  Age = sample(25:65, 10, replace = TRUE),             # Random ages between 25 and 65
  Height_cm = round(runif(10, 150, 190), 1),           # Random height between 150 and 190 cm
  Weight_kg = round(runif(10, 50, 90), 1),             # Random weight between 50 and 90 kg
  BMI = round(runif(10, 18, 30), 1)                    # Random BMI between 18 and 30
)

# Print the health data (optional)
print(health_data)

# Convert the health data to a flextable
health_flextable <- flextable(health_data)

# Optionally, format the table (adjust column widths, add borders, etc.)
health_flextable <- autofit(health_flextable)

# Add the flextable to slide 9 at a specified location
ppt <- ph_with(ppt, value = health_flextable, 
               location = ph_location(left = 1, top = 1, width = 8, height = 5))


# Save the modified PowerPoint presentation
print(ppt, target = "updated_presentation_with_date.pptx")



#Creating tables with R.

library(officer)
library(flextable)
library(dplyr)



# Select slide 11
ppt <- on_slide(ppt, index = 11)

# Generate random health data for the table
set.seed(123)  # Set seed for reproducibility
names <- paste("Person", 1:10)  # Names of the persons
ages <- sample(25:65, 10, replace = TRUE)  # Random ages between 25 and 65
heights <- round(runif(10, 150, 190), 1)  # Random height between 150 and 190 cm
weights <- round(runif(10, 50, 90), 1)  # Random weight between 50 and 90 kg
bmi <- round(weights / ((heights / 100) ^ 2), 1)  # Calculate BMI

# Create a data frame for health data
health_data <- data.frame(
  Name = names,
  Age = ages,
  Height_cm = heights,
  Weight_kg = weights,
  BMI = bmi
)

# Convert the health data to a flextable
health_flextable <- flextable(health_data)

# Optionally, format the table (adjust column widths, add borders, etc.)
health_flextable <- autofit(health_flextable)

# Add the flextable to slide 9 at a specified location
ppt <- ph_with(ppt, value = health_flextable, 
               location = ph_location(left = 8, top = 3, width = 8, height = 5))

# Save the modified PowerPoint presentation
print(ppt, target = "updated_presentation_with_date.pptx")