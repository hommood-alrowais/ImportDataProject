doc_1 <- read_pptx()
sz <- slide_size(doc_1)

# add text and a table ----
doc_1 <- add_slide(doc_1, layout = "Two Content", master = "Office Theme")
doc_1 <- ph_with(
  x = doc_1, value = c("Table cars"),
  location = ph_location_type(type = "title"))

doc_1 <- ph_with(
  x = doc_1, value = names(cars),
  location = ph_location_left()
) 

doc_1 <- ph_with(
  x = doc_1, value = cars,
  location = ph_location_right()
) 


# add a base plot ----
anyplot <- plot_instr(code = {
  col <- c(
    "#440154FF", "#443A83FF", "#31688EFF",
    "#21908CFF", "#35B779FF", "#8FD744FF", "#FDE725FF"
  )
  barplot(1:7, col = col, yaxt = "n")
})


doc_1 <- add_slide(doc_1)
doc_1 <- ph_with(doc_1, anyplot,
                 location = ph_location_fullsize(),
                 bg = "#006699"
)




# add a ggplot2 plot ----
if (require("ggplot2")) {
  doc_1 <- add_slide(doc_1)
  gg_plot <- ggplot(data = iris) +
    geom_point(
      mapping = aes(Sepal.Length, Petal.Length),
      size = 3
    ) +
    theme_minimal()
  doc_1 <- ph_with(
    x = doc_1, value = gg_plot,
    location = ph_location_type(type = "body"),
    bg = "transparent"
  )
  doc_1 <- ph_with(
    x = doc_1, value = "graphic title",
    location = ph_location_type(type = "title")
  )
}


# Save the modified presentation
print(doc_1, target = "tablecars.pptx") 


