
#style for tab 1 title
t1_title_style <- createStyle(
  fontName = "Arial",
  fontSize = 26,
  textDecoration = "bold",
  fontColour = "white",
  fgFill = "#005EB8",
  valign = "center",
  halign = "center"
)

#style for tab 1 table
t1_table_style <- createStyle(
  fontName = "Arial",
  fontSize = 11)

#tab 2 table body
t2_table_style_header <- 
  createStyle(
    fontName = "Arial",
    fontSize = 18,
    border = "LeftRight",
    borderStyle ="thick"
  )

t2_table_style_even <- 
  createStyle(
    fontName = "Arial",
    fontSize = 12,
    fontColour = "#03FCD7",
    fgFill = "#C6FC03",
    border = "LeftRight",
    borderStyle ="thick"
  )

t2_table_style_odd <- 
  createStyle(
    fontName = "Arial",
    fontSize = 12,
    fontColour = "#03FCD7",
    fgFill = "#FC03F0",
    border = "LeftRight",
    borderStyle ="thick"
  )