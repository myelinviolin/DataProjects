library(openxlsx)

filepath <- "S:/Study"
filename <- "Stream-QueryDetail Post"
date <- "10 March 2022"
{
data <- read.csv(file = paste0(filepath, filename, ".csv"), header = TRUE)

wb <- createWorkbook()
sheet <- addWorksheet(wb, sheetName = "Query Detail")
csb <- createStyle(textDecoration = "bold", wrapText = TRUE)
cswrap <- createStyle(wrapText = TRUE)

colw <- matrix(c(5, 15.47,
                 13, 17.27,
                 17, 11.73,
                 19, 50,
                 21, 16.73,
                 23, 15.13,
                 24, 12,
                 25, 33,
                 27, 13, 
                 29, 5.2,
                 31, 8.4,
                 32, 9), byrow = TRUE, ncol = 2)


setColWidths(wb, sheet = 1, cols = colw[,1], widths = colw[,2])

writeData(wb, sheet, data, rowNames = FALSE)

addStyle(wb, sheet, style = csb, rows = 1, cols = 1:32)

# Auto Columns Widths
setColWidths(wb, sheet = 1, cols = c(1:4, 6:12, 14:16, 18, 20, 22, 26, 28, 30), widths = "auto")

addStyle(wb, sheet, style = cswrap, rows = 2:nrow(data), cols = c(19, 25), gridExpand = TRUE)
setHeaderFooter(wb, sheet, header = c("Study", filename, date), footer = c(NA, NA, "&[Page] of &[Pages]"))
outfile <- paste0(filename, "_convert.xlsx")
saveWorkbook(wb, paste0("U:/Study/", outfile), overwrite = TRUE)

beepr::beep()
}
