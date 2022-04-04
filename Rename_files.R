library("magrittr")
library("icesTAF")
options(warn = 2)

filepath <- "U:/Study"
siteno <- "108"

message("Deleting all empty directories")
icesTAF::rmdir(filepath, recursive = TRUE)
beepr::beep()

message("Putting in generic names")
{files <- list.files(filepath, pattern = "^[0-9]{3}-[0-9]{3}")
names <- gsub(x = files, pattern = ".pdf", replacement = "")
names <- paste0("Subject ", names, " Final Subject Data_", siteno, "_.pdf")

file.rename(from = paste0(filepath, files), to = paste0(filepath, names))}
# Read in PDFs to determine dates
message("Getting dates")
{files <- list.files(filepath, pattern = "_.pdf$")
date <- NA

for(x in 1:length(files)){
  message(x, " of ", length(files))
  text <- pdftools::pdf_ocr_text(paste0(filepath, files[x]), pages = 1, language = "eng", dpi = 600 )
  date1 <- stringr::str_split(text, pattern = "\\n")[[1]][5]
  
    date2 <- stringr::str_split(string = date1, pattern = ": ")[[1]][2]
    date3 <- as.Date(stringr::str_split(string = date2, pattern = " ")[[1]][1], 
                     format = "%d-%b-%Y")
    date %<>% append(format(date3, "%d %b %Y"))
}
date <- date[-1]}
date
beepr::beep()


message("Adding dates to filename")
{names <- gsub(x = files, pattern = ".pdf", replacement = "")
names <- paste0(names, date, ".pdf")
file.rename(from = paste0(filepath, files), to = paste0(filepath, names))}
beepr::beep()

