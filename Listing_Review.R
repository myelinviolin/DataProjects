library("tibble")
library("magrittr")
library("xlsx")
library("lubridate")
library("zoo")
options(warn = 2)

test <- TRUE
# Functions/Lists -----
visit_order <- c("SCREEN", "BASELINE", "DAY_1", "DAY_2", "DAY_14", "DAY_28",
                 "WEEK_8", "WEEK_12", "WEEK_16", "WEEK_20", "WEEK_24", "WEEK_28",
                 "WEEK_32", "WEEK_36", "WEEK_40", "WEEK_44", "WEEK_48", "WEEK_52",
                 "WEEK_56", "WEEK_60", "WEEK_64", "WEEK_68", "WEEK_72", "WEEK_76",
                 "WEEK_80", "WEEK_84", "WEEK_88", "WEEK_92", "WEEK_96", "WEEK_100",
                 "WEEK_104", "UNSCHED", "RECONSENT", "EOS", "F_AE", "F_CM",
                 "F_CMSTER", "F_CP", "F_ET", "F_RETREAT", "AFLIB", "DEATH")
list_aes <- function(sub){
  subae <- ae[ae$Subject == sub,]
  text <- NA
  for(y in 1:nrow(subae)){
    text %<>% paste0(subae$AETERM[y], "/", subae$AELOC[y], " started ",
                                 subae$AESTDAT_RAW[y])
    if(subae$AEONGO[y] == "Yes"){
      text %<>% paste0(" and is ongoing.\n")
    }
    if(subae$AEONGO[y] == "No"){
      text %<>% paste0(" and ended ", subae$AEENDAT_RAW[y], ".\n")
    }
  }
  text %<>% gsub(pattern = "^NA", replacement = "")
  text
}
list_oh <- function(sub){
  suboh <- oh[oh$Subject == sub,]
  text <- NA
  for(y in 1:nrow(suboh)){
    text %<>% paste0(suboh$OHTERM[y], "/", suboh$OHLOC_STD[y], " started ",
                     suboh$OHSTDAT_RAW[y])
    if(suboh$OHONGO[y] == 1){
      text %<>% paste0(" and is ongoing.\n")
    }
    if(suboh$OHONGO[y] == 0){
      text %<>% paste0(" and ended ", suboh$OHENDAT_RAW[y], ".\n")
    }
  }
  text %<>% gsub(pattern = "^NA", replacement = "")
  text
}
list_mh <- function(sub){
  submh <- mh[mh$Subject == sub,]
  text <- NA
  for(y in 1:nrow(submh)){
    text %<>% paste0(submh$MHTERM[y], " started ",
                     submh$MHSTDAT_RAW[y])
    if(submh$MHONGO[y] == 1){
      text %<>% paste0(" and is ongoing.\n")
    }
    if(submh$MHONGO[y] == 0){
      text %<>% paste0(" and ended ", submh$MHENDAT_RAW[y], ".\n")
    }
  }
  text %<>% gsub(pattern = "^NA", replacement = "")
  text
}
UNKN_dates <- function(start, end){
  data[grepl(x = data[,end], pattern = "UNKN"), end] <- format(Sys.Date(), "%d %b %Y") %>% 
    toupper() %>% as.character()
  data[is.na(data[,end]), end] <- format(Sys.Date(), "%d %b %Y") %>% 
    toupper() %>% as.character()
  data[,end] %<>% gsub(pattern = "UNK ", replacement = "DEC ")
  un_dates <- which(grepl(x = data[,end], pattern = "^UN"))
  for(x in un_dates){
    data[x,end] %<>% gsub(pattern = "^UN ", replacement = "01 ")
    data[x,end] <- as.character(data[x,end] %>% 
                                    as.Date(format = "%d %b %Y") %>%
                                    lubridate::ceiling_date(unit = "month") - days(1))
    data[x,end] <- format(as.Date(data[x,end], format("%Y-%m-%d")), "%d %b %Y") %>% 
      toupper()
  }  
  for(x in 1:nrow(data)){
    if(!is.na(data[x,start])){
      # If start year is unknown, make start date their birthday
      if(grepl(data[x,start], pattern = "UNKN") | data[x,start] == "UN UNK UNK"){
        if(data$Subject[x] %in% dm$Subject){
          if(dm$BRTHDAT_RAW[dm$Subject == data$Subject[x]] != ""){
            data[x,start] <- dm$BRTHDAT_RAW[dm$Subject == data$Subject[x]]
          }else{
            data[x,start] <- "MISSING"
          }
        }else{
          data[x,start] <- "MISSING"
        }
      }
      # If start month is unknown, make it JAN.  If start date is unknown, make it the first.
      data[x,start] %<>% gsub(pattern = "UNK", replacement = "JAN") %>%
        gsub(pattern = "^UN", replacement = "01")
    }else{
      if(data$Subject[x] %in% dm$Subject){
        if(dm$BRTHDAT_RAW[dm$Subject == data$Subject[x]] != ""){
          data[x,start] <- dm$BRTHDAT_RAW[dm$Subject == data$Subject[x]]
        }else{
          data[x,start] <- "MISSING"
        }
      }else{
        data[x,start] <- "MISSING"
      }
    }
  }
  data
}

# Read in files -----
message("Reading in files ...")
if(!test){
  folderpath <- "U:/Study"
}else{folderpath <- "U:/Study_test"}

csv_files <- list.files(path = folderpath, pattern = ".CSV")

for(x in 1:length(csv_files)){
  assign(x = gsub(x = csv_files[x], pattern = ".CSV", replacement = ""), 
         value = read.csv(file = paste0(folderpath,csv_files[x]), sep = ",", 
                          stringsAsFactors = FALSE))
}
beepr::beep()

# Demographics -----
message("Demographics ... Calculating")

data <- dm %>% dplyr::select(-c(Site, siteid)) %>% unique()

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

# Put in  flags
data$RACE_OTHER_SPECIFY[1] <- "HAWAIIAN"
data$RACE_OTHER_SPECIFY[2] <- "PUERTO RICAN"
race_opts <- c("HAWAIIAN", "AMERICAN", "ASIAN", "BLACK", "WHITE")

message("Check these Race Other, Specify fields for Nationality or present Race options.")
unique(data$RACE_OTHER_SPECIFY)
nationalities <- c("PUERTO RICAN")
message("Already in nationalities list:")
nationalities
message("Race options:")
race_opts

for(x in 1:nrow(data)){
  ### Query if provided other, specify value is a value already listed for Race. -----
  if(data$RACE_OTHER_SPECIFY[x] %in% race_opts){
    data$Comments[x] %<>% paste0("Specify race is already listed.\n")
    data$Query[x] %<>% paste0(data$RACE_OTHER_SPECIFY[x], " is already an option for Race. Please select that box instead of using the Race Other, Specify free text field.\n")
  }
  ### Query if provided other, specify for Race is a nationality (e.g., Puerto Rican) or an ethnicity. -----
  if(data$RACE_OTHER_SPECIFY[x] %in% nationalities){
    data$Comments[x] %<>% paste0("Specify race is a nationality.\n")
    data$Query[x] %<>% paste0(data$RACE_OTHER_SPECIFY[x], " is a nationality, not a race. Please reconcile.\n")
  }
}

### Print -----
message("Demographics ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

data %<>% dplyr::select(Subject, ETHNIC_STD, RACE_AMERICAN, 
                 RACE_ASIAN, RACE_BLACK, RACE_HAWAIIAN, RACE_WHITE, RACE_OTHER,
                 RACE_OTHER_SPECIFY, Comments, Query, Reviewer)

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Demographics.xlsx", 
                 sheetName = "Demographics", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()


# Subject Checks -----
message("Subject Checks ... Calculating")
data <- subject %>% dplyr::select(project, Subject, SUBSTAT) %>% unique()

# Pull in Medical  History Y/N
data2 <- dplyr::select(mhyn, Subject, MHYN) %>% unique()
data <- merge(x = data, y = data2, by = c("Subject"), all.x = TRUE)

# Pull in ConMeds Y/N
data3 <- dplyr::select(cmyn, Subject, CMYN) %>% unique()
data <- merge(x = data, y = data3, by = "Subject", all.x = TRUE)

# Pull in Ocular History Y/N
data4 <- dplyr::select(ohyn, Subject, OHYN) %>% unique()
data <- merge(x = data, y = data4, by = "Subject", all.x = TRUE)

# Pull in Date of Administration
data5 <- ex %>% dplyr::select(Subject, EXDAT_RAW) %>% unique()
data <- merge(x = data, y = data5, by = "Subject", all.x = TRUE)

# Pull in Date of Informed Consent
data6 <- ic %>% dplyr::select(Subject, ICDAT_RAW) %>% unique()
data <- merge(x = data, y = data6, by = "Subject", all.x = TRUE)

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Put in flag triggers -----
if(test){
  # Have a subject with a MHYN answer of No
  data$MHYN[1] <- "No"
  # Have a subject with a CMYN answer of No
  data$CMYN[2] <- "No"
  # Have a subject with a OHYN answer of No
  data$OHYN[3] <- "No"
}
### Calculate -----
for(x in 1:nrow(data)){
  ### Query if the subject has was included in study but does not have MHYN filled out. -----
  if(data$SUBSTAT[x] %in% c("RANDOMIZED", "COMPLETED", "EARLY TERMINATED") & 
     is.na(data$MHYN[x])){
    data$Comments[x] %<>% paste0("Should have a MHY/N Answer.\n")
    data$Query[x] %<>% paste0("This subject has a status of ", data$SUBSTAT[x], 
                              ", and the Medical History Y/N form has not been filled out. Please reconcile.\n")
  }
  ### Query if the subject has was included in study but does not have CMYN filled out. -----
  if(data$SUBSTAT[x] %in% c("RANDOMIZED", "COMPLETED", "EARLY TERMINATED") & 
     is.na(data$CMYN[x])){
    data$Comments[x] %<>% paste0("Should have a CMY/N Answer.\n")
    data$Query[x] %<>% paste0("This subject has a status of ", data$SUBSTAT[x], 
                              ", and the Concomitant Medication Y/N form has not been filled out. Please reconcile.\n")
  }
  ### Query if the subject has was included in study but does not have OHYN filled out. -----
  if(data$SUBSTAT[x] %in% c("RANDOMIZED", "COMPLETED", "EARLY TERMINATED") & 
     is.na(data$OHYN[x])){
    data$Comments[x] %<>% paste0("Should have a OHY/N Answer.\n")
    data$Query[x] %<>% paste0("This subject has a status of ", data$SUBSTAT[x], 
                              ", and the Ocular History Y/N form has not been filled out. Please reconcile.\n")
  }
  ### Query for confirmation, if subject has no medical history -----
  if(!is.na(data$MHYN[x])){
    if(data$MHYN[x] == "No"){
      data$Comments[x] %<>% paste0("MHY/N is No.\n")
      data$Query[x] %<>% paste0("Please confirm that the subject has no Medical History.\n")
    }
    ### Query if the subject does not have a status but MHYN is filled out. -----
    if(is.na(data$SUBSTAT[x])){
      data$Comments[x] %<>% paste0("MHY/N is filled out with no subject status. Check subject for errors.\n")
    }
  }
  ### Query for confirmation, if subject has no ConMeds -----
  if(!is.na(data$CMYN[x])){
    if(data$CMYN[x] == "No"){
      data$Comments[x] %<>% paste0("CMY/N is No.\n")
      data$Query[x] %<>% paste0("Please confirm that the subject has no Concomitant Medications.\n")
    }
    ### Query if the subject does not have a status but CMYN is filled out. -----
    if(is.na(data$SUBSTAT[x])){
      data$Comments[x] %<>% paste0("CMY/N is filled out with no subject status. Check subject for errors.\n")
    }
  }
  ### Query for confirmation, if subject has no ocular history. -----
  if(!is.na(data$OHYN[x])){
    if(data$OHYN[x] == "No"){
      data$Comments[x] %<>% paste0("OHY/N is No.\n")
      data$Query[x] %<>% paste0("Please confirm that the subject has no Ocular History.\n")
    }
    ### Query if the subject does not have a status but OHYN is filled out. -----
    if(is.na(data$SUBSTAT[x])){
      data$Comments[x] %<>% paste0("OHY/N is filled out with no subject status. Check subject for errors.\n")
    }
  }
  ### Query if the subject received the study drug before Informed Consent -----
  ic_date <- as.Date(data$ICDAT_RAW[x], "%d %b %Y")
  admin_date <- as.Date(data$EXDAT_RAW[x], "%d %b %Y")
  if(!is.na(admin_date) & !is.na(ic_date)){
    if(admin_date < ic_date){
      data$Comments[x] %<>% paste0("Admin < IC.\n")
      data$Query[x] %<>% paste0("The study drug administration date is before the Informed Consent date. Please reconcile.\n")
      
    }
  }
}

### Print -----
message("Subject Checks ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

data %<>% dplyr::select(project, Subject, Site, siteid, SUBSTAT, MHYN, OHYN, CMYN,
                        Comments, Query, Reviewer)

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Subject_Checks.xlsx", 
                 sheetName = "Subject Checks", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()


# Inclusion/Exclusion -----
message("Inclusion/Exclusion ... Calculating")

data <- ie
data %<>% dplyr::select(Subject, Folder, IEYN, IECAT_STD, 
                        IETESTCD) %>% unique()
data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

for(sub in unique(data$Subject)){
  subject <- data[data$Subject == sub,]
  ### Query if Criterion Type is provided but no Criterion Identifier is listed.
  for(x in 1:nrow(subject)){
    if(!is.na(subject$IECAT_STD[x]) & is.na(subject$IETESTCD[x])){
      data$Comments[data$Subject == sub][x] %<>% paste0("IETESTCD missing.")
      data$Query[data$Subject == sub][x] %<>% paste0("Criterion Type is entered, but no Criterion Identifier is listed. Please reconcile.")
    }
  }
  
  ### Query if No is selected for "Did the subject meet all eligibility criteria?" and form is not filled out. -----
  if(subject$IEYN[x] == "No" & is.na(subject$IECAT_STD[x])){
    data$Comments[data$Subject == sub][x] %<>% paste0("IEYN is No, but no I/E entries.")
    data$Query[data$Subject == sub][x] %<>% paste0("Subject did not meet all eligibility criteria, but no Inclusion/Exclusion criteria entered. Please reconcile.")
  }
}

### Print -----
message("Inclusion/Exclusion ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Inclusion_Exclusion.xlsx", 
                 sheetName = "IE", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()








# Enrollment -----
message("Enrollment ... Calculating")

data <- enroll
data %<>% dplyr::select(Subject, Folder, DSYN_ENROLL, 
                        SFREAS_STD, SFREAS_SPECIFY, PHASE_STD) %>% unique()
### Bring in Blood Pressure -----
bp <- vslog
bp$VSTPT_STD[bp$VSTPT_STD == ""] <- NA
bp %<>% 
  # We're only interested in the readings at Screening
  dplyr::filter(Folder == "SCREEN") %>% 
  # Pare down columns
  dplyr::select(Subject, VSTPT_STD, VSBPSYS, VSBPDIA) %>% 
  # Remove lines with no bp readings
  dplyr::filter(!(is.na(VSBPSYS) & is.na(VSBPDIA))) %>%
  # Make each Reading into a column.
  tidyr::gather(key = BP, value = mmHg, VSBPSYS:VSBPDIA) %>%
  tidyr::unite(col = READING, BP, VSTPT_STD, sep = "_")

bp$READING %<>% gsub(pattern = " ", replacement = "_")

bp %<>%
  tidyr::spread(key = READING, value = mmHg) %>%
  dplyr::select(Subject, VSBPSYS_READING_1, VSBPSYS_READING_2, VSBPSYS_READING_3,
                VSBPDIA_READING_1, VSBPDIA_READING_2, VSBPDIA_READING_3)

bp$SYS_AVG <- NA
bp$DIA_AVG <- NA

# Calculate averages
for(x in 1:nrow(bp)){
  bp$SYS_AVG[x] <- mean(c(bp$VSBPSYS_READING_1[x], 
                          bp$VSBPSYS_READING_2[x], 
                          bp$VSBPSYS_READING_3[x]),
                        na.rm = TRUE) %>% round(digits = 0)
  bp$DIA_AVG[x] <- mean(c(bp$VSBPDIA_READING_1[x], 
                          bp$VSBPDIA_READING_2[x], 
                          bp$VSBPDIA_READING_3[x]),
                        na.rm = TRUE) %>% round(digits = 0)
}

# Combine with data
data %<>%  merge(bp, all.x = TRUE)

### Bring in Anti-VEGF Injections -----
inj <- dplyr::select(cm_avegf, Subject, CMLOC_AVEGF_STD)
colnames(inj) %<>% dplyr::recode("CMLOC_AVEGF_STD" = "INJEYE")
inj$INJEYE[inj$INJEYE == ""] <- NA

# Get rid of rows where the injected eye is not listed
inj %<>% dplyr::filter(!is.na(inj$INJEYE)) %>% table() %>% as.data.frame() %>%
  tidyr::spread(key = INJEYE, value = Freq)
colnames(inj) %<>% dplyr::recode("OD" = "INJ_OD",
                                 "OS" = "INJ_OS")

data %<>% merge(inj, all.x = TRUE)
data$INJ_OD[is.na(data$INJ_OD)] <-  0
data$INJ_OS[is.na(data$INJ_OS)] <-  0

### Bring in Study Eye -----
se <- ds_eye %>% dplyr::select(Subject, DSEYE_STD)
colnames(se) %<>% dplyr::recode("DSEYE_STD" = "SE")
data %<>% merge(se, all.x = TRUE)


### Put in flagging data -----
# SYS is too high, DIA is fine
data$DSYN_ENROLL[1] <- "Yes"
data$SYS_AVG[1] <- 160
data$DIA_AVG[1] <- 70
# DIA is too high, SYS is fine
data$DSYN_ENROLL[2] <- "Yes"
data$SYS_AVG[2] <- 150
data$DIA_AVG[2] <- 100
# Both are too high
data$DSYN_ENROLL[3] <- "Yes"
data$SYS_AVG[3] <- 160
data$DIA_AVG[3] <- 100
# Both are too high, but subject not enrolled
data$DSYN_ENROLL[4] <- "No"
data$SYS_AVG[4] <- 160
data$DIA_AVG[4] <- 100
# 6 injections in study eye
data$DSYN_ENROLL[5] <- "Yes"
data$SE[5] <- "OD"
data$INJ_OD[5] <- 6
# 6 injections in fellow eye, none in study eye
data$DSYN_ENROLL[6] <- "Yes"
data$SE[6] <- "OD"
data$INJ_OD[6] <- 0
data$INJ_OS[6] <- 6

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Calculate -----
for(x in 1:nrow(data)){
  if(data$DSYN_ENROLL[x] == "Yes"){
    ### Query if the subject is enrolled and a study eye is not chosen -----
    if(is.na(data$SE[x])){
      data$Comments[x] %<>% paste0("SE not selected for enrollment.\n")
      data$Query[x] %<>% paste0("The study eye was not selected at screening, yet the subject was enrolled. Please reconcile.\n")
    }else{
      ### Query if the the subject is enrolled and anti-vegf page does not have at least 6 injections in the study eye -----
      if((data$SE[x] == "OD" & data$INJ_OD[x] < 6) || (data$SE[x] == "OS" & data$INJ_OS[x] < 6)){
        data$Comments[x] %<>% paste0("SE has < 6 VEGF injections.\n")
        data$Query[x] %<>% paste0("The study eye has fewer than 6 Anti-VEGF injections, yet the subject was enrolled. Please reconcile.\n")
      }
    }
    ### Query if the the subject is enrolled and Screening Blood pressure measurements were not taken -----
    if(is.na(data$SYS_AVG[x]) || is.na(data$DIA_AVG[x])){
      data$Comments[x] %<>% paste0("BP not measured for enrollment.\n")
      data$Query[x] %<>% paste0("The subject's blood pressure was not measured at screening, yet the subject was enrolled. Please reconcile.\n")
    }else{
      ### Query if the the subject is enrolled and Screening Blood pressure average of triplicate >= 160 mmHg systolic or >= 100mmHg diastolic -----
      if(data$SYS_AVG[x]>=160 || data$DIA_AVG[x]>=100){
        data$Comments[x] %<>% paste0("BP too high for enrollment.\n")
        if(data$SYS_AVG[x]>=160){
          data$Query[x] %<>% paste0("The subject's systolic blood pressure was ", data$SYS_AVG[x] ," at screening, yet the subject was enrolled. Please reconcile.\n")
        }
        if(data$DIA_AVG[x]>=100){
          data$Query[x] %<>% paste0("The subject's diastolic blood pressure was ", data$DIA_AVG[x] ," at screening, yet the subject was enrolled. Please reconcile.\n")
        }
      }
    }
  }
}
### Print -----
message("Enrollment ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Enrollment.xlsx", 
                 sheetName = "Enrollment", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()








# Ocular/Medical History -----
message("Ocular/Medical History ... Calculating")

data_oh <- oh %>% dplyr::select(Subject, OHTERM, OHLOC_STD,
                        OHSTDAT_RAW, OHENDAT_RAW, DataPageName) %>% unique()
colnames(data_oh) %<>% dplyr::recode("OHTERM" = "TERM",  
                                     "OHLOC_STD" = "OHLOC",
                                     "OHSTDAT_RAW" = "STDAT_RAW",
                                     "OHENDAT_RAW" = "ENDAT_RAW",
                                     "DataPageName" = "FORM")

data_mh <- mh %>% dplyr::select(Subject, MHTERM,
                                MHSTDAT_RAW, MHENDAT_RAW, DataPageName) %>% unique()
colnames(data_mh) %<>% dplyr::recode("MHTERM" = "TERM",  
                                     "MHSTDAT_RAW" = "STDAT_RAW",
                                     "MHENDAT_RAW" = "ENDAT_RAW",
                                     "DataPageName" = "FORM")
data_mh$OHLOC <- NA
data <- rbind(data_oh, data_mh)
data$FORM %<>% dplyr::recode("Ocular Medical History" = "OH",
                             "Non-Ocular Medical History" = "MH")

# Pull in date of discontinuation
data3 <- dplyr::select(ds_eos, Subject, DSSTDAT_RAW) %>% unique()
data <- merge(x = data, y = data3, by = "Subject", all.x = TRUE)

colnames(data) %<>% dplyr::recode("DSSTDAT_RAW" = "DSSTDAT")

data$ENDAT <- data$ENDAT_RAW

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

ocular_conds <- c("DRY EYE", "EYE SIDE EDEMA", "LOST SIGHT", "RED EYES")

### Replace all UNKs - push all end dates latest -----
data %<>% add_column(STDAT = data$STDAT_RAW, .after = "STDAT_RAW")
data %<>% add_column(ENDAT = data$ENDAT_RAW, .after = "ENDAT_RAW")

data <- UNKN_dates("STDAT", "ENDAT")

### Put in flagging data -----
# No end date but subject is discontinued
data$ENDAT[4] <-  NA
data$DSSTDAT[4] <- "01 JAN 2000"

# Ocular term under MH
data$TERM[5] <-  "DRY EYE"
data$FORM[5] <- "MH"

### Calculate -----
for(x in 1:nrow(data)){
  ### Query if OHLOC is not selected -----
  if(is.na(data$OHLOC[x]) & data$FORM[x] == "OH"){
    data$Comments[x] %<>% paste0("No OHLOC.\n")
    data$Query[x] %<>% paste0("The eye location has not been selected. Please reconcile.\n")
  }
  start <- as.Date(data$STDAT[x], format ="%d %B %Y")
  end <- as.Date(data$ENDAT[x], format ="%d %B %Y")
  discon <- as.Date(data$DSSTDAT[x], format ="%d %B %Y")

  ### Query if the there is no end date but the subject is discontinued -----
  if(is.na(end) & !is.na(discon)){
    data$Comments[x] %<>% paste0("Discontinued, but no end date\n")
    data$Query[x] %<>% paste0("The subject is discontinued, but the ocular history has no end date. Please reconcile.\n")
  }
  
  ### Query if End Date is after the Date of Discontinuation. -----
  if(!is.na(end) & !is.na(discon)){
    if(end > discon){
      data$Comments[x] %<>% paste0("End date after discontinuation.\n")
      data$Query[x] %<>% paste0("The end date occurs after the date of discontinuation for this subject. Please reconcile.\n")
    }
  }
  
  ### Query if non/ocular disease/condition is present. --------------------
  if(!data$TERM[x] %in% ocular_conds & data$FORM[x] == "OH"){
    data$Comments[x] %<>% paste0("Potentially non-ocular condition\n")
    data$Query[x] %<>% paste0("Please confirm if this condition is non-ocular, or update the condition to be ocular-specific.\n")
  }
  if(data$TERM[x] %in% ocular_conds & data$FORM[x] == "MH"){
    data$Comments[x] %<>% paste0("Potentially ocular condition\n")
    data$Query[x] %<>% paste0("Please confirm if this condition is ocular. If so, it belongs under Ocular History.\n")
  }
  
  ### Query if a surgery is present that a corresponding condition is not also recorded -----
  
  
  
}

### Print -----
message("Ocular/Medical History ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

data %<>% dplyr::select(Subject, FORM, TERM, OHLOC, 
                        STDAT_RAW, ENDAT_RAW, ENDAT, DSSTDAT, Comments, Query, 
                        Reviewer)


xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Medical_Ocular_History.xlsx", 
                 sheetName = "MHOH", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()

# Discontinuation -----
message("Discontinuation ... Calculating")

data <- ds_eos

data %<>% dplyr::select(Subject, DSSTDAT_RAW, 
                        DSYN_COMPLETE_STD, DSDECOD_STD, DSCOVAL, DSAEID) %>% 
  unique()
colnames(data) %<>% dplyr::recode("DSSTDAT_RAW" = "DSSTDAT",
                                  "DSYN_COMPLETE_STD" = "DSYN_COMPLETE",
                                  "DSDECOD_STD" = "DSDECOD")

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

reasons <- c("DEATH", "ADVERSE EVENT", "LOST TO FOLLOW-UP", "PHYSICIAN DECISION", "PROTOCOL DEVIATION", "STUDY TERMINATED", "WITHDRAWAL BY SUBJECT", "OTHER")

### Put in flagging data -----
data$DSAEID[1] <- 456
data$DSDECOD[2] <- "DEATH"
data$DSAEID[2] <- NA
data$DSDECOD[2] <- "ADVERSE EVENT"
data$DSAEID[2] <- NA
data$DSDECOD[3] <- "OTHER"
data$DSCOVAL[3] <- "LOST TO FOLLOW-UP"


### Calculate -----
for(x in 1:nrow(data)){
  if(!is.na(data$DSDECOD[x])){
    ### Query if "The subject was discontinued due to:" is "Death" or "Adverse Event" and there is no corresponding AE. -----
    if((data$DSDECOD[x] == "DEATH" || data$DSDECOD[x] == "ADVERSE EVENT") & is.na(data$DSAEID[x])){
      data$Comments[x] %<>% paste0("No AE recorded for Reason DEATH/AE.\n")
      data$Query[x] %<>% paste0("No AE recorded for a Discontinuation Reason of ", data$DSDECOD[x], ". Please reconcile.\n")
    }
  }
  ### Query if provide details regarding discontinuation information is a value already listed for "The subject was discontinued due to." -----
  if(data$DSCOVAL[x] %in% reasons){
    data$Comments[x] %<>% paste0("OTHER value already listed.\n")
    data$Query[x] %<>% paste0(data$DSCOVAL[x], " is already an option for the reason for discontinuation. Please reconcile.\n")
  }
  ### Make sure corresponding AE makes sense if listed. -----
  if(!is.na(data$DSAEID[x])){
    data$Comments[x] %<>% paste0(list_aes(data$Subject[x]))
  }
}
### Print -----
message("Discontinuation ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Discontinuation.xlsx", 
                 sheetName = "DS", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()



# 4D Administration -----
data <- ex
# Bring in study eye
se <- ds_eye %>% dplyr::select(Subject, DSEYE_STD)
colnames(se) %<>% dplyr::recode("DSEYE_STD" = "SE")
data %<>% merge(se, all.x = TRUE)

data %<>% dplyr::select(Subject, EXPERF, EXEYE_STD, SE,
                        EXDAT_RAW, EXTIM, EXYN_COMP, EXAEID) %>% unique()

colnames(data) %<>% dplyr::recode("EXEYE_STD" = "EXEYE",
                                  "EXDAT_RAW" = "EXDAT")
data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Put in flagging data -----
data$EXYN_COMP[1] <- "Yes"
data$EXAEID[1] <- NA

### Calculate -----
for(x in 1:nrow(data)){
  ### Query if not injected in study eye -----
  if(!is.na(data$EXEYE[x]) & !is.na(data$SE[x])){
    if(data$EXEYE[x] != data$SE[x]){
      data$Comments[x] %<>% paste0("Injected eye not same as study eye.\n")
      data$Query[x] %<>% paste0("The injection was performed on ", data$EXEYE[x],
                                ", but the study eye for this subject is ", 
                                data$SE[x],". Please reconcile.\n")
    }
  }
  ### Query if there are complications but no AE listed. -----
  if(!is.na(data$EXYN_COMP[x])){
    if(data$EXYN_COMP[x] == "Yes" & is.na(data$EXAEID[x])){
      data$Comments[x] %<>% paste0("AE needs to be listed.\n")
      data$Query[x] %<>% paste0("There were complications following injection, but the corresponding AE is not listed. Please reconcile.\n")
    }
  }
  ### List AE that is provided. -----
  if(!is.na(data$EXAEID[x])){
    if(data$Subject[x] %in% ae$Subject){
      data$Comments[x] %<>% paste0(list_aes(sub = data$Subject[x]))
    }else{
      data$Comments[x] %<>% paste0("No AEs listed for subject.\n")
      data$Query[x] %<>% paste0("No AEs are listed for subject, yet an AE was referenced due to the complication with injection. Please reconcile.\n")
    }
  }
}
### Print -----
message("4D Administration ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/4D_Admin.xlsx", 
                 sheetName = "4D", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()


# IOP -----
data_pre <- oe_iop
data_pre %<>% dplyr::select(project, Subject, Site, siteid, Folder, OEEYE_STD, OESTAT,
                        OEDAT_RAW, OEORRES_IOP, OEMETHOD_IOP_STD, 
                        OEMETHOD_IOP_SPECIFY)

data_post <- oe_iop_po
data_post %<>% dplyr::select(project, Subject, Site, siteid, Folder, OEEYE_STD, 
                             OEPERF_IOP, OEORRES_IOP, OEMETHOD_IOP_STD, 
                            OEMETHOD_IOP_SPECIFY)
data <- merge(data_pre, data_post, 
              by = c("project", "Subject", "Site", "siteid", "Folder", "OEEYE_STD"), 
              all.x = TRUE, all.y = TRUE, suffixes = c("_PRE", "_POST"))

colnames(data) %<>% dplyr::recode("OEPERF_IOP" = "OEPERF_IOP_POST")

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())


### Print -----
message("IOP ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/IOP.xlsx", 
                 sheetName = "IOP", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()

# Slit Lamp -----
data <- oe_sl
colnames(data) %<>% 
  gsub(pattern = "OEORRES_", replacement = "") %>%
  gsub(pattern = "OEDESC", replacement = "D")
  
data %<>% dplyr::select(Subject, Folder, OEYN_CHANGES_OD,
                        OEYN_CHANGES_OS, OEEYE_STD, OEDAT_RAW, EYELID_STD,
                        D_EYELID, EYELID_ERYTHEMA, EYELID_EDEMA,
                        CONJ_STD, D_CONJ, CONJ_ERYTHEMA,
                        CONJ_EDEMA, SCLERA_STD, D_SCLERA,
                        CORNEA_STD, D_CORNEA, CORNEA_EDEMA,
                        CORNEA_STAINING, ANTCHM_STD, D_ANTCHM,
                        ANTCHM_CELLS, ANTCHM_CELLS_HYPO,
                        OESTAT_ANTCHM_CELLS_HYPO, ANTCHM_FLARE,
                        IRISPUP_STD, D_IRISPUP, LENS_STATUS_STD,
                        OEPERF_AREDS, LENS_STD, D_LENS, ANTVIT_STD,
                        D_ANTVIT)
colnames(data) %<>% 
  gsub(pattern = "_STD", replacement = "") %>%
  gsub(pattern = "_RAW", replacement = "")

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Put in flagging data -----
# data$Folder <- "SCREEN"
data$EYELID[3] <- "ABNORMAL (NCS)"
### Calculate -----
for(x in 1:nrow(data)){
  ### Check for the corresponding condition in MH/OH for all abnormal findings at Screening. Query if the corresponding condition is missing in MH/OH -----
  if(data$Folder[x] == "SCREEN" & 
     (grepl(pattern = "ABNORMAL", x = paste0(data$EYELID[x], data$CONJ[x],
                                            data$SCLERA[x], data$CORNEA[x],
                                            data$ANTCHM[x], data$IRISPUP[x],
                                            data$LENS[x], data$ANTVIT[x])) ||
     grepl(pattern = "DEFECT", x = data$IRISPUP[x]))){
    
    data$Comments[x] %<>% paste0("Abnormality at SCREEN.\n")
    
    if(!data$Subject[x] %in% oh$Subject){
      data$Comments[x] %<>% paste0("No OH present for subject.\n")
    }else{
      data$Comments[x] %<>% paste0(list_oh(data$Subject[x]))
    }
    if(!data$Subject[x] %in% mh$Subject){
      data$Comments[x] %<>% paste0("No MH present for subject.\n")
    }else{
      data$Comments[x] %<>% paste0(list_mh(data$Subject[x]))
    }
    # Subject does not have MH or OH
    if(!data$Subject[x] %in% mh$Subject & !data$Subject[x] %in% oh$Subject){
      data$Query[x] %<>% paste0("There is a significant abnormality at Screening, but the subject does not have any medical or ocular history. Please add OH/MH records as applicable, or otherwise reconcile.\n")
    }
  }
}

for(sub in unique(data$Subject)){
 
  subject <- data[data$Subject == sub & !is.na(data$OEEYE) & !is.na(data$OEDAT),
                  c("Subject", "Folder", "OEEYE",
                    "OEDAT", "EYELID", "CONJ", "SCLERA", "CORNEA", "ANTCHM", 
                    "IRISPUP", "LENS", "ANTVIT")]

  od <- subject[subject$OEEYE == "OD",]
  od <- od[order(as.Date(od$OEDAT, format = "%d %b %Y")),]
  os <- subject[subject$OEEYE == "OS",]
  os <- os[order(as.Date(os$OEDAT, format = "%d %b %Y")),]
  
  for(lat in list(od, os)){
    assign(x = "eye", value = lat)
    change <- FALSE
    if(nrow(eye)>1){
      for(x in 2:nrow(eye)){
        for(area in c("EYELID", "CONJ", "SCLERA", "CORNEA", "ANTCHM", "IRISPUP", 
                      "LENS", "ANTVIT")){
          if(!is.na(eye[x-1,area]) & !is.na(eye[x,area])){
            ### Query if Abnormal Findings increased from Normal to Abnormal (CS) or Abnormal (NCS) to Abnormal (CS) when compared to previous visit and no corresponding AE is recorded -----
            if((eye[x-1,area] == "NORMAL" & 
                (grepl(eye[x,area], pattern = "ABNORMAL")||grepl(eye[x,area], pattern = "DEFECT")))
               ||
               (eye[x-1,area] == "ABNORMAL (NCS)" & eye[x,area] == "ABNORMAL (CS)")
            ){
              change <- TRUE
              data$Comments[data$Subject == sub & 
                              data$Folder == eye$Folder[x] &
                              data$OEEYE == eye$OEEYE[x]] %<>% 
                paste0(area, " worsens from ", eye[x-1, area], " to ", eye[x, area], ".\n")
            }
            ### Query if Abnormal Finding decreased from Abnormal (CS) to Normal/Abnormal (NCS) and there is no end date for the corresponding condition -----
            if((eye[x,area] == "NORMAL" & 
                (grepl(eye[x-1,area], pattern = "ABNORMAL")||grepl(eye[x-1,area], pattern = "DEFECT")))
               ||
               (eye[x,area] == "ABNORMAL (NCS)" & eye[x-1,area] == "ABNORMAL (CS)")
            ){
              change <- TRUE
              data$Comments[data$Subject == sub & 
                              data$Folder == eye$Folder[x] &
                              data$OEEYE == eye$OEEYE[x]] %<>% 
                paste0(area, " improves from ", eye[x-1, area], " to ", eye[x, area], ".\n")
            }
          }
        }
      }
      if(change){
        if(sub %in% ae$Subject){
          data$Comments[data$Subject == sub & 
                          data$Folder == eye$Folder[x] &
                          data$OEEYE == eye$OEEYE[x]] %<>% paste0(list_aes(sub))
        }else{
          data$Comments[data$Subject == sub & 
                          data$Folder == eye$Folder[x] &
                          data$OEEYE == eye$OEEYE[x]] %<>% paste0("Subject has no AEs.\n")
        }
      }
    }
  }
}








### Print -----
message("Slit Lamp ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Slit_Lamp.xlsx", 
                 sheetName = "SL", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()


# Dilated Indirect Ophthalmoscopy -----
data <- oe_dio
colnames(data) %<>% 
  gsub(pattern = "OEORRES_", replacement = "") %>%
  gsub(pattern = "OEDESC", replacement = "D") %>%
  gsub(pattern = "DIO_", replacement = "")

data %<>% dplyr::select(Subject, Folder, OEYN_CHANGES_OD,
                        OEYN_CHANGES_OS, OEEYE_STD, OEDAT_RAW, OPTNERV_STD,
                        D_OPTNERV, MACULA_STD, D_MACULA, RETINA_STD, D_RETINA,
                        CHOROID_STD, D_CHOROID, RETINA_PERI_STD, D_RETINA_PERI,
                        VITHAZE_STD, D_VITHAZE, VITCELLS_STD, D_VITCELLS,
                        VITREOUS_STD, D_VITREOUS, VITBODY_STD,
                        D_VITBODY, FUNDHEM_STD, D_FUNDHEM) %>% unique()
colnames(data) %<>% 
  gsub(pattern = "_STD", replacement = "") %>%
  gsub(pattern = "_RAW", replacement = "")

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Put in flagging data -----
# data$Folder <- "SCREEN"
# data$EYELID[3] <- "ABNORMAL (NCS)"
### Calculate -----
for(x in 1:nrow(data)){
  ### Check for the corresponding condition in MH/OH for all abnormal findings at Screening. Query if the corresponding condition is missing in MH/OH -----
  if(data$Folder[x] == "SCREEN" & !all(is.na(data$OPTNERV[x]), is.na(data$MACULA[x]),
                                       is.na(data$RETINA[x]), is.na(data$CHOROID[x]),
                                       is.na(data$RETINA_PERI[x]), is.na(data$VITHAZE[x]),
                                       is.na(data$VITCELLS[x]), is.na(data$VITREOUS[x]),
                                       is.na(data$VITBODY[x]), is.na(data$FUNDHEM[x])
                                       )){
    if(grepl(pattern = "ABNORMAL", x = paste0(data$OPTNERV[x], data$MACULA[x],
                                              data$RETINA[x], data$CHOROID[x],
                                              data$RETINA_PERI[x], data$VITHAZE[x],
                                              data$VITCELLS[x], data$VITREOUS[x],
                                              data$VITBODY[x], data$VITBODY[x]))){

      data$Comments[x] %<>% paste0("Abnormality at SCREEN.\n")
      
      if(!data$Subject[x] %in% oh$Subject){
        data$Comments[x] %<>% paste0("No OH present for subject.\n")
      }else{
        data$Comments[x] %<>% paste0(list_oh(data$Subject[x]))
      }
      if(!data$Subject[x] %in% mh$Subject){
        data$Comments[x] %<>% paste0("No MH present for subject.\n")
      }else{
        data$Comments[x] %<>% paste0(list_mh(data$Subject[x]))
      }
      # Subject does not have MH or OH
      if(!data$Subject[x] %in% mh$Subject & !data$Subject[x] %in% oh$Subject){
        data$Query[x] %<>% paste0("There is a significant abnormality at Screening, but the subject does not have any medical or ocular history. Please add OH/MH records as applicable, or otherwise reconcile.\n")
      }
    }
  }
}

for(sub in unique(data$Subject)){
  
  subj <- data[data$Subject == sub & !is.na(data$OEEYE) & !is.na(data$OEDAT),
               c("Subject", "Folder", "OEEYE",
                 "OEDAT", "OPTNERV", "MACULA", "RETINA", "CHOROID", "RETINA_PERI",
                 "VITHAZE", "VITCELLS", "VITREOUS", "VITBODY", "FUNDHEM")]
  subj %<>% dplyr::filter(!is.na(OPTNERV), !is.na(MACULA), !is.na(RETINA),
                          !is.na(CHOROID), !is.na(RETINA_PERI), !is.na(VITHAZE),
                          !is.na(VITCELLS), !is.na(VITBODY), !is.na(FUNDHEM))
  
  od <- subj[subj$OEEYE == "OD",]
  od <- od[order(as.Date(od$OEDAT, format = "%d %b %Y")),]
  os <- subj[subj$OEEYE == "OS",]
  os <- os[order(as.Date(os$OEDAT, format = "%d %b %Y")),]
  
  for(lat in list(od, os)){
    assign(x = "eye", value = lat)
    change <- FALSE
    if(nrow(eye)>1){
      for(x in 2:nrow(eye)){
        for(area in c("OPTNERV", "MACULA", "RETINA", "CHOROID", "RETINA_PERI",
                      "VITHAZE", "VITCELLS", "VITREOUS", "VITBODY", "FUNDHEM")){
          if(!is.na(eye[x-1,area]) & !is.na(eye[x,area])){
            ### Query if Abnormal Findings increased from Normal to Abnormal (CS) or Abnormal (NCS) to Abnormal (CS) when compared to previous visit and no corresponding AE is recorded -----
            if((eye[x-1,area] == "NORMAL" & 
                (grepl(eye[x,area], pattern = "ABNORMAL")))
               |
               (eye[x-1,area] == "ABNORMAL (NCS)" & eye[x,area] == "ABNORMAL (CS)")
            ){
              change <- TRUE
              data$Comments[data$Subject == sub & 
                              data$Folder == eye$Folder[x] &
                              data$OEEYE == eye$OEEYE[x]] %<>% 
                paste0(area, " worsens from ", eye[x-1, area], " to ", eye[x, area], ".\n")
            }
            ### Query if Abnormal Finding decreased from Abnormal (CS) to Normal/Abnormal (NCS) and there is no end date for the corresponding condition -----
            if((eye[x,area] == "NORMAL" & 
                (grepl(eye[x-1,area], pattern = "ABNORMAL")))
               |
               (eye[x,area] == "ABNORMAL (NCS)" & eye[x-1,area] == "ABNORMAL (CS)")
            ){
              change <- TRUE
              data$Comments[data$Subject == sub & 
                              data$Folder == eye$Folder[x] &
                              data$OEEYE == eye$OEEYE[x]] %<>% 
                paste0(area, " improves from ", eye[x-1, area], " to ", eye[x, area], ".\n")
            }
          }
        }
      }
      if(change){
        if(sub %in% ae$Subject){
          data$Comments[data$Subject == sub & 
                          data$Folder == eye$Folder[x] &
                          data$OEEYE == eye$OEEYE[x]] %<>% paste0(list_aes(sub))
        }else{
          data$Comments[data$Subject == sub & 
                          data$Folder == eye$Folder[x] &
                          data$OEEYE == eye$OEEYE[x]] %<>% paste0("Subject has no AEs.\n")
        }
      }
    }
  }
}

### Print -----
message("Dilated Indirect Ophthalmoscopy ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/DIO.xlsx", 
                 sheetName = "DIO", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()


# Visual Acuity History -----
message("Visual Acuity History ... Calculating")

data <-  oe_va

data %<>% dplyr::select(project, Subject, Folder, OEEYE_STD,
                        OESTAT_VA, OEDAT_VA_RAW, OEYN_METHOD, OEMETHOD_SPECIFY,
                        OEORRES_SCORE, OEORRES_VA_ALT_STD, OEORRES_FINGERS_COUNT,
                        OEORRES_SNELLEN) %>% unique()
colnames(data) %<>% gsub(pattern = "_RAW", replacement = "") %>%
  gsub(pattern = "OEORRES_", replacement = "") %>%
  gsub(pattern = "_STD", replacement = "") %>%
  gsub(pattern = "_VA", replacement = "")

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Print -----
message("Visual Acuity History ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Visual_Acuity_History.xlsx", 
                 sheetName = "VAH", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()


# ConMeds -----
message("ConMeds ... Calculating")

data <- cm

data %<>% dplyr::select(Subject, CMTRT, CMTRT_PRODUCT,
                        CMLOC_STD, CMINDC_STD, CMINDC_SPECIFY,
                        CMAE_STD, CMMH_STD, CMOH_STD, CMDOSTXT, CMDOSU_STD, 
                        CMDOSU_SPECIFY, CMDOSFRQ_STD, CMDOSFRQ_SPECIFY, CMROUTE_STD, 
                        CMROUTE_SPECIFY, CMSTDAT_RAW, CMENDAT_RAW, CMINDC_SUPP,
                        CMCP_STD) %>% unique()
colnames(data) %<>% gsub(pattern = "_STD", replacement = "")

# Pull in informed consent date
data_ic <- ic %>% dplyr::select(Subject, ICDAT_RAW) %>% unique()
data <- merge(x = data, y = data_ic, all.x = TRUE)
colnames(data) %<>% dplyr::recode("ICDAT_RAW" = "ICDAT")

data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

### Query if the value in the Other field is already provided in the drop down lists for Indication, Dose Units, Frequency, and Route. -----
message("INDICATION OTHER")
cat(unique(data$CMINDC_SPECIFY[data$CMINDC == "OTH"]), sep = "\n")
message("ROUTE OTHER")
cat(unique(data$CMROUTE_SPECIFY[data$CMROUTE == "OTH"]), sep = "\n")
message("DOSE UNITS OTHER")
cat(unique(data$CMDOSU_SPECIFY[data$CMDOSU == "OTH"]), sep = "\n")
message("DOSE FREQ OTHER")
cat(unique(data$CMDOSFRQ_SPECIFY[data$CMDOSFRQ == "OTH"]), sep = "\n")

### Query if Location does not correspond with Indication. -----
test <- data %>% dplyr::select(CMLOC, CMINDC, CMINDC_SPECIFY)
  
test %<>% dplyr::filter(test$CMINDC != "PROPHYLAXIS" & !is.na(test$CMLOC))
test <- test[!(test$CMLOC %in% c("OU", "OD", "OS") & (test$CMINDC == "OCULAR HISTORY")),]
test <- test[!(test$CMINDC == "OTH" & test$CMINDC_SPECIFY == "NA"),]
test <- test[!(test$CMLOC %in% c("OU", "OD", "OS") & test$CMINDC == "MEDICAL HISTORY"),]
test <- test[!(test$CMLOC == "NON-OCULAR" & test$CMINDC == "OCULAR HISTORY"),]
test %<>% unique()

message("Compare Location and Indication:")
knitr::kable(test[,c("CMLOC", "CMINDC", "CMINDC_SPECIFY")])

### Query if Route does not correspond with the Drug Name -----
test <- data %>% dplyr::select(CMTRT, CMDOSU, CMDOSU_SPECIFY, CMROUTE, CMROUTE_SPECIFY) %>%
  dplyr::filter(!is.na(CMTRT)) %>% dplyr::filter(!is.na(CMDOSU)) %>% 
  dplyr::filter(!is.na(CMROUTE))
test <- test[!(test$CMDOSU == "INHALATION" & test$CMROUTE == "INTRAMUSCULAR"),]
test <- test[!(test$CMDOSU == "INHALATION" & test$CMROUTE == "INTRAVENOUS"),]
test <- test[!(test$CMDOSU == "CAPSULE" & test$CMROUTE == "SUBTENON"),]
test <- test[!(test$CMDOSU == "CAPSULE" & test$CMROUTE == "INTRAVENOUS"),]
test <- test[!(test$CMDOSU == "gtt" & test$CMROUTE == "RECTAL"),]
test <- test[!(test$CMDOSU == "APPLICATION" & test$CMROUTE == "INTRAMUSCULAR"),]
test <- test[!(test$CMDOSU == "APPLICATION" & test$CMROUTE == "INTRAVITREAL"),]
test %<>% unique()

message("Compare Drug and Route:")
knitr::kable(test[,c("CMTRT", "CMDOSU", "CMDOSU_SPECIFY", "CMROUTE", "CMROUTE_SPECIFY")])

### Edit start/end dates for calculations -----
data %<>% add_column(CMSTDAT = data$CMSTDAT_RAW, .after = "CMSTDAT_RAW")
data %<>% add_column(CMENDAT = data$CMENDAT_RAW, .after = "CMENDAT_RAW")

data <- UNKN_dates(start = "CMSTDAT", end = "CMENDAT")

### Calculate -----
for(x in 1:nrow(data)){
  ### Specific drop down list errors. -----
  if(!is.na(data$CMROUTE_SPECIFY[x])){
    if(data$CMROUTE_SPECIFY[x] == c("BY MOUTH")){
      data$Comments[x] %<>% paste0("Route already specified.\n")
      data$Query[x] %<>% paste0("Please select ORAL as the Route.\n")
    }
  }
  ### Query if Location does not correspond with Indication. -----
  if(!is.na(data$CMLOC[x]) & !is.na(data$CMINDC[x])){
    if((data$CMLOC[x] %in% c("OU", "OS", "OD") & data$CMINDC[x] == "MEDICAL HISTORY")
       ||
       (data$CMLOC[x] == "NON-OCULAR" & data$CMINDC[x] == "OCULAR HISTORY")
    ){
      data$Comments[x] %<>% paste0("Location does not correspond with Indication.\n")
      data$Query[x] %<>% paste0("Location does not correspond with Indication. Please reconcile.\n")
    }
  }
  ### Query if Route does not correspond with the Drug Name -----
  if(!is.na(data$CMDOSU[x]) & !is.na(data$CMROUTE[x])){
    if((data$CMDOSU[x] == "INHALATION" & data$CMROUTE[x] == "INTRAMUSCULAR")
       ||
       (data$CMDOSU[x] == "INHALATION" & data$CMROUTE[x] == "INTRAVENOUS")
       ||
       (data$CMDOSU[x] == "CAPSULE" & data$CMROUTE[x] == "SUBTENON")
       ||
       (data$CMDOSU[x] == "CAPSULE" & data$CMROUTE[x] == "INTRAVENOUS")
       ||
       (data$CMDOSU[x] == "gtt" & data$CMROUTE[x] == "RECTAL")
       ||
       (data$CMDOSU[x] == "APPLICATION" & data$CMROUTE[x] == "INTRAMUSCULAR")
       ||
       (data$CMDOSU[x] == "APPLICATION" & data$CMROUTE[x] == "INTRAVITREAL")
    ){
      data$Comments[x] %<>% paste0("Drug does not correspond with Route.\n")
      data$Query[x] %<>% paste0("The drug/application does not correspond with the Route. Please reconcile.\n")
    }
  }
  ### Query if CM Indication does not have a corresponding Medical History or Adverse Event record. No query is needed for prophylactic, contraceptive or health maintenance medications. -----
  if(!is.na(data$CMINDC[x])){
    if(data$CMINDC[x] != "PROPHYLAXIS" & 
       (is.na(data$CMAE[x]) & is.na(data$CMMH[x]) & is.na(data$CMOH[x]))
    ){
      data$Comments[x] %<>% paste0("Missing AE/MH/OH association.\n")
      data$Query[x] %<>% paste0("Drug should be associated with an AE, MH, or OH. Please reconcile.\n")
    }
  
  if(!is.na(data$ICDAT[x]) & !is.na(data$CMSTDAT_RAW[x])){
    cm_start <- as.Date(data$CMSTDAT[x], "%d %b %Y")
    ic_dat <- as.Date(data$ICDAT[x], "%d %b %Y")
    ### Query if start date is before Informed Consent, but Indication is listed as ADVERSE EVENT -----
    if(cm_start < ic_dat & data$CMINDC[x] == "ADVERSE EVENT"){
      data$Comments[x] %<>% paste0("Indication AE, but start < IC.\n")
    }
    ### Query if start date is after Informed Consent, but Indication is listed as MH/OH -----
    if(cm_start > ic_dat & 
       (data$CMINDC[x] == "OCULAR HISTORY" | data$CMINDC[x] == "MEDICAL HISTORY")){
      data$Comments[x] %<>% paste0("Indication MH/OH, but start > IC.\n")
    }
  }
  }
}

for(sub in unique(data$Subject)){
  subj <- data[data$Subject == sub,]
  subj %<>% dplyr::filter(!is.na(CMSTDAT) & !is.na(CMENDAT))
  if(nrow(subj) > 0){
    if(any(duplicated(subj$CMTRT))){
      for(x in 1:nrow(subj)){
      cm_start <- as.Date(subj$CMSTDAT[x], "%d %b %Y")
      cm_end <- as.Date(subj$CMENDAT[x], "%d %b %Y")
      cm_span <- interval(start = cm_start, end = cm_end)
      
      others <- subj[-x,]
      other_cm_start <- as.Date(others$CMSTDAT, "%d %b %Y")
      other_cm_end <- as.Date(others$CMENDAT, "%d %b %Y")
      other_cm_span <- interval(start = other_cm_start, end = other_cm_end)
      
      cm_span_overlaps <- lubridate::int_overlaps(cm_span, other_cm_span)
      cm_term_overlaps <- subj$CMTRT[x] == others$CMTRT
      cm_match <- cm_span_overlaps & cm_term_overlaps
      
      data$Comments[data$Subject == sub][x] %<>%
      paste0("CM Term/Date Overlaps:\n", others$CMTRT[cm_match], "/", 
             others$CMSTDAT[cm_match], "---", others$CMENDAT[cm_match], "\n")
      }
    }
  }
}

### Consolidate and print -----
message("ConMeds ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/ConMeds.xlsx", 
                 sheetName = "CM", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()

# AE/CM/MH/OH -----
message("AE/CM/MH/OH ... Calculating")
data_oh <- oh %>% 
  dplyr::select(Subject, OHTERM,
                OHSTDAT_RAW, OHENDAT_RAW, DataPageName) %>% 
  unique()
colnames(data_oh) %<>% dplyr::recode("OHTERM" = "TERM",
                                     "OHSTDAT_RAW" = "STDAT_RAW",
                                     "OHENDAT_RAW" = "ENDAT_RAW",
                                     "DataPageName" = "FORM")
data_mh <- mh %>% 
  dplyr::select(Subject, MHTERM,
                MHSTDAT_RAW, MHENDAT_RAW, DataPageName) %>%
  unique()
colnames(data_mh) %<>% dplyr::recode("MHTERM" = "TERM",  
                                     "MHSTDAT_RAW" = "STDAT_RAW",
                                     "MHENDAT_RAW" = "ENDAT_RAW",
                                     "DataPageName" = "FORM")
data_ae <- ae %>% 
  dplyr::select(Subject, AETERM,
                AESTDAT_RAW, AEENDAT_RAW, DataPageName) %>%
  unique()
colnames(data_ae) %<>% dplyr::recode("AETERM" = "TERM",  
                                     "AESTDAT_RAW" = "STDAT_RAW",
                                     "AEENDAT_RAW" = "ENDAT_RAW",
                                     "DataPageName" = "FORM")
data_cm <- cm %>% 
  dplyr::select(Subject, CMTRT,
                CMSTDAT_RAW, CMENDAT_RAW, CMINDC_STD, DataPageName) %>%
  unique()
colnames(data_cm) %<>% dplyr::recode("CMTRT" = "TERM",  
                                     "CMSTDAT_RAW" = "STDAT_RAW",
                                     "CMENDAT_RAW" = "ENDAT_RAW",
                                     "CMINDC_STD" = "CMINDC",
                                     "DataPageName" = "FORM")
data_oh$CMINDC <- NA
data_mh$CMINDC <- NA
data_ae$CMINDC <- NA

data <- rbind(data_oh, data_mh, data_ae, data_cm)
data$FORM %<>% dplyr::recode("Ocular Medical History" = "OH",
                             "Non-Ocular Medical History" = "MH",
                             "Adverse Events" = "AE",
                             "Prior/Concomitant Medications" = "CM")
# Pull in informed consent date
data_ic <- ic %>% dplyr::select(Subject, ICDAT_RAW) %>% unique()
data <- merge(x = data, y = data_ic, all.x = TRUE, by = "Subject")

# Pull in CM Links to AE/MH/OH
data_link <- cm %>% dplyr::select(Subject, CMTRT, CMAE, CMMH, CMOH, CMSTDAT_RAW,
                                CMENDAT_RAW) %>% unique()
colnames(data_link) %<>% dplyr::recode("CMTRT" = "TERM",
                                       "CMSTDAT_RAW" = "STDAT_RAW",
                                       "CMENDAT_RAW" = "ENDAT_RAW")

data_link %<>% tidyr::unite(link, c(CMAE, CMMH, CMOH), sep = "", na.rm = TRUE)
data_link$link %<>% gsub(pattern = "^.*-.*- ", replacement = "")
data <-  merge(x = data, y = data_link, all.x = TRUE, 
               by = c("Subject", "TERM", "STDAT_RAW", "ENDAT_RAW"))


data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

# Fill in missing things
data$TERM[is.na(data$TERM)] <- "MISSING"
data$CMINDC[is.na(data$CMINDC) & data$FORM == "CM"] <- "MISSING"

### Edit start/end dates -----
data %<>% add_column(STDAT = data$STDAT_RAW, .after = "STDAT_RAW")
data %<>% add_column(ENDAT = data$ENDAT_RAW, .after = "ENDAT_RAW")
data %<>% add_column(ICDAT = data$ICDAT_RAW, .after = "ICDAT_RAW")

data <- UNKN_dates(start = "STDAT", end = "ENDAT")

data$ICDAT[is.na(data$ICDAT)] <- "MISSING"
data %<>% dplyr::select(-ICDAT_RAW)

### Calculate -----
# Sort by subject, start date
data <- data[order(data$Subject, as.Date(data$STDAT, format = "%d %b %Y")),]
rownames(data) <- 1:nrow(data)

for(sub in unique(data$Subject)){
  subj <- data[data$Subject == sub,]
  if("AE" %in% subj$FORM){
    ae_start <- as.Date(subj$STDAT[subj$FORM == "AE"], format = "%d %b %Y")
    ae_end <- as.Date(subj$ENDAT[subj$FORM == "AE"], format = "%d %b %Y")
    ae_span <- interval(start = ae_start, end = ae_end)
  }else{
    ae_start <- NA
    ae_end <- NA
    ae_span <- NA
  }
  if("MH" %in% subj$FORM | "OH" %in% subj$FORM){ 
    mhoh_start <- as.Date(subj$STDAT[subj$FORM == "MH" | subj$FORM == "OH"], format = "%d %b %Y")
    mhoh_end <- as.Date(subj$ENDAT[subj$FORM == "MH" | subj$FORM == "OH"], format = "%d %b %Y")
    mhoh_span <- interval(start = mhoh_start, end = mhoh_end)
  }else{
    mhoh_start <- NA
    mhoh_end <- NA
    mhoh_span <- NA
  }
  if(any(subj$FORM == "CM")){
    subj_cm <- subj[subj$FORM == "CM",]
    for(x in 1:nrow(subj_cm)){
      # Only look at subjects with informed consent dates and CM start dates, and not prophylaxis
      if((!all(subj$ICDAT == "MISSING")) & subj_cm$STDAT[x] != "MISSING" & 
         subj_cm$CMINDC[x] != "PROPHYLAXIS"){
        cm_start <- as.Date(subj_cm$STDAT[x], format = "%d %b %Y")
        cm_end <- as.Date(subj_cm$ENDAT[x], format = "%d %b %Y")
        cm_span <- interval(cm_start, cm_end)
        cm_ic <- as.Date(subj_cm$ICDAT[x], format = "%d %b %Y")
        # ConMed was started after Informed Consent
        if(cm_start > cm_ic){
          # Subject has no AEs.
          if(all(is.na(ae_span))){
            data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
              paste0("CM start > IC, No AE listed for subject.\n")
            data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
              paste0("The CM start date is after Informed Consent, but there are no AEs listed for the subject. Please reconcile.\n")
          }
          # Subject has AEs.
          if(!all(is.na(ae_span))){
            # There are CM/AE Overlaps
            if(any(lubridate::int_overlaps(cm_span, ae_span))){
              data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                paste0("AE Overlaps.\n")
              # Which ones overlap?
              AE_OLAPS <-  data.frame(AE_TERM = subj$TERM[subj$FORM == "AE"][lubridate::int_overlaps(cm_span, ae_span)],
                                      AE_SPAN = ae_span[lubridate::int_overlaps(cm_span, ae_span)])
              if(!is.na(subj_cm$link[x])){
                # Terms match exactly.
                if(subj_cm$link[x] %in% AE_OLAPS$AE_TERM){
                  link_span <- AE_OLAPS$AE_SPAN[subj_cm$link[x] == AE_OLAPS$AE_TERM]
                  if(length(link_span) > 1){
                    data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                      paste0("More than one matching AE term/date overlap.\n")
                    data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
                      paste0("There is more than one overlapping AE Term for this CM. Please reconcile.\n")
                  }
                  if(length(link_span) == 1){
                    # CM starts before AE.
                    if(cm_start < int_start(link_span)){
                      data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                        paste0("Explained by Link, CM starts before AE.\n")
                      
                    }else{
                      # CM started after AE.
                      data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                        paste0("CM started after linked AE.\n")
                      data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
                        paste0("CM started after the corresponding AE. Please reconcile.\n")
                    }
                  }
                }else{
                  for(y in 1:nrow(AE_OLAPS)){
                    data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                      paste0(AE_OLAPS$AE_TERM[y], "/", AE_OLAPS$AE_SPAN[y], "\n")
                  }
                }
              }else{
                for(y in 1:nrow(AE_OLAPS)){
                  data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                    paste0(AE_OLAPS$AE_TERM[y], "/", AE_OLAPS$AE_SPAN[y], "\n")
                }
              }
            }else{
              # There are no CM/AE Overlaps
              data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                paste0("CM start > IC, No AE overlap.\n")
              data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
                paste0("The CM start date is after Informed Consent, but there is no overlapping AE. Please reconcile.\n")
            }
          }
        }
        # ConMed was started before Informed Consent
        if(cm_start < cm_ic){
          # Subject has no MHOH.
          if(all(is.na(mhoh_span))){
            data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
              paste0("CM start < IC, No MH/OH listed for subject.\n")
            data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
              paste0("The CM start date is before Informed Consent, but there are no MH/OH listed for the subject. Please reconcile.\n")
          }
          # Subject has MHOH.
          if(!all(is.na(mhoh_span))){
            # There are CM/MH/OH Overlaps
            if(any(lubridate::int_overlaps(cm_span, mhoh_span))){
              data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                paste0("MH/OH Overlaps.\n")
              # Which ones overlap?
              MHOH_OLAPS <-  data.frame(MHOH_TERM = subj$TERM[subj$FORM == "MH" | subj$FORM == "OH"][lubridate::int_overlaps(cm_span, mhoh_span)],
                                        MHOH_SPAN = mhoh_span[lubridate::int_overlaps(cm_span, mhoh_span)])
              if(!is.na(subj_cm$link[x])){
                # Terms match exactly.
                if(subj_cm$link[x] %in% MHOH_OLAPS$MHOH_TERM){
                  link_span <- MHOH_OLAPS$MHOH_SPAN[subj_cm$link[x] == MHOH_OLAPS$MHOH_TERM]
                  if(length(link_span)>1){
                    data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                      paste0("More than one matching MHOH term/date overlap.\n")
                    data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
                      paste0("There is more than one overlapping MHOH Term for this CM. Please reconcile.\n")
                  }
                  if(length(link_span) == 1){
                    # CM starts before MHOH.
                    if(cm_start < int_start(link_span)){
                      data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                        paste0("Explained by Link, CM starts before MHOH.\n")
                    }else{
                      # CM started after MHOH.
                      data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                        paste0("CM started after linked MHOH.\n")
                      data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
                        paste0("CM started after the corresponding MHOH. Please reconcile.\n")
                    }
                  }
                }else{
                  for(y in 1:nrow(MHOH_OLAPS)){
                    data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                      paste0(MHOH_OLAPS$MHOH_TERM[y], "/", MHOH_OLAPS$MHOH_SPAN[y], "\n")
                  }
                }
              }else{
                for(y in 1:nrow(MHOH_OLAPS)){
                  data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                    paste0(MHOH_OLAPS$MHOH_TERM[y], "/", MHOH_OLAPS$MHOH_SPAN[y], "\n")
                }
              }
            }else{
              # There are no CM/MH/OH Overlaps
              data$Comments[data$Subject == sub & data$FORM == "CM"][x] %<>%
                paste0("CM start < IC, No MH/OH overlap.\n")
              data$Query[data$Subject == sub & data$FORM == "CM"][x] %<>%
                paste0("The CM start date is before Informed Consent, but there is no overlapping MH/OH. Please reconcile.\n")
            }
          }
        }
      }
    }
  }
}

### Print -----
message("AE/CM/MH/OH ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/AE_CM_MH_OH.xlsx", 
                 sheetName = "AE_CM_MH_OH", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()

# Adverse Events -----
message("Adverse Events ... Calculating")

data <- ae %>% dplyr::select(project, Subject, AETERM, AELOC_STD, AESTDAT_RAW,
                             AEENDAT_RAW, AEONGO) %>% unique()
colnames(data) %<>% dplyr::recode("AETERM" = "TERM",
                                "AELOC_STD" = "LOC")
# Pull in Date of Discontinuation
data2 <- dplyr::select(ds_eos, Subject, DSSTDAT_RAW) %>% unique()
data <- merge(x = data, y = data2, by = "Subject", all.x = TRUE)

colnames(data) %<>% dplyr::recode("DSSTDAT_RAW" = "DSSTDAT")
data[data == ""] <- NA
data$Comments <- NA
data$Query <- NA
data$Reviewer <- paste("CLK", format(as.Date(Sys.Date()),"%d%b%Y") %>% toupper())

# Fill in missing things
data$TERM[is.na(data$TERM)] <- "MISSING"
data$LOC[is.na(data$LOC)] <- "MISSING"
data$AEONGO[is.na(data$AEONGO)] <- "MISSING"

### Edit start/end dates -----
data %<>% add_column(STDAT = data$AESTDAT_RAW, .after = "AESTDAT_RAW")
data %<>% add_column(ENDAT = data$AEENDAT_RAW, .after = "AEENDAT_RAW")

data <- UNKN_dates(start = "STDAT", end = "ENDAT")

### Location does not match Term -----
test <- data.frame(TERM = data$TERM, LOC = data$LOC) %>% unique()
test <- test[!(test$TERM == "MISSING" | test$LOC == "MISSING"),]
knitr::kable(test[,c("TERM", "LOC")])

### Calculate -----
# Sort
data <- data[order(data$Subject, as.Date(data$STDAT, format = "%d %b %Y")),]
rownames(data) <- 1:nrow(data)

for(x in 1:nrow(data)){
  ### Query if Resolution Date is missing/after the Date of Discontinuation -----
  if(!is.na(data$DSSTDAT[x]) & data$AEONGO[x] != "Yes"){
    dis_date <- as.Date(data$DSSTDAT[x], "%d %b %Y")
    end_date <- as.Date(data$ENDAT[x], "%d %b %Y")
    if(end_date > dis_date){
      data$Comments[x] %<>% paste0("End date > Discon Date, ongoing not Yes.\n")
      data$Query[x] %<>% paste0("The end date for this subject is missing or occurs after the Discontinuation Date, and Ongoing is not Yes. Please reconcile.\n")
    }
  }
}

for(sub in unique(data$Subject)){
  subj <- data[data$Subject == sub,]
  ### Query if there are overlapping AEs with the same TERM -----
  for(x in 1:nrow(subj)){
    others <- subj[-x,]
    if(subj$TERM[x] %in% others$TERM){
      data$Comments[data$Subject == sub][x] %<>%
        paste0("Duplicate AE TERM.\n")
      # Do they overlap?
      ae_start <- as.Date(data$STDAT[x], "%d %b %Y")
      ae_end <- as.Date(data$ENDAT[x], "%d %b %Y")
      ae_span <- lubridate::interval(ae_start, ae_end)

      other_ae_start <- as.Date(others$STDAT, "%d %b %Y")
      other_ae_end <- as.Date(others$ENDAT, "%d %b %Y")
      other_ae_span <- interval(start = other_ae_start, end = other_ae_end)
      
      ae_span_overlaps <- lubridate::int_overlaps(ae_span, other_ae_span)
      ae_term_overlaps <- subj$TERM[x] == others$TERM
      ae_match <- ae_span_overlaps & ae_term_overlaps
      
      data$Comments[data$Subject == sub][x] %<>%
        paste0("AE Term/Date Overlaps:\n", others$TERM[ae_match], "/", 
               others$STDAT[ae_match], "---", others$ENDAT[ae_match], "\n")
    }
  }
}

### Print -----
message("Adverse Events ... Printing Data")
data$Query %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Query %<>% gsub(pattern = "\n$", replacement = "") 
data$Query[is.na(data$Query)] <- "NA"
data$Comments %<>% gsub(pattern = "^NA |^NA", replacement = "") 
data$Comments %<>% gsub(pattern = "\n$", replacement = "") 
data$Comments[is.na(data$Comments)] <- "NA"

xlsx::write.xlsx(x = data, file = "U:/Study/Data Listing/Adverse_Events.xlsx", 
                 sheetName = "AE", showNA = FALSE, row.names = FALSE,
                 col.names = TRUE, append = FALSE)

beepr::beep()
