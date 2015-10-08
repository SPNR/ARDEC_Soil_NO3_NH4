################################################################################
#
# Assign lab numbers to plant samples from a specific study, and create a 
# weigh sheet for CN analysis
#
################################################################################

# This function creates a printable weigh sheet for plant CN analysis
labSheet <- function() {
  
  # Silence warnings throughout this function
  oldWarn <- getOption("warn")
  options(warn = -1)
  
  # Prompt user for weigh sheet details
  cat('Choose a study:\n')
  cat('\n\t1. Rotation 1\n\t2. Rotation 2\n\t3. Rotation 3\n\t4. Rotation 4')
  cat('\n\t5. Rotation 5\n\t6. Strip Till\n\t7. DMP')
  stNum <- readline('\n\     ? ')
  st <- NA_character_
  if(stNum == '1') st <- 'Rotation 1'
  if(stNum == '2') st <- 'Rotation 2'
  if(stNum == '3') st <- 'Rotation 3'
  if(stNum == '4') st <- 'Rotation 4'
  if(stNum == '5') st <- 'Rotation 5'
  if(stNum == '6') st <- 'Strip Till'
  if(stNum == '7') st <- 'DMP'
  if(is.na(st)) fatalError('Invalid selection')
  
  # Initialize plot suffix as NA
  pltsfx <- NA_character_
  
  # Prompt user for E or W suffix if applicable
  if(stNum == '1' | stNum == '2' | stNum == '5') {
    cat('\nChoose a suffix:\n')
    cat('\n\t1. East\n\t2. West')
    pltsfxNum <- readline('\n     ? ')
    if(pltsfxNum == '1') pltsfx <- 'E'
    if(pltsfxNum == '2') pltsfx <- 'W'
    if(is.na(pltsfx)) fatalError('Invalid selection')
    
  # Prompt user for CC or SS-C if applicable  
  }else if (stNum == '6') {
    cat('\nChoose a suffix:\n')
    cat('\n\t1. CC\n\t2. SS-C')
    pltsfxNum <- readline('\n     ? ')
    if(pltsfxNum == '1') pltsfx <- 'CC'
    if(pltsfxNum == '2') pltsfx <- 'SS-C'
    if(is.na(pltsfx)) fatalError('Invalid selection')
  }

  # Prompt user for crop type
  cat('\nSelect crop type:\n')
  cat('\n\t1. Corn\n\t2. Pinto Bean\n\t3. Soybean\n\t4. Sorghum Sudangrass')
  cat('\n\t5. Pea')
  crop <- NA_character_
  crop <- readline('\n\     ? ')
  segments <- NA_character_
  if(crop == '1') {
    crop <- 'Corn'
    segments <- c('Stalk', 'Cob', 'Grain')
  }
  if(crop == '2') {
    crop <- 'Pinto Bean'
    segments <- c('Stem', 'Grain')
  }
  if(crop == '3') {
    crop <- 'Soybean'
    segments <- c('Stem', 'Grain')
  }
  if(crop == '4') {
    crop <- 'Sorghum Sudangrass'
    segments <- c('Stalk')
  }
  if(crop == '5') {
    crop <- 'Pea'
    segments <- c('Stem', 'Grain')
  if(is.na(crop)) fatalError('Invalid selection')
  }
  
  # Prompt user for sampling date
  cat('\nEnter sampling date (mm-dd-yyyy):')
  textDate <- readline('\n\     ? ')
  # Test for valid date
  dateTest <- as.Date(textDate, '%m-%d-%Y')
  if(is.na(dateTest)) fatalError('Invalid date')
  
  # Split textDate on non-digit characters, giving three strings
  dateList <- strsplit(textDate, '[^0-9]')
  # Extract date components as numeric values
  month <- as.integer(dateList[[1]][1])
  day <- as.integer(dateList[[1]][2])
  year <- as.integer(dateList[[1]][3])
  
  # Determine sampling season for weighSheet title
  if(month >= 3 & month <= 5) {
    season <- 'Spring'
  } else if(month >= 6 & month <= 8) {
    season <- 'Summer'
  } else if(month >= 9 & month <= 11) {
    season <- 'Fall'
  } else {
    season <- 'Winter'
  }
  
  # Prompt user for starting lab number
  cat('\n\nStarting lab number? (<ENTER> for autonumbering)  ')
  startNum <- as.integer(readline('\n\     ? '))
#   cat('\n\nPress <ENTER> to select a plant template file... ')
#   templatePrompt <- readline()
  
  # Is this input check necessary?
  # if(templatePrompt != '') fatalError('Invalid selection')
  
  # Spawn file chooser to have user select a template file from which to build
  # the weigh sheet
  #path <- 'C:/Users/Robert/Documents/GitHub/ARDEC_Soil_NO3_NH4/'
  path <- 'W:/R/SPNR/spnr tables/'
  fileExt <- '.xlsx'
  templateFile <- paste(path, 'ARDEC_Plant_CN_Template', fileExt,
                               sep = '')
#   filename <- choose.files(default = defaultTemplateFile)
  
  # Output status message
  cat('\n\nReading template file...')
  
  # Package xlsx provides read/write functions for Excel files
  
  library(xlsx)
  # Worksheet 2 contains the classes for each column used on worksheet 1
  templateSheet2 <- read.xlsx2(templateFile, sheetIndex = 2,
                               stringsAsFactors = FALSE)
  colClassVector <- templateSheet2[, 2]
  names(colClassVector) <- templateSheet2[, 1]
  # Now read worksheet 1, using colClassVector
  templateSheet <- read.xlsx2(templateFile, sheetIndex = 1,
                              colClasses = colClassVector,
                              stringsAsFactors = FALSE)
  
  # Open existing data file
  dataFileName <- paste('ARDEC_Plant_CN_', as.character(year), '_test',
                        sep = '')
  defaultDataFile <- paste(path, dataFileName, fileExt, sep = '')
  colClassSheet <- read.xlsx2(defaultDataFile, sheetIndex = 2,
                           stringsAsFactors = FALSE)
  colClassVector <- colClassSheet[, 2]
  names(colClassVector) <- colClassSheet[, 1]
  # Now read worksheet 1, using colClassVector
  dataSheet <- read.xlsx2(defaultDataFile, sheetIndex = 1,
                          colClasses = colClassVector, stringsAsFactors = FALSE)
  
  # Output status message
  cat('\n\n...creating weigh sheet...')
  
  # If DMP was specified for the weighSheet...
  if(st == 'DMP') {
    # ... then subset by study
    weighSheet <- subset(templateSheet, study == st)
  # Otherwise subset any non-DMP weighSheet based on whether pltsfx is defined
  } else {
    if(is.na(pltsfx)) {
      weighSheet <- subset(templateSheet, study == st)
    } else {
      weighSheet <- subset(templateSheet, study == st & plotSuffix == pltsfx)
    }
  }

#   # Add columns for plant segment and tray/well identifiers
#   #weighSheet$segment <- NA_character_
#   weighSheet$Tray <- NA_integer_
#   weighSheet$Well <- NA_character_
#-------- template file now includes these columns
  
  
  # Package dplyr provides arrange(), bind_rows(), etc.
  library(dplyr)
  
  # Duplicate all rows in weighSheet for each additional plant segment
  if(length(segments) == 2) weighSheet <- bind_rows(weighSheet, weighSheet)
  if(length(segments) == 3) weighSheet <-
    bind_rows(weighSheet, weighSheet, weighSheet)
  
  # Sort rows by plot number and suffix (if applicable)
  if(st == 'DMP') {
    weighSheet <- arrange(weighSheet, plotNumber, plotSuffix)
    # Else only one suffix appears on the weigh sheet
  } else weighSheet <- arrange(weighSheet, plotNumber)
  
  # Populate segment column
  weighSheet$plantSegment <- segments
  # Sort again, first by segment (because the segment column is populated by
  # cycling through the segment list)
  if(st == 'DMP') {
    weighSheet <- arrange(weighSheet, plantSegment, plotNumber, plotSuffix) 
  } else weighSheet <- arrange(weighSheet, plantSegment, plotNumber)

  # Create vector of tray well names
  wellNames <- c('A1', 'A2', 'A3', 'A4', 'A5', 'A6',
                 'B1', 'B2', 'B3', 'B4', 'B5', 'B6',
                 'C1', 'C2', 'C3', 'C4', 'C5', 'C6',
                 'D1', 'D2', 'D3', 'D4', 'D5', 'D6')
  
  # Number of trays needed, at 24 wells per tray
  # numTrays <- ceiling(nrow(weighSheet) / 24)
  
  # Add tray and well names to weighSheet
  rowsPerSeg <- nrow(weighSheet) / length(segments)
  
  for(i in 1:length(segments)) {
    # Populate tray number for each segment subset
    weighSheet$tray[((i - 1) * rowsPerSeg + 1):(i * rowsPerSeg)] <- i
    # Populate well names for each segment subset
    weighSheet$well[((i - 1) * rowsPerSeg + 1):(i * rowsPerSeg)] <- wellNames
  }

  

#   # Check for duplicate sampling events (existing dates are same as current)
#   dupCheckSub <- subset(dataSheet, study == st)
#   if(!is.na(pltsfx) & st != 'DMP') dupCheckSub <- subset(dupCheckSub,
#                                                          plotSuffix == pltsfx)
#   # If either sampDay or sampMonth exists then throw an error
#   if(!is.na(dupCheckSub[1, 'sampDay']) | !is.na(dupCheckSub[1, 'sampMonth'])) {
#     fatalError(paste('Duplicate sampling event exists on', textDate,
#                      'for specified study.'))
#   }
  
  # Check for duplicate sampling events by comparing existing dates to current
  # date.
  #
  # Extract a subset of current study and suffix if applicable
  if(st == 'DMP' | is.na(pltsfx)) {
    dupCheckSub <- filter(dataSheet, study == st)
  } else {
    dupCheckSub <- filter(dataSheet, study == st & plotSuffix == pltsfx)
  }
  dupDateSub <- filter(dupCheckSub, sampDay == day & sampMonth == month &
                           sampYear == year)
  if(nrow(dupDateSub > 0)) fatalError(
    paste('Duplicate sampling event exists on', textDate,
          'for specified study.'))
  
  # If startNum is not specified by the user then identify the largest existing
  # lab number in excelDataSheet, and start the new lab numbers after it.
  if(is.na(startNum)) {
    hiNum <- max(dataSheet$labNum, na.rm = TRUE)
    # If all labNum values are NA then the previous statement assigns
    # hiNum a value of -Inf.
    if(hiNum > 0) {
      startNum <- hiNum + 1
    } else {
      startNum <- 1
    }
  }
  
  # Generate list of desired lab numbers
  endNum <- startNum + nrow(weighSheet) - 1  # Last number to assign
  numVec <- c(startNum:endNum)
  
  # Check for duplicate lab numbers
  dupIndices <- which(numVec %in% dataSheet$labNum)
  # Halt script if duplicates are detected
  if(length(dupIndices) > 0) {
    dups <- numVec[dupIndices]
    fatalError(paste('The following lab numbers already exist in the', year,
                 'data file:\n', dups, '\n** Program halted. **'))
  }
  
  # If no rows match input criteria then halt script with an error message
  if(nrow(weighSheet) == 0) fatalError(
    'No match in plant file.  Check input values.')
  
  # Output status message
  cat('\n\n...assigning lab numbers...')
  
  # Assign lab numbers to sorted weigh sheet
  weighSheet$labNum <- numVec
  weighSheet$sampDay <- day
  weighSheet$sampMonth <- month
  weighSheet$sampYear <- year

  # Delete original rows from datasheet for current plot and suffix because
  # they will have no lab numbers
  if(is.na(pltsfx)) {
    dataSheet <- filter(dataSheet, !(study == st & is.na(labNum)))
  } else {
    dataSheet <- filter(dataSheet, !(study == st & plotSuffix == pltsfx &
                          is.na(labNum)))
  }
  
  # Remove tray and well columns before binding to dataSheet
  weighSheetLite <- weighSheet
  weighSheetLite$tray <- NULL
  weighSheetLite$well <- NULL
  # Populate crop column in weighSheetLite
  weighSheetLite$crop <- crop
  # Append weighSheet to dataSheet
  dataSheetNew <- as.data.frame(bind_rows(dataSheet, weighSheetLite))

  # Rename weighSheet column headers for clarity
  names(weighSheet)[names(weighSheet) == 'labNum'] <- 'LabNum'
  names(weighSheet)[names(weighSheet) == 'trtLevel'] <- 'Trt'
  names(weighSheet)[names(weighSheet) == 'rep'] <- 'Rep'
  names(weighSheet)[names(weighSheet) == 'plotNumber'] <- 'Plot'
  names(weighSheet)[names(weighSheet) == 'plotSuffix'] <- 'Suffix'
  names(weighSheet)[names(weighSheet) == 'plantSegment'] <- 'Segment'
  names(weighSheet)[names(weighSheet) == 'tray'] <- 'Tray'
  names(weighSheet)[names(weighSheet) == 'well'] <- 'Well'

  # If no plot suffix is present in weighSheet then remove that column.  Also,
  # arrange columns specifically in this order:
  if(weighSheet$Suffix[1] == '' | is.na(weighSheet$Suffix[1])) {
    keepCols <- c('Trt', 'Rep', 'Plot', 'Segment', 'Tray', 'Well', 'LabNum')
  } else {keepCols <- c('Trt', 'Rep', 'Plot', 'Suffix',  'Segment', 'Tray',
                        'Well', 'LabNum')
  }
  
  # Ensure that weighSheet is a data frame
  weighSheet <- as.data.frame(weighSheet)
  # Specify columns to keep in weighSheet
  weighSheet <- weighSheet[keepCols]
  
  # Output status message
  cat('\n\n...saving weigh sheet...')
  
  # Create weigh file name
  if(is.na(pltsfx)) pltsfx <- ''
  weighFileName <- paste(path, season, ' ', year, ' ARDEC ', st, ' ', pltsfx,
                         ' plant weigh sheet.xlsx', sep = '')
  # Create a workbook for the weigh sheet
  weighSheetWB <- createWorkbook()
  # Create a worksheet
  weighSheetWS <- createSheet(weighSheetWB, sheetName = 'Weigh Sheet')
  # Define cell styles for weigh sheet workbook
  titleStyle <- CellStyle(weighSheetWB) +
    Font(weighSheetWB, heightInPoints=14, isBold=TRUE)
  subtitleStyle1 <- CellStyle(weighSheetWB) +
    Font(weighSheetWB,  heightInPoints=12, isBold=TRUE)
  subtitleStyle2 <- CellStyle(weighSheetWB) +
    Font(weighSheetWB,  heightInPoints=12, isBold=FALSE)
  tableColNamesStyle <- CellStyle(weighSheetWB) +
    Font(weighSheetWB, isBold = TRUE) +
    Alignment(horizontal = 'ALIGN_CENTER')
  # This function formats the worksheet's title and subtitle
  addTitle <- function(sheet, rowIndex, title, titleStyle) {
    rows <- createRow(sheet, rowIndex = rowIndex)  # createRow is from xlsx
    sheetTitle <- createCell(rows, colIndex = 1)
    setCellValue(sheetTitle[[1,1]], title)
    setCellStyle(sheetTitle[[1,1]], titleStyle)
  }
  
  # Create main title
  titleText <- paste(year, ' ', season, ',', sep = '' )
  titleText <- paste(titleText, 'ARDEC', st, pltsfx, crop)
  dateSubtitleText <- paste('Sampling date:', textDate)
  addTitle(weighSheetWS, rowIndex = 1, title = titleText,
           titleStyle = titleStyle)
  # Create subtitle for analysis type
  addTitle(weighSheetWS, rowIndex = 2, title = 'Total Carbon and Nitrogen',
           titleStyle = subtitleStyle1)
  # Create subtitle for analysis type
  addTitle(weighSheetWS, rowIndex = 3, title = dateSubtitleText,
           titleStyle = subtitleStyle2)  
  # Add weigh sheet table to worksheet
  addDataFrame(weighSheet, weighSheetWS, startRow = 5, row.names = FALSE,
               colnamesStyle = tableColNamesStyle)
  # Save weigh sheet workbook
  saveWorkbook(weighSheetWB, weighFileName)
  
  # Output status message
  cat('\n\n...updating ARDEC plant data file...')
  
  # Create new plant file name
  filename <- paste(path, dataFileName, fileExt, sep = '')
  # Write new plant file
  write.xlsx2(dataSheetNew, file = filename, sheetName = 'Data',
              showNA = FALSE, row.names = FALSE)
  write.xlsx2(templateSheet2, file = filename, sheetName = 'Columns',
              showNA = FALSE, row.names = FALSE, append = TRUE)
  
  # Output status message
  cat('\n\n...done.')
  
  # Restore previous warning level
  options(warn = oldWarn)
}

# This function handles fatal errors with message output
fatalError <- function (errorMessage) {
  cat('\n\n', errorMessage, '\n')
  stop('Program halted', call. = FALSE)
}
