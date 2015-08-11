################################################################################
#
# Assign soil lab numbers to a study, and create a soil weigh sheet
#
################################################################################

# This function creates a printable weigh sheet
createWeighSheet <- function() {
  
  # Prompt user for weigh sheet details
  cat('Choose a study:\n')
  cat('\n\t1. Rotation 1\n\t2. Rotation 2\n\t3. Rotation 3\n\t4. Rotation 4')
  cat('\n\t5. Rotation 5\n\t6. Strip Till\n\t7. DMP')
  stNum <- readline('\n\n\     ? ')
  st <- NA_character_
  if(stNum == '1') st <- 'Rotation 1'
  if(stNum == '2') st <- 'Rotation 2'
  if(stNum == '3') st <- 'Rotation 3'
  if(stNum == '4') st <- 'Rotation 4'
  if(stNum == '5') st <- 'Rotation 5'
  if(stNum == '6') st <- 'Strip Till'
  if(stNum == '7') st <- 'DMP'
  if(is.na(st)) errorHandler('Invalid selection')
  
  # Initialize plot suffix as NA
  pltsfx <- NA_character_
  
  # Prompt user for E or W suffix if applicable
  if(stNum == '1' | stNum == '2' | stNum == '5') {
    cat('Choose a suffix:\n')
    cat('\n\t1. East\n\t2. West')
    pltsfxNum <- readline('\n\n     ? ')
    if(pltsfxNum == '1') pltsfx <- 'E'
    if(pltsfxNum == '2') pltsfx <- 'W'
    if(is.na(pltsfx)) errorHandler('Invalid selection')
    
    # Prompt user for CC or SS-C if applicable  
  }else if (stNum == '6') {
    cat('Choose a suffix:\n')
    cat('\n\t1. CC\n\t2. SS-C')
    pltsfxNum <- readline('\n\n     ? ')
    if(pltsfxNum == '1') pltsfx <- 'CC'
    if(pltsfxNum == '2') pltsfx <- 'SS-C'
    if(is.na(pltsfx)) errorHandler('Invalid selection')
  }
  
  # Prompt user for additional input.
  # Will add 'autonumber all' feature in the future.
  #  auto <- as.character(readline('Autonumber all remaining studies? (Y/N)  '))
  
  # Prompt user for sampling date
  textDate <- readline('Enter sampling date (mm-dd-yyyy)  ')
  # Test for valid date
  dateTest <- as.Date(textDate, '%m-%d-%Y')
  if(is.na(dateTest)) errorHandler('Invalid date')
  
  # Split textDate on non-digit characters, giving three strings
  dateList <- strsplit(textDate, '[^0-9]')
  # Extract date components as numeric values
  month <- as.integer(dateList[[1]][1])
  day <- as.integer(dateList[[1]][2])
  year <- as.integer(dateList[[1]][3])
  
  # Prompt user for starting lab number
  startNum <-
    as.integer(readline('Starting lab number? (<ENTER> for autonumbering)  '))
  templatePrompt <-
    readline('\n\nPress <ENTER> to select a weigh sheet template file... ')
  
  # Spawn file chooser to have user select a template file from which to build
  # the weigh sheet
  path <- 'W:/R/SPNR/spnr tables/'
  fileExt <- '.xlsx'
  defaultTemplateFile <- paste(path, 'ARDEC_Soil_N_Template', fileExt, sep = '')
  filename <- choose.files(default = defaultTemplateFile)
  
  # Output status message
  cat('\n\nReading template file...')
  
  # Package xlsx provides read/write functions for Excel files
  library(xlsx)
  # Worksheet 2 contains the classes for each column used on worksheet 1
  templateSheet2 <- read.xlsx2(filename, sheetIndex = 2,
                               stringsAsFactors = FALSE)
  colClassVector <- templateSheet2[, 2]
  names(colClassVector) <- templateSheet2[, 1]
  # Now read worksheet 1, using colClassVector
  templateSheet <- read.xlsx2(filename, sheetIndex = 1,
                              colClasses = colClassVector,
                              stringsAsFactors = FALSE)
  
  # Open existing data file
  dataFileName <- paste('ARDEC_Soil_N_', as.character(year), sep = '')
  defaultDataFile <- paste(path, dataFileName, fileExt, sep = '')
  dataSheet2 <- read.xlsx2(defaultDataFile, sheetIndex = 2,
                           stringsAsFactors = FALSE)
  colClassVector <- dataSheet2[, 2]
  names(colClassVector) <- dataSheet2[, 1]
  # Now read worksheet 1, using colClassVector
  dataSheet <- read.xlsx2(defaultDataFile, sheetIndex = 1,
                          colClasses = colClassVector, stringsAsFactors = FALSE)
  
  # Output status message
  cat('\n\n...creating weigh sheet...')
  
  # If DMP was specified for the weighSheet...
  if(st == 'DMP') {
    # ... then subset by study
    weighSheet <- subset(templateSheet, study == st)
    # copy non-subsetted rows into excelSheetLite
    #    excelSheetLite <- subset(templateSheet, !(study == st))
    # and sort weighSheet by plot number, suffix and depth (unique to DMP)
    weighSheet <- weighSheet[order(weighSheet$plotNumber, weighSheet$plotSuffix,
                                   weighSheet$depthTop), ]
    
    # Otherwise subset any non-DMP weighSheet based on whether pltsfx is defined
  } else {
    if(is.na(pltsfx)) {
      weighSheet <- subset(templateSheet, study == st)
    } else {
      weighSheet <- subset(templateSheet, study == st & plotSuffix == pltsfx)
    }
    # Sort non-DMP weigh sheet
    weighSheet <- weighSheet[order(weighSheet$plotNumber,
                                   weighSheet$depthTop), ]
  }
  
  # Check for duplicate sampling events
  dupCheckSub <- subset(dataSheet, study == st)
  if(!is.na(pltsfx) & st != 'DMP') dupCheckSub <- subset(dupCheckSub,
                                                         plotSuffix = pltsfx)
  if(dupCheckSub[1, 'sampDay'] == day & dupCheckSub[1, 'sampMonth'] == month &
       dupCheckSub[1, 'sampYear'] == year) {
    errorHandler(paste('Duplicate sampling event exists on', textDate,
                       'for specified study.'))
  }
  
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
    errorHandler(paste('The following lab numbers already exist in the', year,
                 'data file:\n', dups, '\n** Program halted. **'))
  }
  
  # If no rows match input criteria then halt script with an error message
  if(nrow(weighSheet) == 0) errorHandler(
    'No match in soil file.  Check input values.')
  
  # Output status message
  cat('\n\n...assigning lab numbers...')
  
  # Assign lab numbers and sampling date to sorted weigh sheet
  weighSheet$labNum <- numVec
  weighSheet$sampDay <- day
  weighSheet$sampMonth <- month
  weighSheet$sampYear <- year
  
  # Package dplyr provides the bind_rows function
  library(dplyr)
  # Append weighSheet to dataSheet, inserting NAs where columns don't match
  dataSheetNew <- as.data.frame(bind_rows(dataSheet, weighSheet))
  
  # Rename weighSheet column headers for clarity
  names(weighSheet)[names(weighSheet) == 'labNum'] <- 'LabNum'
  names(weighSheet)[names(weighSheet) == 'trtLevel'] <- 'Trt'
  names(weighSheet)[names(weighSheet) == 'rep'] <- 'Rep'
  names(weighSheet)[names(weighSheet) == 'plotNumber'] <- 'Plot'
  names(weighSheet)[names(weighSheet) == 'plotSuffix'] <- 'Suffix'
  names(weighSheet)[names(weighSheet) == 'depthTop'] <- 'From'
  names(weighSheet)[names(weighSheet) == 'depthBottom'] <- 'To'
  
  # If no plot suffix is present in weighSheet then remove that column.  Also,
  # arrange columns specifically in this order:
  if(weighSheet$Suffix[1] == '' | is.na(weighSheet$Suffix[1])) {
    keepCols <- c('Trt', 'Rep', 'Plot', 'From', 'To', 'LabNum')
  } else {keepCols <- c('Trt', 'Rep', 'Plot', 'Suffix', 'From', 'To', 'LabNum')
  }
  # Specify columns to keep in weighSheet
  weighSheet <- weighSheet[keepCols]
  
  # Output status message
  cat('\n\n...saving weigh sheet...')
  
  # Create weigh file name
  if(is.na(pltsfx)) pltsfx <- ''
  weighFileName <- paste(path, st, pltsfx, ' soil weigh sheet.xlsx', sep = '')
  # Create a workbook for the weigh sheet
  weighSheetWB <- createWorkbook()
  # Create a worksheet
  weighSheetWS <- createSheet(weighSheetWB, sheetName = 'Weigh_Sheet')
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
    rows <- createRow(sheet, rowIndex = rowIndex)
    sheetTitle <- createCell(rows, colIndex = 1)
    setCellValue(sheetTitle[[1,1]], title)
    setCellStyle(sheetTitle[[1,1]], titleStyle)
  }
  
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
  
  # Create main title
  titleText <- paste(year, ' ', season, ',', sep = '' )
  titleText <- paste(titleText, 'ARDEC', st, pltsfx)
  dateSubtitleText <- paste('Sampling date:', textDate)
  addTitle(weighSheetWS, rowIndex = 1, title = titleText,
           titleStyle = titleStyle)
  # Create subtitle for analysis type
  addTitle(weighSheetWS, rowIndex = 2, title = 'Soil Nitrate and Ammonium',
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
  cat('\n\n...saving updated soil file...')
  
  # Create new soil file name
  filename <- paste(path, dataFileName, " - Updated", fileExt, sep = '')
  # Write new soil file
  write.xlsx2(dataSheetNew, file = filename, showNA = FALSE, row.names = FALSE)
  
  # Output status message
  cat('\n\n...done.')
}

# This function responds to invalid inputs
errorHandler <- function (errorMessage) {
  cat('\n\n', errorMessage, '\n')
  stop('Program halted', call. = FALSE)
}
