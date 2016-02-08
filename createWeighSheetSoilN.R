#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#
# Assign soil lab numbers to a study, and create a soil weigh sheet for soil N
#
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# This function creates a printable weigh sheet
createWeighSheet <- function() {
  
  # Prompt user for weigh sheet details
  cat('Choose a study:\n')
  cat('\n\t1. Rotation 1\n\t2. Rotation 2\n\t3. Rotation 3\n\t4. Rotation 4')
  cat('\n\t5. Rotation 5\n\t6. Strip Till\n\t7. DMP\n\t8. 200/300B\n\t9. 300A')
  stNum <- readline('\n\     ? ')
  st <- NA_character_
  if(stNum == '1') st <- 'Rotation 1'
  if(stNum == '2') st <- 'Rotation 2'
  if(stNum == '3') st <- 'Rotation 3'
  if(stNum == '4') st <- 'Rotation 4'
  if(stNum == '5') st <- 'Rotation 5'
  if(stNum == '6') st <- 'Strip Till'
  if(stNum == '7') st <- 'DMP'
  if(stNum == '8') st <- '200/300B'
  if(stNum == '9') st <- '300A'
  if(is.na(st)) fatalError('Invalid selection')
  
  # Initialize plot suffix as NA
  pltsfx <- NA_character_
  
  # Prompt user for E or W suffix if applicable
  if(stNum == '1' | stNum == '2' | stNum == '5' | stNum == '300A') {
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
  
  # Prompt user for maximum soil depth. (Use cat for text color consistency.)
  cat ('\nEnter maximum depth sampled, in feet
       (whole number only, or "0" for top foot composite):  ')
  maxDepth <- readline('\n\     ? ')
  # Convert to numeric
  maxDepth <- as.numeric(maxDepth)
  
  # If selection is invalid then throw an error
  if(maxDepth != 1 & maxDepth != 2 & maxDepth != 3 & maxDepth != 4 &
       maxDepth != 5 & maxDepth != 6 & maxDepth != 0) fatalError (
                                                        'Invalid selection')
    
  # Will add 'autonumber all' feature in the future.
  #  auto <- as.character(readline('Autonumber all remaining studies? (Y/N)  '))
  
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
  } else if(month >= 9 & month <= 12) {  #Fall sampling can occur in December
    season <- 'Fall'
#   } else {
#     season <- 'Winter'
  }
  
  # Prompt user for starting lab number
  cat('\n\nStarting lab number? (<ENTER> for autonumbering)  ')
  startNum <- as.integer(readline('\n\     ? '))
  cat('\n\nPress <ENTER> to select a soil template file... ')
  templatePrompt <- readline()
  if(templatePrompt != '') fatalError('Invalid selection')
  
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
  
  # If max depth is less than 6 feet then subset appropriate rows
  if(maxDepth < 6 & maxDepth > 0) weighSheet <-
    weighSheet[weighSheet$depthBottom <= maxDepth * 12, ]
  # A zero value of maxDepth indicates a composite sample of the top 12 inches
  if(maxDepth == 0) {
    # Subset 0 to 3 inches
    weighSheet <- weighSheet[weighSheet$depthBottom == 3, ]
    # Establish a 0-12 inch increment
    weighSheet$depthBottom <- 12
  }
  # Check for duplicate sampling events
#   dupCheckSub <- subset(dataSheet, study == st)
  if(!is.na(pltsfx) & st != 'DMP') dupCheckSub <- subset(dupCheckSub,
                                                         plotSuffix = pltsfx)
  if(dupCheckSub[1, 'sampDay'] == day & dupCheckSub[1, 'sampMonth'] == month &
       dupCheckSub[1, 'sampYear'] == year) {
    fatalError(paste('Duplicate sampling event exists on', textDate,
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
    fatalError(paste('The following lab numbers already exist in the', year,
                 'data file:\n', dups, '\n** Program halted. **'))
  }
  
  # If no rows match input criteria then halt script with an error message
  if(nrow(weighSheet) == 0) fatalError(
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
  names(weighSheet)[names(weighSheet) == 'depthTop_in'] <- 'From'
  names(weighSheet)[names(weighSheet) == 'depthBottom_in'] <- 'To'
  
  # If no plot suffix is present in weighSheet then remove that column.  Also,
  # arrange columns specifically in this order:
  if(weighSheet$Suffix[1] == '' | is.na(weighSheet$Suffix[1])) {
    keepCols <- c('Trt', 'Rep', 'Plot', 'From', 'To', 'LabNum')
  } else {keepCols <- c('Trt', 'Rep', 'Plot', 'Suffix', 'From', 'To', 'LabNum')
  }
  # Specify columns to keep in weighSheet
  weighSheet <- weighSheet[, keepCols]
  
  # Output status message
  cat('\n\n...saving weigh sheet...')
  
  # Create weigh file name
  if(is.na(pltsfx)) pltsfx <- ''
  weighFileName <- paste(path, season, ' ', year, ' ARDEC ', st, ' ', pltsfx,
                         ' soil NO3 NH4 weigh sheet.xlsx', sep = '')
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
  cat('\n\n...updating soil file...')
  
  # Create new soil file name
  filename <- paste(path, dataFileName, fileExt, sep = '')
  # Write new soil file
  write.xlsx2(dataSheetNew, file = filename, sheetName = 'Data',
              showNA = FALSE, row.names = FALSE)
  write.xlsx2(templateSheet2, file = filename, sheetName = 'Columns',
              showNA = FALSE, row.names = FALSE, append = TRUE)
  
  # Output status message
  cat('\n\n...done.')
}

# This function handles fatal errors with message output
fatalError <- function (errorMessage) {
  cat('\n\n', errorMessage, '\n')
  stop('Program halted', call. = FALSE)
}
