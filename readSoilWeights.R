################################################################################
#
# Read multiple soil weight files, and join them together in a data frame
#
################################################################################

readSoilWeights <- function() {
  
  # Prompt user to select a soil data file
  promptInput <-
    readline('\n\nPress <ENTER> to select a soil data file... ')
  # Check input character
  if(promptInput != '') errorHandler('Invalid input')
  # Spawn file chooser to have user select the appropriate data file
  path <- 'W:/R/SPNR/spnr tables/'
  fileExt <- '.xlsx'
  defaultDataFile <- paste(path, 'ARDEC_Soil_N_2015', fileExt, sep = '')
  dataFileName <- choose.files(default = defaultDataFile)
  if(length(dataFileName) == 0) errorHandler('No data file selected')
  cat('\n\n...reading data file...')
  
  # Convert soil data file to data frame
  #
  # Package xlsx provides read/write functions for xls and xlsx files
  library(xlsx)
  #
  # Sheet 2 lists the classes for each column on sheet 1
  colClassSheet <- read.xlsx2(dataFileName, sheetIndex = 2,
                              stringsAsFactors = FALSE)
  colClassVector <- colClassSheet[, 2]
  names(colClassVector) <- colClassSheet[, 1]
  # Now read worksheet 1, using colClassVector
  dataSheet <- read.xlsx2(dataFileName, sheetIndex = 1,
                          colClasses = colClassVector, stringsAsFactors = FALSE)
  
  # Check for duplicate lab numbers in dataSheet by first creating an NA-free
  # copy of dataSheet$labNum
  dupTestVector <- na.omit(dataSheet$labNum)
  # The object 'dups' is a logical vector
  dups <- duplicated(dupTestVector)
  # Create a vector of indices
  dupIndices <- which(dups)
  # If duplicates are present then throw a fatal error
  if(length(dupIndices) > 1) {
    errorHandler(cat('Duplicate lab number(s) present in soil data file:',
                     dataSheet$labNum[dupIndices]))
    
    # Prompt user to select a folder that contains soil weight files
    promptInput <-
      readline(
        '\n\nPress <ENTER> to select a folder containing soil weight files... ')
    # Choose folder containing soil weights
    folder <- choose.dir(default =
                           'V:/Halvorson/ADHLab/Soil Weights/2015 Soil Weights/')
    if(is.na(folder)) errorHandler('No folder selected')
    
    # Print status message
    cat('\n\n...reading soil weights...')  
    
    # Capture filenames with xls and xlsx extensions into list soilWeightFiles
    soilWeightFiles <- Sys.glob(paste(folder, '/*.xls*', sep = ''))
    
    # The xlsx package contains read.xlsx2, which reads both xls and xlsx files
    library(xlsx)
    # Concatenate the names of all files in the default folder
    for(i in 1:length(soilWeightFiles)){
      soilWeightList <- lapply(soilWeightFiles,
                               function(x) read.xlsx2(x,
                                                      sheetIndex = 1, stringsAsFactors = FALSE, startRow = 4,
                                                      colIndex = 1:2, header=FALSE,
                                                      colClasses = c('numeric', 'numeric')))
    }
    
    # Transform the list of two-column data frames into a single data frame having
    # two columns
    soilWeightDF <- do.call('rbind', soilWeightList)
    # Retain only rows with no NA values
    soilWeightDF <- na.omit(soilWeightDF)
    # Provide descriptive column names
    names(soilWeightDF) <- c('labNum', 'labWeight')
    # Check for duplicate lab numbers in soilWeightDF
    dups <- NULL  # This object name was used above
    dups <- duplicated(soilWeightDF$labNum)
    dupIndices <- which(dups)
    # If duplicates are present then throw a fatal error
    if(length(dupIndices) > 1) {
      errorHandler(cat('Duplicate lab number(s) present in soil weight files:',
                       soilWeightDF$labNum[dupIndices]))
    }
    
    # Find number of rows in soilWeightDF
    weightRows <- nrow(soilWeightDF)
    # Initialize object to record matchless lab numbers
    missingNums <- integer()
    # For each row in soilWeightDF...
    for(weightRow in 1:weightRows) {
      # ... find row in dataSheet matching current labNum
      dataRow <- which(dataSheet$labNum == soilWeightDF$labNum[weightRow])
      # If no such row exists then record current labNum
      if(length(dataRow) == 0) {
        missingNums <- c(missingNums, soilWeightDF$labNum[weightRow])
        # Otherwise copy labWeight to dataSheet for current labNum  
      } else {
        dataSheet$labWeight[dataRow] <- soilWeightDF$labWeight[weightRow]
      }
    }
    
    # Output status message
    cat('\n\n...saving updated soil file...')
    
    # Create new soil file name
    #
    # Position of character preceding '.'
    nameEndPos <- nchar(dataFileName) - nchar(fileExt)
    # Data file name with dot and extension removed
    nameEnd <- substr(dataFileName, start = 1, stop = nameEndPos)
    # Insert ' - Updated' at end of file name, and reattach fileExt
    filename <- paste(nameEnd, ' - Updated', fileExt, sep = '')
    
    # Write new soil file
    write.xlsx2(dataSheet, file = filename, showNA = FALSE, row.names = FALSE)
    
    # Output status message
    cat('\n\n...done.')
    
    # Display list of lab numbers that didn't appear in dataSheet
    if(length(missingNums) != 0) {
      missingNums <- sort(missingNums)
      cat('\n\nLab number(s) not found in data sheet:\n\n', missingNums)
    }
  }
}


################################################################################
#
# This function handles fatal errors with message output
#
################################################################################

errorHandler <- function (errorMessage) {
  cat('\n\n')
  fullMessage <- paste(errorMessage, '. Program halted.', sep = '')
  stop(fullMessage, call. = FALSE)
}
