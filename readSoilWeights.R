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
  # Sheet 2 lists the classes for each column on sheet 1
  colClassSheet <- read.xlsx2(dataFileName, sheetIndex = 2,
                           stringsAsFactors = FALSE)
  colClassVector <- colClassSheet[, 2]
  names(colClassVector) <- colClassSheet[, 1]
  # Now read worksheet 1, using colClassVector
  dataSheet <- read.xlsx2(dataFileName, sheetIndex = 1,
                          colClasses = colClassVector, stringsAsFactors = FALSE)
  
  # Prompt user to select a folder that contains soil weight files
  promptInput <-
    readline(
      '\n\nPress <ENTER> to select a folder containing soil weight files... ')
  # Check input character
  if(promptInput != '') errorHandler('Invalid input')
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
    soilWeightList <- lapply(soilWeightFiles, function(x) read.xlsx2(x,
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
  
  # Print status message
  cat('\n\n...merging soil weights with data...')  

  # A better approach may be to rename the second column in soilWeightDF as
  # 'newWeight' for example.  Then merge both dfs on labNum, and then manipulate
  # the merged df such that newWeight values replace blank labWeight slots,
  # first checking for conflicts (i.e., two different weights for same labNum).
  mergedDF <- merge(dataSheet, soilWeightDF, by = 'labNum')

  
    #
  # Add weight values to soil file by matching lab numbers
  #mergedDF <- merge(dataSheet, soilWeightDF, by = c('labNum', 'labWeight'))
  
  # Output status message
  cat('\n\n...saving updated soil file...')
  
  # Rewrite the following lines to parse dataFileName and add ' - Updated'
  # before saving updated file.
  #
  # Create new soil file name
  filename <- paste(path, "Data File - Updated", fileExt, sep = '')
  
  # Write new soil file
  write.xlsx2(mergedDF, file = filename, showNA = FALSE, row.names = FALSE)
  
  # Output status message
  cat('\n\n...done.')
  
}


# This function handles fatal errors with message output
errorHandler <- function (errorMessage) {
  cat('\n\n')
  fullMessage <- paste(errorMessage, '. Program halted.', sep = '')
  stop(fullMessage, call. = FALSE)
}




################################################################################

## Example expression to access the 157th element of the X2 column of the 2nd
## data frame in soilWeightList:
# soilWeightList[[2]]$X2[157]