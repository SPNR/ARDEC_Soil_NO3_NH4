################################################################################
# 
# The function readLachatOutput reads the data table from Omnion's HTML output
# file into a data frame, and then reformats it to facilitate data analysis.
#
################################################################################


# Define sample year
sampleYear <- '2015'
# Default file path for Omnion HTML file
omnionFile <- paste('V:/Halvorson/ADHLab/Lachat/Lachat data/', sampleYear,
                    '/*.htm', sep = '')
# Prompt user to select an Omnion output file
promptInput <-
  readline('\n\nPress <ENTER> to select an Omnion .htm file... ')
# Check input character
if(promptInput != '') errorHandler('Invalid input')
# Launch file chooser
filename <- choose.files(default = omnionFile)

# Required for readHTMLTable function
library(XML)
# Import table from HTML file
lachatTable <- readHTMLTable(filename, which = 1, skip.rows = c(1:17),
                             colClasses = c('character', rep('numeric', 3)),
                             stringsAsFactors = FALSE)  
# Delete second column (Rep column)
lachatTable$V2 <- NULL
# Rename remaining columns
names(lachatTable) <- c('labNum', 'rawNO3NO2', 'rawNH4')
# Create a working copy of lachatTable
tableClean <- lachatTable
# Capitalize all letters to facilitate pattern matching
tableClean$labNum <- toupper(lachatTable$labNum)

# Add columns to tableClean
tableClean$dilution <- NA_integer_
tableClean$keep <- FALSE


# Initialize objects for upcoming for-loop
checkValues <- data.frame(labNum = character(0), rawNO3NO2 = numeric(0),
                          rawNH4 = numeric(0))

# Search for keywords in Omnion sample names, one row at a time
for(i in 1:nrow(tableClean)) {
  
  # If current name contains keyword 'check' then copy NO3NO2 and NH4 values
  # to checkValues
  if(grepl('CHECK', tableClean[i, 1])) {
    checkValues <- rbind(checkValues, c(tableClean$rawNO3NO2[i],
                                        tableClean$rawNH4[i]))
    
    # Else if current name contains '1:' then extract the dilution value and
    # lab number
  } else if(grepl('1:', tableClean$labNum[i])) {
    dilPos <- regexpr('1:', tableClean$labNum[i])
    # Read dilution value
    tableClean$dilution[i] <- suppressWarnings(
      as.numeric(substring(tableClean$labNum[i], dilPos + 2)))
    # Parse previous characters to extract corresponding lab number
    prevChars <- substring(tableClean$labNum[i], first = 1, last = dilPos - 1)
    # Acquire lab number as text. (Split function outputs a list.)
    labNumText <- split(prevChars, 'SAMPLE')
    # Transform labNumText into a numeric value
    tableClean$labNum[i] <- suppressWarnings(
      as.numeric(paste(unlist(labNumText), collapse = '')))
    # If current dilution or labNum is NA then print error message
    if(is.na(tableClean$dilution[i]) | is.na(tableClean$labNum[i])) {
      cat("** Invalid dilution value or lab number at row", i)
      cat('\n** This observation will be excluded from data summary!\n\n')
    } else {
      # Keep this row
      tableClean$keep[i] <- TRUE
    }
    
    # Else if current name contains keyword 'sample' then extract lab number
  } else if(grepl('SAMPLE', tableClean$labNum[i])) {
    labNumText <- substring(tableClean$labNum[i], first = 7)
    tableClean$labNum[i] <- suppressWarnings(as.numeric(labNumText))
    # if current labNum is NA then print error message
    if(is.na(tableClean$labNum[i])) {
      cat("** Invalid lab number at row", i)
      cat('\n** This observation will be excluded from data summary!\n\n')
    } else {
      # Otherwise keep this row
      tableClean$keep[i] <- TRUE
    }
    
    # Else (not a valid check, dilution or sample)      
  } else {
    cat('** Unreadable sample description at row', i)
    cat('\n** This observation will be excluded from data summary!\n\n')
  }
  
}  # End for-loop

# Check for negative measured values, and set those observations to
# keep = FALSE
negValRows <- which(tableClean$rawNO3NO2 < 0 | tableClean$rawNH4 < 0)
if(length(negValRows) > 0) {
  tableClean$keep[negValRows] <- FALSE
  cat('** Negative value(s) detected at these rows:', negValRows)
  cat('\n** These observations will be excluded from data summary!\n\n')
}

# Check for rows in which rawNO3NO2 rawNH4 values are out of range
hiValRows <- which(tableClean$rawNO3NO2 > 20 | tableClean$rawNH4 > 4)
if(length(hiValRows) > 0) {
  tableClean$keep[hiValRows] <- FALSE
  cat('** Out of range value(s) detected at these rows:', hiValRows)
  cat('\n** These observations will be excluded from data summary!\n\n')
}

# Check for duplicate lab numbers in otherwise valid observations
keepRows <- tableClean[tableClean$keep == TRUE, ]
firstDupRow <- anyDuplicated(keepRows$labNum)
if(firstDupRow > 0 ) {
  errorHandler(cat('** More than one instance of lab number',
                   keepRows$labNum[firstDupRow]))
}  

# Subset rows for final table
tableFinal <- tableClean[tableClean$keep == TRUE, ]

# Calculate dilution values
dilRows <- which(!is.na(tableFinal$dilution))
for(rowNum in dilRows) {
  tableFinal$rawNO3NO2[rowNum] <-
    tableFinal$rawNO3NO2[rowNum] * tableFinal$dilution[rowNum]
  tableFinal$rawNH4[rowNum] <-
    tableFinal$rawNH4[rowNum] * tableFinal$dilution[rowNum]
}

# Delete unneeded columns
tableFinal$dilution <- NULL
tableFinal$keep <- NULL

##############################################################################
#
# Insert code here to calculate adjusted nitrate & ammonium concentrations,
# as well as lb N per ac and kg N per ha for each depth segment.  These
# calculations require soil bulk density, soil lab weight and depth segment
# values, which are available in the soil file.
# 
##############################################################################

# Define default soil file path and name
path <- 'W:/R/SPNR/spnr tables/'
fileExt <- '.xlsx'
defaultDataFile <- paste(path, 'ARDEC_Soil_N_', sampleYear, fileExt, sep = '')

# Prompt user to select a soil data file
promptInput <-
  readline('\n\nPress <ENTER> to select a soil data file... ')
# Check input character
if(promptInput != '') errorHandler('Invalid input')
# Launch file chooser
filename <- choose.files(default = defaultDataFile)

# xlsx provides read/write functions for Excel files
library(xlsx)
# Worksheet 2 of the soil file contains the classes for each column used on
# worksheet 1
excelSheet2 <- read.xlsx2(filename, sheetIndex = 2, stringsAsFactors = FALSE)
colClassVector <- excelSheet2[, 2]
names(colClassVector) <- excelSheet2[, 1]
# Now read worksheet 1, using colClassVector
excelSheet <- read.xlsx2(filename, sheetIndex = 1,
                         colClasses = colClassVector, stringsAsFactors = FALSE)

# For each value of labNum in tableFinal...
lapply(tableFinal$labNum, function(x) {
  
  # ... find corresponding dataSheet row number
  dataSheetIndex <- which(dataSheet$labNum == x)
  # ... find corresponding tableFinal row number
  tableFinalIndex <- which(tableFinal$labNum == x)
  
  # Copy rawNO3NO2
  dataSheet$rawNO3NO2[dataSheetIndex] <- tableFinal$rawNO3NO2[tableFinalIndex]
  # Calculate adjNO3NO2
  dataSheet$adjNO3NO2[dataSheetIndex] <- 30 *
    dataSheet$rawNO3NO2[dataSheetIndex] / dataSheet$labWeight[dataSheetIndex]
  # Calculate kg_N_per_ha_NO3NO2
  dataSheet$kg_N_per_ha_NO3NO2[dataSheetIndex] <-
    dataSheet$bulkDensity[dataSheetIndex] *
    (dataSheet$depthBottom - dataSheet$depthTop) * 0.254 *
    dataSheet$adjNO3NO2[dataSheetIndex]
  
  # Copy rawNH4
  dataSheet$rawNH4[dataSheetIndex] <- tableFinal$rawNH4[tableFinalIndex]
  # Calculate adjNH4
  dataSheet$adjNH4[dataSheetIndex] <- 30 *
    dataSheet$rawNH4[dataSheetIndex] / dataSheet$labWeight[dataSheetIndex]
  # Calculate kg_N_per_ha_NH4
  dataSheet$kg_N_per_ha_NH4[dataSheetIndex] <-
    dataSheet$bulkDensity[dataSheetIndex] *
    (dataSheet$depthBottom - dataSheet$depthTop) * 0.254 *
    dataSheet$adjNH4[dataSheetIndex]
  
})  # End of function(x) and lapply statement



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
