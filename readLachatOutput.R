################################################################################
# 
# The function readLachatOutput reads the data table from Omnion's HTML output
# file into a data frame, and then reformats it to facilitate data analysis.
#
################################################################################


readLachatOutput <- function() {
  # Define sample year
  sampleYear <- '2015'
  # Default file path for Omnion HTML file
  omnionFile <- paste('V:/Halvorson/ADHLab/Lachat/Lachat data/', sampleYear,
                      '/*.htm', sep = '')
  # Launch file chooser
  filename <- choose.files(default = omnionFile)
  
  # Default file path for soil weight file
  soilWeightFiles <- paste('V:/Halvorson/ADHLab/Soil Weights/2015 Soil Weights/',
                       'NO3 soils 2015 spring/*.xls', sep = '')
  # Launch file chooser - this won't be necessary if weights are read
  # automatically
  filename <- choose.files(default = soilWeights)
  
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
  
  # Capitalize all letters
  tableClean$labNum <- toupper(lachatTable$labNum)
  
  # Add columns to tableClean
  tableClean$dilution <- NA_integer_
  tableClean$keep <- FALSE
  tableClean$rowNumber <- 1:nrow(tableClean)
  
  # Reorder columns, making rowNumber the first for easier troubleshooting
  tableClean <- tableClean[c(6, 1:5)]
  
  # Initialize objects for upcoming for-loop
  checkValues <- data.frame(labNum = character(0), rawNO3NO2 = numeric(0),
                            rawNH4 = numeric(0))
  deletedRows <- data.frame(labNum = character(0), rawNO3NO2 = numeric(0),
                            rawNH4 = numeric(0))
  
  # Search for keywords in Omnion sample names, one row at a time
  for(i in 1:nrow(tableClean)) {
    
    # If current name contains keyword 'check' then copy NO3NO2 and NH4 values
    # to checkValues and update checkRowNums
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
        # Copy row of original data frame to deletedRows
        deletedRows <- rbind(deletedRows, lachatTable[i,])
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
        # Copy row of original data frame to deletedRows
        deletedRows <- rbind(deletedRows, lachatTable[i,])
      } else {
      # Otherwise keep this row
      tableClean$keep[i] <- TRUE
      }
      
    # Else (not a valid check, dilution or sample)      
    } else {
      cat('** Unreadable sample description at row', i)
      cat('\n** This observation will be excluded from data summary!\n\n')
      deletedRows <- rbind(deletedRows, lachatTable[i,])
    }
    
  }  # End for-loop
  
  # Check for negative measured values, and set those observations to
  # keep = FALSE
  negValRows <- which(tableClean$rawNO3NO2 < 0 | tableClean$rawNH4 < 0)
  if(length(negValRows) > 1) {
    tableClean$keep[negValRows] <- FALSE
    cat('** Negative value(s) detected at these rows:', negValRows)
    cat('\n** These observations will be excluded from data summary!\n\n')
    deletedRows <- rbind(deletedRows, lachatTable[negValRows,])
  }

  # Check for rows in which rawNO3NO2 rawNH4 values are out of range
  hiValRows <- which(tableClean$rawNO3NO2 > 20 | tableClean$rawNH4 > 4)
  if(length(hiValRows) > 1) {
    tableClean$keep[hiValRows] <- FALSE
    cat('** Out of range value(s) detected at these rows:', hiValRows)
    cat('\n** These observations will be excluded from data summary!\n\n')
    deletedRows <- rbind(deletedRows, lachatTable[hiValRows,])
  }

  # Check for duplicate lab numbers in otherwise valid observations
  keepRows <- tableClean[tableClean$keep == TRUE, ]
  firstDupRow <- anyDuplicated(keepRows$labNum)
  if(firstDupRow > 0 ) {
    cat('** More than one instance of lab number',
        keepRows$labNum[firstDupRow])
    stop(' Program halted.')
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
  
  # Remove unneeded columns
  dropCols <- c('rowNumber', 'dilution', 'keep')
  tableFinal <- tableFinal[, !names(tableFinal) %in% dropCols]
  
  # Add sampling year column to tableFinal
  tableFinal$sampYear <- NA_real_
  
  
  
  
  
  # Cross-reference the lab numbers in tableFinal with the ranges specified in 
  # readLachatOutput_settings.xlsx to ensure the correct sampling year.  
  settingsFile <- 'W:/R/SPNR/spnr tables/readLachatOutput_settings.xlsx'
  settingsSheet <- read.xlsx2(settingsFile, sheetIndex = 1,
                              colClasses = 'integer')
  
  # For each row entry in settingsSheet...
  for(rowNum in 1:nrow(settingsSheet)) {
    # Extract the range of lab numbers for the current row
    labNumRange <- settingsSheet$startNum[rowNum]:settingsSheet$endNum[rowNum]
    # Identify the row indices of tableFinal lab numbers in this range
    currentRangeMatch <- which(tableFinal$labNum %in% labNumRange)
    # Populate tableFinal's sampYear with the specified year
    tableFinal$sampYear[currentRangeMatch] <- settingsSheet$sampYear[rowNum]
  }
  
  
  # Join tableFinal with ARDEC_soil spreadsheet, checking for existing values in
  # the latter, in which case results are reruns.
  currentYear <- 2015
  excelYearSubset <- excelSheet[sampleYear == currentYear, ]
  excelLabNumSubset <- excelYearSubset[excelYearSubset$labNum %in%
                                         tableCleaned$labNum]
  

  

##############################################################################
#
# Insert code here to calculate adjusted nitrate & ammonium concentrations,
# as well as lb N per ac and kg N per ha for each depth segment.  These
# calculations require soil bulk density, soil lab weight and depth segment
# values, which are available in the soil file.
# 
##############################################################################

# Spawn file chooser to have user select appropriate soil file
path <- 'W:/R/SPNR/spnr tables/'
fileExt <- '.xlsx'
omnionFile <- paste(path, 'ARDEC_soil', fileExt, sep = '')
filename <- choose.files(default = omnionFile)

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

# Merge lachatTableCleaned with excelSheet
merge(x = lachatTableCleaned, y = excelSheet, by.x = )


# Calcuate adjusted NO3/NO2 and NH4 values











# Copy objects of interest to global environment
lachatTable <<- lachatTable
checkValues <<- checkValues
lachatTableCleaned <<- lachatTableCleaned

}  # Close readLachatOutput



