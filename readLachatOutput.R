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
  # Prompt user to select an Omnion output file
  promptInput <-
    readline('\n\nPress <ENTER> to select an Omnion .htm file... ')
  # Check input character
  if(promptInput != '') fatalError('Invalid input')
  # Launch file chooser
  htmlFile <- choose.files(default = omnionFile)
  
  # Required for readHTMLTable function
  library(XML)
  # Import table from HTML file
  lachatTable <- readHTMLTable(htmlFile, which = 1, skip.rows = c(1:17),
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
  tableClean$hiVal <- NA_character_
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
    }else if(grepl('1:', tableClean$labNum[i])) {
      # Character position of dilution value
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
      # If this diluted sample is still out of range then exclude it  
      }else if(tableClean$rawNO3NO2[i] > 20 | tableClean$rawNH4[i] > 4){
        cat('Diluted sample still out of range')
        
        # ************ invoke exclusion error here *******************
      }else{
        tableClean$keep[i] <- TRUE
      }
      
      # Else if current name contains keyword 'sample' then extract lab number
    }else if(grepl('SAMPLE', tableClean$labNum[i])) {
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
    
    # Subset only rows marked to keep
    tableClean <- tableClean[tableClean$keep == TRUE, ]
    
    # Identify results that are out of range
    if(tableClean$rawNO3NO2[i] > 20) tableClean$hiVal[i] <- 'hiNO3NO2'
    if(tableClean$rawNH4[i] > 4) tableClean$hiVal[i] <- 'hiNH4'
    if(tableClean$rawNO3NO2[i] > 20 &
         tableClean$rawNH4[i] > 4) tableClean$hiVal[i] <- 'hiBoth'
    
  }  # End for-loop
  
  # Identify rows with out of range values
  hiValIndices <- which(!is.na(tableClean$hiVal))
  # Identify rows with ok values
  okValIndices <- which(is.na(tableClean$hiVal))

  # For each hiVal row
  for(hiValInd in hiValIndices) {
    # Identify labNum of current hiVal row
    currentLabNum <- tableClean$labNum[hiValInd]
    # Find labNum of row with ok values that matches the labNum of current
    # hiVal row.  The former should be a dilution.
    labNumMatchRows <- which(tableClean$labNum[okValIndices] %in% currentLabNum)
    
    # If multiple dilutions share the same lab number then throw an error
    if(length(labNumMatchRows > 1)) {
      fatalError (cat('Multiple occurrences of lab number', currentLabNum))
    # Initialize noMatch flag
    noMatch <- FALSE
    # If current hiVal sample has no corresponding dilution then exclude row
    }else if(length(labNumMatchRows = 0)) {
      tableClean$keep(hiValInd) == FALSE
      cat('** No corresponding dilution for observation', i)
      cat('\n** This observation will be excluded from data summary!\n\n')
      # Update flag
      noMatch <- TRUE
    }else{
      # Check expected dilution for presence of dilution ratio
      if(is.na(tableClean$dilution[labNumMatchRows])) fatalError('????????????')
      # Check type of hiVal
      if(tableClean$hiVal[hiValInd] == 'hiNH4') {
        tableClean$
      }
    }
    
  
### remove out of range rows  
  
  
  
  # Check for negative measured values, and set those observations to
  # keep = FALSE
  negValRows <- which(tableClean$rawNO3NO2 < 0 | tableClean$rawNH4 < 0)
  if(length(negValRows) > 0) {
    tableClean$keep[negValRows] <- FALSE
    cat('** Negative value(s) detected at these rows:', negValRows)
    cat('\n** These observations will be excluded from data summary!\n\n')
  }
  
  # Check for duplicate lab numbers in otherwise valid observations
  keepRows <- tableClean[tableClean$keep == TRUE, ]
  firstDupRow <- anyDuplicated(keepRows$labNum)
  if(firstDupRow > 0 ) {
    fatalError(cat('** More than one instance of lab number',
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
  
  # Define default soil file path and name
  path <- 'W:/R/SPNR/spnr tables/'
  fileExt <- '.xlsx'
  defaultDataFile <- paste(path, 'ARDEC_Soil_N_', sampleYear, fileExt, sep = '')
  
  # Prompt user to select a soil data file
  promptInput <-
    readline('\n\nPress <ENTER> to select a soil data file... ')
  # Check input character
  if(promptInput != '') fatalError('Invalid input')
  # Launch file chooser
  dataFile <- choose.files(default = defaultDataFile)
  
  # xlsx provides read/write functions for Excel files
  library(xlsx)
  # Worksheet 2 of the soil file contains the classes for each column used on
  # worksheet 1
  excelSheet2 <- read.xlsx2(dataFile, sheetIndex = 2, stringsAsFactors = FALSE)
  colClassVector <- excelSheet2[, 2]
  names(colClassVector) <- excelSheet2[, 1]
  # Now read worksheet 1, using colClassVector
  excelSheet <- read.xlsx2(dataFile, sheetIndex = 1,
                           colClasses = colClassVector, stringsAsFactors = FALSE)
  
  ## Extract run date from Omnion file
  #
  # Read lines of raw HTML
  htmlText = suppressWarnings(readLines(htmlFile))
  # Line 8 contains date and time
  dateTimeLine <- htmlText[8]
  # Identify character position of keyword 'Created: '
  createdStart <- gregexpr(pattern ='Created: ', dateTimeLine)
  # Read date and time
  dateTimeSub <- substr(dateTimeLine, start = createdStart[[1]][1] + 9,
                        stop = (nchar(dateTimeLine)-4))
  # Convert date and time to POSIX format
  dateTimeConverted <- strptime(dateTimeSub, '%m/%d/%Y %I:%M:%S %p')
  
  # Extract run date and copy to dataSheet
  dataSheet$runDay <- as.numeric(format(dateTimeConverted, '%d'))
  dataSheet$runMonth <- as.numeric(format(dateTimeConverted, '%m'))
  dataSheet$runYear <- as.numeric(format(dateTimeConverted, '%Y'))
  
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
  
}


################################################################################
#
# This function outputs a message indicating an excluded row
#
################################################################################

excludeRow <- function (errorMessage) {
  cat('\n\n')
  fullMessage <- paste(errorMessage, '. Program halted.', sep = '')
  stop(fullMessage, call. = FALSE)
}



################################################################################
#
# This function handles fatal errors with message output
#
################################################################################

fatalError <- function (errorMessage) {
  cat('\n\n')
  fullMessage <- paste(errorMessage, '. Program halted.', sep = '')
  stop(fullMessage, call. = FALSE)
}
