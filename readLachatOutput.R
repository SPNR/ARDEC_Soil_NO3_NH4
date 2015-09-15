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
  omnionFolder <- paste('V:/Halvorson/ADHLab/Lachat/Lachat data/', sampleYear,
                        '/', sep = '')
  # Prompt user to select an Omnion output file
  promptInput <-
    readline('\n\nPress <ENTER> to select a folder containing .htm files... ')
  # Choose folder containing Omnion .htm files
  folder <- choose.dir(default = omnionFolder)
  if(is.na(folder)) errorHandler('No folder selected')
  # Print status message
  cat('\n\n...reading HTML files...')
  # Capture filenames with htm and html extensions into list htmlFiles
  htmlFiles <- Sys.glob(paste(folder, '/*.htm*', sep = ''))
  
  # Required for readHTMLTable function
  library(XML)
  # Create a list to contain individual DFs of HTML tables
  tableList <- list()
  # For each filename...
  for(i in 1:length(htmlFiles)) {
    # Read the HTML data table
    tableList[[i]] <- readHTMLTable(htmlFiles[i], which = 1, skip.rows =
                          c(1:17), colClasses = c('character',
                          rep('numeric', 3)), stringsAsFactors = FALSE)
    # Read lines of raw HTML
    htmlText = suppressWarnings(readLines(htmlFiles[i]))
    # Line 8 contains date and time
    dateTimeLine <- htmlText[8]
    # Identify character position of keyword 'Created: '
    createdStart <- gregexpr(pattern ='Created: ', dateTimeLine)
    # Read date and time
    dateTimeSub <- substr(dateTimeLine, start = createdStart[[1]][1] + 9,
                          stop = (nchar(dateTimeLine) - 4))
    # Extract run date and copy to current table
    tableList[[i]]$runDate <- dateTimeSub
    
  }  # End for-loop

  # Concatenate list of DFs into a single DF
  lachatTable <- NULL
  # The data.table package provides rbindlist()
  library(data.table)
  lachatTable <- rbindlist(tableList)
  
  # Convert date and time to POSIX format
  lachatTable$runDate <- as.POSIXct(lachatTable$runDate,
                                    format = '%m/%d/%Y %I:%M:%S %p')
  
  # Delete second column (Rep column)
  lachatTable$V2 <- NULL
  # Rename remaining columns
  names(lachatTable) <- c('labNum', 'rawNO3NO2', 'rawNH4', 'runDate')
  # Create a working copy of lachatTable
  tableClean <- lachatTable
  # Capitalize all letters to facilitate pattern matching
  tableClean$labNum <- toupper(lachatTable$labNum)
  
  # Add columns to tableClean
  tableClean$dilution <- NA_integer_
  tableClean$hiVal <- NA_character_
  tableClean$keep <- FALSE
  
  # Initialize data frame to contain soil check results
  checkValues <- data.frame(labNum = character(0), rawNO3NO2 = numeric(0),
                            rawNH4 = numeric(0), runDate = numeric(0),
                            stringsAsFactors = FALSE)
  
  ##################
  #
  # Error messages below for "unreadable sample description" and "invalid lab
  # number" can refer to row numbers that will be different in the modified
  # tableClean due to removed 'keep=FALSE' lines.
  #
  
  # Search for keywords in Omnion sample names, one row at a time
  for(i in 1:nrow(tableClean)) {
    
    # If current name contains keyword 'check' then copy NO3NO2 and NH4 values
    # to checkValues
    if(grepl('CHECK', tableClean[i, 1])) {
      checkValues <- rbind(checkValues,
                           c(tableClean$rawNO3NO2[i], tableClean$rawNH4[i],
                             tableClean$runDay[i], tableClean$runMonth[i],
                             tableClean$runYear[i]))
      # Else if current name contains '1:' then extract the dilution value and
      # lab number
    }else if(grepl('1:', tableClean$labNum[i])) {
      # Character position of dilution value
      dilPos <- regexpr('1:', tableClean$labNum[i])
      # Read dilution value
      tableClean$dilution[i] <- suppressWarnings(
        as.numeric(substring(tableClean$labNum[i], dilPos + 2)))
      # Parse previous characters to extract lab number.  Sample
      # descriptor is assumed to be of the form "1234 1:8"
      labNumText <- substring(tableClean$labNum[i], first = 1,
                              last = dilPos - 2)
      # Transform labNumText into a numeric value
      #       tableClean$labNum[i] <- suppressWarnings(
      #         as.numeric(paste(unlist(labNumText), collapse = '')))
      tableClean$labNum[i] <- suppressWarnings(as.numeric(labNumText))
      
      # If current dilution or labNum is NA then print error message
      if(is.na(tableClean$dilution[i]) | is.na(tableClean$labNum[i])) {
        cat("** Invalid dilution value or lab number at sample",
            tableClean$labNum[i])
        cat('\n** This observation will be excluded from data summary!\n\n')
        # If this diluted sample is still out of range then exclude it  
      }else if(tableClean$rawNO3NO2[i] > 20 | tableClean$rawNH4[i] > 4){
        cat('Diluted sample', tableClean$labNum[i], 'still out of range')
        cat('\n** This observation will be excluded from data summary!\n\n')
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
      }else{
        # Otherwise keep this row
        tableClean$keep[i] <- TRUE
      }
      
      # Else (not a valid check, dilution or sample)      
    }else{
      cat('** Unreadable sample description at row', i)
      cat('\n** This observation will be excluded from data summary!\n\n')
    }
    
    if(tableClean$keep[i] == TRUE) {
      # Identify results with a single out-of-range value, as the corresponding
      # in-range value will be take the place of the diluted in-range value
      if(tableClean$rawNO3NO2[i] > 20 &
           tableClean$rawNH4[i] <= 4) tableClean$hiVal[i] <- 'hiNO3NO2'
      if(tableClean$rawNO3NO2[i] <= 20 &
           tableClean$rawNH4[i] > 4) tableClean$hiVal[i] <- 'hiNH4'
    }  # End if-statement
    
  }  # End for-loop
  
  # Simplify tableClean by subsetting only rows marked to keep
  tableClean <- tableClean[tableClean$keep == TRUE, ]
  
  # Indices of rows with out-of-range values
  hiValRows <- which(!is.na(tableClean$hiVal))
  # Indices of rows with ok values
  okValRows <- which(is.na(tableClean$hiVal))
  # Indices of rows having out-of-range values, but no corresponding dilutions
  noDilRows <- numeric(0)
  
  # For each hiVal row
  for(hiValRow in hiValRows) {
    # Identify labNum of current row
    currentLabNum <- tableClean$labNum[hiValRow]
    # Find labNum of the row with ok values that matches the labNum of the
    # current hiVal row.  The former should be a dilution.
    labNumMatchRows <- which(tableClean$labNum[okValRows] %in% currentLabNum)
    
    # If multiple in-range dilutions share the same lab number then throw an
    # error.  No out-of-range dilutions (underdilutions) should occur here
    # because they were filtered out in the previous for-loop
    if(length(labNumMatchRows > 1)) {
      fatalError (cat('Multiple dilutions with lab number', currentLabNum))
      # Else if current hiVal sample has no corresponding dilution then exclude
      # current row
    }else if(length(labNumMatchRows == 0)) {
      tableClean$keep(hiValRow) == FALSE
      # Add this row to noDilRows
      noDilRows <- c(noDilRows, tableClean$labNum[hiValRow])
      # Else a single match exists: length(labNumMatchRows) = 1
    }else{
      # Check expected dilution for presence of dilution ratio
      if(is.na(tableClean$dilution[labNumMatchRows])) fatalError(cat(
        'Dilution ratio not found for lab number', tableClean$labNum[hiValRow]))
      # If hiVal is an ammonium... 
      if(tableClean$hiVal[hiValRow] == 'hiNH4') {
        # ... then retain the original nitrate/nitrite value
        tableClean$rawNO3NO2[labNumMatchRows] <-
          tableClean$rawNO3NO2[hiValRow]
        # Else if hiVal is a nitrate/nitrite...
      }else if(tableClean$hiVal[hiValRow] == 'hiNO3NO2') {
        # ... then retain original ammonium value
        tableClean$rawNH4[labNumMatchRows] <-
          tableClean$rawNH4[hiValRow]
      }
      # Label current row for deletion
      tableClean$keep[hiValRow] <- FALSE
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
  defaultDataFile <- paste(path, 'ARDEC_Soil_N_', sampleYear, '_test', fileExt,
                           sep = '')
  
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
    
    # Copy run date
    dataSheet$runDay[dataSheetIndex] <- tableFinal$runDay[tableFinalIndex]
    dataSheet$runMonth[dataSheetIndex] <- tableFinal$runMonth[tableFinalIndex]
    dataSheet$runYear[dataSheetIndex] <- tableFinal$runYear[tableFinalIndex]
    
  })  # End of function(x) and lapply statement
  
}  # End of readLachatOutput()



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
