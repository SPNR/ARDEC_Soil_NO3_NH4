################################################################################
#
# Function chooseExcelFile
# Spawn a file chooser for selecting appropriate Excel soil file
#
################################################################################

chooseExcelFile <- function() {
  
  path <- 'W:/R/SPNR/test data/'
  fileExt <- '.xlsx'
  defaultFile <- paste(path, '2014_R1_and_R2_summer_NO3_NH4_analysis',
                       fileExt, sep = '')
  filename <- choose.files(default = defaultFile)
  
  # xlsx provides read/write functions for Excel files
  library(xlsx)
  excelSheet <- read.xlsx2(filename, sheetIndex = 1, colClasses =
                             c(rep('numeric', 8), rep('character', 2),
                               rep('numeric', 4)))
  excelSheet  # Return value
}

################################################################################
#
# main code
#
################################################################################

excelSheet <- chooseExcelFile()

# Replace original values with rerun values that satisfy given condition
for(i in 1:nrow(excelSheet)) {
  # If current NO3 rerun value is not NA...
  if(!is.na(excelSheet$concNO3rr[i])) {
    # Are NO3 and NH4 rerun values for each observation both less than their
    # corresponding original measurements?
    if((excelSheet$concNO3rr[i] < excelSheet$concNO3[i]) &
         (excelSheet$concNH4rr[i] < excelSheet$concNH4[i])) {
      # If so, substitute rerun values
      excelSheet$concNO3[i] <- excelSheet$concNO3rr[i]
      excelSheet$concNH4[i] <- excelSheet$concNH4rr[i]
    }
  }
}

# Change calendar day to days after fertilization
excelSheet$sampDay <- excelSheet$sampDay + 25

# Columns to remove
rmCols <- names(excelSheet) %in% c('sampYear', 'sampMonth')
excelSheet <- excelSheet[,!rmCols]

# Subset by sampling day
day1 <- subset(excelSheet, sampDay == 27)
day2 <- subset(excelSheet, sampDay == 43)

# NO3 calculations

# Calculate mean by treatment, study and suffix
meanNO3_day1 <- aggregate(concNO3 ~ treatment + study + suffix + depthTop,
                          day1, mean)
# Order by study, treatment and suffix
meanNO3_day1 <- meanNO3_day1[order(meanNO3_day1$study, meanNO3_day1$treatment,
                                   meanNO3_day1$depthTop, meanNO3_day1$suffix),]
# Calculate sd by treatment, study and suffix
sdNO3_day1 <- aggregate(concNO3 ~ treatment + study + suffix + depthTop,
                        day1, sd)
# Order by study, treatment and suffix
sdNO3_day1 <- sdNO3_day1[order(sdNO3_day1$study, sdNO3_day1$treatment,
                               sdNO3_day1$depthTop, sdNO3_day1$suffix),]

# Calculate mean by treatment, study and suffix
meanNO3_day2 <- aggregate(concNO3 ~ treatment + study + suffix + depthTop,
                          day2, mean)
# Order by study, treatment and suffix
meanNO3_day2 <- meanNO3_day2[order(meanNO3_day2$study, meanNO3_day2$treatment,
                                   meanNO3_day2$depthTop, meanNO3_day2$suffix),]
# Calculate sd by treatment, study and suffix
sdNO3_day2 <- aggregate(concNO3 ~ treatment + study + suffix + depthTop,
                        day2, sd)
# Order by study, treatment and suffix
sdNO3_day2 <- sdNO3_day2[order(sdNO3_day2$study, sdNO3_day2$treatment,
                               sdNO3_day2$depthTop, sdNO3_day2$suffix),]

# Rename columns for day 1
colnames(meanNO3_day1)[5] <- 'meanConcNO3'
NO3_day1 <- meanNO3_day1
colnames(sdNO3_day1)[5] <- 'sdConcNO3'
NO3_day1$sdConcNO3 <- sdNO3_day1$sdConcNO3

# Rename columns for day 2
colnames(meanNO3_day2)[5] <- 'meanConcNO3'
NO3_day2 <- meanNO3_day2
colnames(sdNO3_day2)[5] <- 'sdConcNO3'
NO3_day2$sdConcNO3 <- sdNO3_day2$sdConcNO3

# NH4 calculations

# Calculate mean by treatment, study and suffix
meanNH4_day1 <- aggregate(concNH4 ~ treatment + study + suffix + depthTop,
                          day1, mean)
# Order by study, treatment and suffix
meanNH4_day1 <- meanNH4_day1[order(meanNH4_day1$study, meanNH4_day1$treatment,
                                   meanNH4_day1$depthTop, meanNH4_day1$suffix),]
# Calculate sd by treatment, study and suffix
sdNH4_day1 <- aggregate(concNH4 ~ treatment + study + suffix + depthTop,
                        day1, sd)
# Order by study, treatment and suffix
sdNH4_day1 <- sdNH4_day1[order(sdNH4_day1$study, sdNH4_day1$treatment,
                               sdNH4_day1$depthTop, sdNH4_day1$suffix),]

# Calculate mean by treatment, study and suffix
meanNH4_day2 <- aggregate(concNH4 ~ treatment + study + suffix + depthTop,
                          day2, mean)
# Order by study, treatment and suffix
meanNH4_day2 <- meanNH4_day2[order(meanNH4_day2$study, meanNH4_day2$treatment,
                                   meanNH4_day2$depthTop, meanNH4_day2$suffix),]
# Calculate sd by treatment, study and suffix
sdNH4_day2 <- aggregate(concNH4 ~ treatment + study + suffix + depthTop,
                        day2, sd)
# Order by study, treatment and suffix
sdNH4_day2 <- sdNH4_day2[order(sdNH4_day2$study, sdNH4_day2$treatment,
                               sdNH4_day2$depthTop, sdNH4_day2$suffix),]

# Rename columns for day 1
colnames(meanNH4_day1)[5] <- 'meanConcNH4'
NH4_day1 <- meanNH4_day1
colnames(sdNH4_day1)[5] <- 'sdConcNH4'
NH4_day1$sdConcNH4 <- sdNH4_day1$sdConcNH4

# Rename columns for day 2
colnames(meanNH4_day2)[5] <- 'meanConcNH4'
NH4_day2 <- meanNH4_day2
colnames(sdNH4_day2)[5] <- 'sdConcNH4'
NH4_day2$sdConcNH4 <- sdNH4_day2$sdConcNH4


# Vector of data frames
df <- c(NO3_day1, NO3_day2, NH4_day1, NH4_day2)  # Necessary?

# Add a column for p-values
pValue <- rep('-', nrow(NO3_day1))
NO3_day1$pValue <- pValue
NO3_day2$pValue <- pValue
NH4_day1$pValue <- pValue
NH4_day2$pValue <- pValue

# Sort day1 and day2 to prepare for t-tests
day1 <- day1[order(day1$study, day1$depthTop, day1$treatment, day1$suffix),]
day2 <- day2[order(day2$study, day2$depthTop, day2$treatment, day2$suffix),]

# Initialize loop counters
i <- 1
lineCount <- 2
# Perform t-tests for day 1 NO3
while(lineCount <= 72) {
  x <- c(day1$concNO3[i], day1$concNO3[i + 1], day1$concNO3[i + 2])
  y <- c(day1$concNO3[i + 3], day1$concNO3[i + 4], day1$concNO3[i + 5])
  test <- t.test(x, y)
  NO3_day1$pValue[lineCount] <- round(test$p.value, digits = 3)
  print(i)
  print(lineCount)
  lineCount <- lineCount+ 2
  i <- i + 6
}

# Initialize loop counters
i <- 1
lineCount <- 2
# Perform t-tests for day 2 NO3
while(lineCount <= 72) {
  x <- c(day2$concNO3[i], day2$concNO3[i + 1], day2$concNO3[i + 2])
  y <- c(day2$concNO3[i + 3], day2$concNO3[i + 4], day2$concNO3[i + 5])
  test <- t.test(x, y)
  NO3_day2$pValue[lineCount] <- round(test$p.value, digits = 3)
  print(i)
  print(lineCount)
  lineCount <- lineCount+ 2
  i <- i + 6
}

# Initialize loop counters
i <- 1
lineCount <- 2
# Perform t-tests for day 1 NH4
while(lineCount <= 72) {
  x <- c(day1$concNH4[i], day1$concNH4[i + 1], day1$concNH4[i + 2])
  y <- c(day1$concNH4[i + 3], day1$concNH4[i + 4], day1$concNH4[i + 5])
  test <- t.test(x, y)
  NH4_day1$pValue[lineCount] <- round(test$p.value, digits = 3)
  print(i)
  print(lineCount)
  lineCount <- lineCount+ 2
  i <- i + 6
}

# Initialize loop counters
i <- 1
lineCount <- 2
# Perform t-tests for day 2 NH4
while(lineCount <= 72) {
  x <- c(day2$concNH4[i], day2$concNH4[i + 1], day2$concNH4[i + 2])
  y <- c(day2$concNH4[i + 3], day2$concNH4[i + 4], day2$concNH4[i + 5])
  test <- t.test(x, y)
  NH4_day2$pValue[lineCount] <- round(test$p.value, digits = 3)
  print(i)
  print(lineCount)
  lineCount <- lineCount+ 2
  i <- i + 6
}

# Write files
path <- 'W:/R/SPNR/test data/'
fileExt <- '.xlsx'
sum1 <- paste(path, 'sum1', fileExt, sep = '')
sum2 <- paste(path, 'sum2', fileExt, sep = '')
sum3 <- paste(path, 'sum3', fileExt, sep = '')
sum4 <- paste(path, 'sum4', fileExt, sep = '')

write.xlsx2(NO3_day1, file = sum1, sheetName = 'NO3_day1',
            showNA = FALSE, row.names = FALSE)
write.xlsx2(NO3_day2, file = sum2, sheetName = 'NO3_day2',
            showNA = FALSE, row.names = FALSE)
write.xlsx2(NH4_day1, file = sum3, sheetName = 'NH4_day1',
            showNA = FALSE, row.names = FALSE)
write.xlsx2(NH4_day2, file = sum4, sheetName = 'NH4_day2', 
            showNA = FALSE, row.names = FALSE)
