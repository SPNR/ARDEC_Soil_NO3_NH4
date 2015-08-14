################################################################################
#
# Read multiple soil weight files, and join them together in a data frame
#
################################################################################

# Test edit

# Establish default folder for soil weights
soilWeightFiles <- Sys.glob(paste('V:/Halvorson/ADHLab/Soil Weights/',
                                  '2015 Soil Weights/',
                                  'NO3 soils 2015 spring/*.*', sep = ''))

# Concatenate the contents of all files in the default folder
i <- 1:length(soilWeightFiles)
for(i in 1:length(soilWeightFiles)){
  soilWeightList <- lapply(soilWeightFiles, function(x) read.xlsx2(x,
                      sheetIndex = 1, stringsAsFactors = FALSE, startRow = 4,
                      colIndex = 1:2, header=FALSE, colClasses =
                      c('numeric', 'numeric')))
}

# Transform the list of two-column data frames into a single data frame with
# two columns
soilWeightDF <- do.call('rbind', soilWeightList)

names(soilWeightDF) <- c('labNum', 'weight')

# Delete rows containing '' values
soilWeightDF <- soilWeightDF[!if.na(soilWeightDF$X2), ]




# necessary?
df1 <- as.data.frame(soilWeightList)

# Example expression to access the 157th element of the X2 column of the 2nd
# data frame in soilWeightList:
soilWeightList[[2]]$X2[157]
