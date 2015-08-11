################################################################################
#
#  Exploratory data analysis for ARDEC soil
#
################################################################################

# Spawn file chooser for user to select appropriate soil file
path <- 'W:/R/SPNR/spnr tables/'
fileExt <- '.xlsx'
defaultFile <- paste(path, 'ARDEC_Soil_N_2015', fileExt, sep = '')
soilsFile <- choose.files(default = defaultFile)

# xlsx provides read/write functions for Excel files
library(xlsx)
# Worksheet 2 contains the classes for each column used on worksheet 1
dataSheet2 <- read.xlsx2(soilsFile, sheetIndex = 2, stringsAsFactors = FALSE)
colClassVector <- dataSheet2[, 2]
names(colClassVector) <- dataSheet2[, 1]
# Now read worksheet 1, using colClassVector
dataSheet <- read.xlsx2(soilsFile, sheetIndex = 1, colClasses = colClassVector,
                        stringsAsFactors = FALSE)

# Add columns to adjust N concentrations to a per-inch-of-depth basis for
# scale consistency in scatter plots.
#
# Initialize new columns
dataSheet$kg_N_per_ha_per_inch_NO3NO2 <- 0
dataSheet$kg_N_per_ha_per_inch_NH4 <- 0

# Populate new columns
for(i in 1:nrow(dataSheet)) {
  if(dataSheet$depthBottom[i] == 3 | dataSheet$depthBottom[i] == 6) {
    dataSheet$kg_N_per_ha_per_inch_NO3NO2[i] <-
      dataSheet$kg_N_per_ha_NO3NO2[i] / 3
    dataSheet$kg_N_per_ha_per_inch_NH4[i] <-
      dataSheet$kg_N_per_ha_NH4[i] / 3
  }else if(dataSheet$depthBottom[i] == 12) {
    dataSheet$kg_N_per_ha_per_inch_NO3NO2[i] <-
      dataSheet$kg_N_per_ha_NO3NO2[i] / 6
    dataSheet$kg_N_per_ha_per_inch_NH4[i] <-
      dataSheet$kg_N_per_ha_NH4[i] / 6
  }else {
    dataSheet$kg_N_per_ha_per_inch_NO3NO2[i] <-
      dataSheet$kg_N_per_ha_NO3NO2[i] / 12
    dataSheet$kg_N_per_ha_per_inch_NH4[i] <-
      dataSheet$kg_N_per_ha_NH4[i] / 12
  }
}

# Change some column classes to 'factor' for simplified plotting
dataSheet$rep <- as.factor(dataSheet$rep)

# Change trtLevel to a more readable form for plotting (dmp will be a special
# case, a few code sections down)
dataSheet$trtLevel <- paste('N treatment level',
                            as.character(dataSheet$trtLevel))

# Subset by study/suffix (exluding SCS from this summary)
stCC  <- subset(dataSheet, study == 'Strip Till' & plotSuffix == 'CC')
stSSC <- subset(dataSheet, study == 'Strip Till' & plotSuffix == 'SS/C')
dmp  <- subset(dataSheet, study == 'DMP')
r1E  <- subset(dataSheet, study == 'Rotation 1' & plotSuffix == 'E')
r1W  <- subset(dataSheet, study == 'Rotation 1' & plotSuffix == 'W')
r2E  <- subset(dataSheet, study == 'Rotation 2' & plotSuffix == 'E')
r2W  <- subset(dataSheet, study == 'Rotation 2' & plotSuffix == 'W')
r3  <- subset(dataSheet, study == 'Rotation 3')
r4  <- subset(dataSheet, study == 'Rotation 4')
r5E  <- subset(dataSheet, study == 'Rotation 5' & plotSuffix == 'E')
r5W  <- subset(dataSheet, study == 'Rotation 5' & plotSuffix == 'W')



# In each subset, create column 'fullStudy' that is a concatenation of plot name
# and plot suffix.  If no suffix exists, it consists only of plot name.  This
# provides consistency in plotting among datasubsets.
#
# Specify no suffixes for these studies
studyOnlyList <- list(dmp = dmp, r3 = r3, r4 = r4)
# Specify suffixes for these studies
studySfxList <- list(stCC = stCC, stSSC = stSSC, r1E = r1E, r1W = r1W,
                     r2E = r2E, r2W = r2W, r5E = r5E, r5W = r5W)
lapply(studyOnlyList, function(x) x$fullStudy <- x$study)
lapply(studySfxList, function(x) x$fullStudy <- paste(x$study, x$plotSuffix))

# # In each subset, create column 'fullStudy' that is a concatenation of plot name
# # and plot suffix.  If no suffix exists, it consists only of plot name.  This
# # provides consistency in plotting among datasubsets.
# stCC$fullStudy <- paste(stCC$study, stCC$plotSuffix)
# stSSC$fullStudy <- paste(stSSC$study, stSSC$plotSuffix)
# dmp$fullStudy <- dmp$study
# r3$fullStudy <- r3$study
# r4$fullStudy <- r4$study
# r1E$fullStudy <- paste(r1E$study, r1E$plotSuffix)
# r1W$fullStudy <- paste(r1W$study, r1W$plotSuffix)
# r2E$fullStudy <- paste(r2E$study, r2E$plotSuffix)
# r2W$fullStudy <- paste(r2W$study, r2W$plotSuffix)
# r5E$fullStudy <- paste(r5E$study, r5E$plotSuffix)
# r5W$fullStudy <- paste(r5W$study, r5W$plotSuffix)

# Replace trtLevel with trtType for DMP, as all plots have same trtLevel.  This
# provides consistency in plotting among data subsets.
dmp$trtLevel <- paste('N source type:', dmp$trtType1)

# This list will have a different order than the commented alternative below
#subsetList <- list(studyOnlyList, studySfxList)

 subsetList <- list(stCC = stCC, stSSC = stSSC, dmp = dmp, r1E = r1E, r1W = r1W,
                    r2E = r2E, r2W = r2W, r3 = r3, r4 = r4, r5E = r5E, r5W = r5W)

library(ggplot2)

# Create NO3/NO2 plots
lapply(subsetList, function(df) {
  
  # Generate NO3/NO2 plot
  plotTitle <- paste(df$study[1], df$plotSuffix[1])
  if(df$study[1] == 'DMP') plotTitle <- df$study[1]
  g1 <- ggplot(df, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                       color = rep, group = rep)) 
  g1 <- g1 + geom_point(alpha = 0.8, size = 3)
  g1 <- g1 + facet_wrap(~trtLevel, ncol = 2)
  g1 <- g1 + labs(x = 'depthBottom (in)', y = 'kg_N_per_ha_per_inchDepth_NO3NO2',
                  title = plotTitle)
  g1 <- g1 + theme(plot.title = element_text(size = 24, vjust = 2),
                   axis.title.x = element_text(size = 14),
                   axis.title.y = element_text(size = 14, vjust = 1))
  g1 <- g1 + scale_x_continuous(breaks = seq(0, 72, 12))
  g1
  
})


# Script ends here

#------------------------------------------

# Plot each subset individually

library(ggplot2)

# Create facetted NO3/NO2 plots separated by depth
g <- ggplot(stCC, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                    color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)', y = 'kg_N_per_ha_per_inchDepth_NO3NO2',
              title = plotTitle)
g <- g + ggtitle('Strip Till CC')
g

g <- ggplot(stSSC, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                      color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Strip Till SS/C')
g

g <- ggplot(dmp, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                       color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('DMP')
g

g <- ggplot(r3, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                       color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 3')
g

g <- ggplot(r4, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                       color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 4')
g

g <- ggplot(r1E, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                       color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 1E')
g

g <- ggplot(r1W, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 1W')
g

g <- ggplot(r2E, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 2E')
g

g <- ggplot(r2W, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 2W')
g

g <- ggplot(r5E, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 5E')
g

g <- ggplot(r5W, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NO3NO2,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NO3NO2')
g <- g + ggtitle('Rot 5W')
g


# Create facetted NH4 plots separated by depth
g <- ggplot(stCC, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                      color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Strip Till CC')
g

g <- ggplot(stSSC, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                       color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Strip Till SS/C')
g

g <- ggplot(dmp, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('DMP')
g

g <- ggplot(r3, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                    color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 3')
g

g <- ggplot(r4, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                    color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 4')
g

g <- ggplot(r1E, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 1E')
g

g <- ggplot(r1W, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 1W')
g

g <- ggplot(r2E, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 2E')
g

g <- ggplot(r2W, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 2W')
g

g <- ggplot(r5E, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 5E')
g

g <- ggplot(r5W, aes(x = depthBottom, y = kg_N_per_ha_per_inch_NH4,
                     color = rep, group = rep)) 
g <- g + geom_point(alpha = 0.7, size = 3)
g <- g + facet_wrap(~trtLevel, ncol = 2)
g <- g + labs(x = 'depthBottom (in)',
              y = 'kg_N_per_ha_per_inchDepth_NH4')
g <- g + ggtitle('Rot 5W')
g


# ------------------------- Boxplots below ---------------------------------


# Convert depthBottom to factor for each subset individually for boxplots
st$depthBottom <- as.factor(st$depthBottom)
dmp$depthBottom <- as.factor(dmp$depthBottom)
r3$depthBottom <- as.factor(r3$depthBottom)
r4$depthBottom <- as.factor(r4$depthBottom)
r1E$depthBottom <- as.factor(r1E$depthBottom)
r1W$depthBottom <- as.factor(r1W$depthBottom)
r2E$depthBottom <- as.factor(r2E$depthBottom)
r2W$depthBottom <- as.factor(r2W$depthBottom)


# Now create 8 boxplots for soil NO3/NO2 results.  These plots do not
# separate by treatement level.
g <- ggplot(st, aes(x = depthBottom, y = adjNO3NO2, fill = depthBottom))
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Strip Till')
g

g <- ggplot(dmp, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('DMP')
g

g <- ggplot(r3, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Rotation 3')
g

g <- ggplot(r4, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Rotation 4')
g

g <- ggplot(r1E, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Rotation 1 East')
g

g <- ggplot(r1W, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Rotation 1 West')
g

g <- ggplot(r2E, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Rotation 2 East')
g

g <- ggplot(r2W, aes(depthBottom, adjNO3NO2, fill = depthBottom)) 
g <- g + geom_boxplot() + guides(fill = FALSE)
g <- g + ggtitle('Rotation 2 West')
g


