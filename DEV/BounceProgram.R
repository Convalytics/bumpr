#################################
# bumpr package
# Jason Green : Convalytics.com
# 11/29/2014
# Version 0.0.0.0001
#################################
#
# A "bump" or "bounce" is when you would like to merge 2 (or more) separate reports 
# into one. Usually done with a VLookup in Excel. This package aims to make that
# process less painful and more accurate.
#################################

setwd("~/GitHub/bumpr/DEV")


#library(RODBC)    # To pull from database sources.
#install.packages('openxlsx')
require(openxlsx)



# Function for converting Excel dates.
excelToDate <- function(excelDate)  ### Function to convert from excel's numeric date format to an actual date format.
{
   as.POSIXct(excelDate * 60 * 60 * 24, origin="1899-12-30", tz="GMT")
}
# -------------------------------------------------------------------------------------

### Get File 1
doBounce <- function() {
   
##### Get file 1:
file1path <- readline("Enter File 1 Name: ") 
file1path <- as.character(file1path)
print(file1path)
file1path <- paste0("./",file1path)

file1sheet <- readline("Enter File 1 Sheet Index:")
file1sheet <- as.integer(file1sheet)

file1row <- readline("On Which Row Does the Data Begin?: ")
file1row <- as.integer(file1row)

file1 <- read.xlsx(file1path,sheet=file1sheet, startRow = file1row, colNames=TRUE)

##### Get file 2:
file2path <- readline("Enter File 2 Name: ") 
file2path <- as.character(file2path)
print(file2path)
file2path <- paste0("./",file2path)

file2sheet <- readline("Enter File 2 Sheet Index:")
file2sheet <- as.integer(file2sheet)

file2row <- readline("On Which Row Does the Data Begin?: ")
file2row <- as.integer(file2row)

file2 <- read.xlsx(file2path,sheet=file2sheet, startRow = file2row, colNames=TRUE)

### Bounce the files
print(names(file1))
print("-----------------")
print(names(file2))
file1join <- readline("Field to Join on From File 1:") 
file1join <- as.character(file1join)
file2join <- readline("Field to Join on From File 2:") 
file2join <- as.character(file2join)

bothreports <- merge(file1, file2, by.x=file1join, by.y=file2join, all.x=TRUE)

return(bothreports)
}
bounced <- doBounce()

#####----------------------------------------------------------------------

###########################################################################
# Run all code to this point.  
# The code below is for additional analysis/exploration of the bounced data.
###########################################################################
write.xlsx(bounced, file="BouncedFile.xlsx")       # Write data back to excel.
