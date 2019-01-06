install.packages("xlsx")
library("xlsx")

DataDump <- readLines("CrossText.txt")
PropertyType <- readLines("PropertyType.txt")

DataExtractStart <- grep("Required", DataDump) + 1
DataExtractStop <- grep("Available", DataDump) - 1

FieldLocations <- grep("As of", DataDump) - 1
Fields <- DataDump[FieldLocations]

PropertyCheck <- matrix('-', nrow = length(Fields), ncol = length(PropertyType))
colnames(PropertyCheck) <- sort(PropertyType)
rownames(PropertyCheck) <- sort(Fields)

for (i in 2:length(Fields)) {
  TempRequiredCode <- toString(DataDump[DataExtractStart[i]:DataExtractStop[i]])

  for (i2 in 1:length(PropertyType)) {
    if (grepl(PropertyType[i2], TempRequiredCode)) {
      PropertyCheck[i,i2] = paste("Y")
    }
  }

  TempRequiredCode <- c()
}

# Excel R Code
PropertyCheckWorkbook <- createWorkbook(type = "xlsx")
PropertyCheckSheet <- createSheet(PropertyCheckWorkbook, sheetName = "Property Check")

CellStyle <- CellStyle(PropertyCheckWorkbook) + Alignment(horizontal = "ALIGN_CENTER")
cList <- rep(list(CellStyle), ncol(PropertyCheck))
names(cList) <- seq(1, ncol(PropertyCheck), by = 1)

ColumnNameStyle <- CellStyle(PropertyCheckWorkbook) + Font(PropertyCheckWorkbook, isBold = TRUE) + 
  Alignment(horizontal = "ALIGN_CENTER")
RowNameStyle <- CellStyle(PropertyCheckWorkbook) + Font(PropertyCheckWorkbook, isBold = TRUE)

addDataFrame(PropertyCheck, PropertyCheckSheet, colStyle = cList, 
             colnamesStyle = ColumnNameStyle, rownamesStyle = RowNameStyle)

setColumnWidth(PropertyCheckSheet, colIndex = 1, colWidth = max(nchar(Fields) + 2))
setColumnWidth(PropertyCheckSheet, colIndex = 2:(ncol(PropertyCheck) + 1), colWidth = max(nchar(PropertyType) + 2))

saveWorkbook(PropertyCheckWorkbook, "Property Check.xlsx")