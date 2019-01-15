install.packages("xlsx")
library("xlsx")

# Read data
#---------------------------------------------------------------------------
data_dump <- readLines("cross-text.txt")
prop_type <- readLines("property-types.txt")

data_ext_beg <- grep("Required", data_dump) + 1
data_ext_end <- grep("Available", data_dump) - 1

field_locs <- grep("As of", data_dump) - 1
fields <- data_dump[field_locs]

prop_chk <- matrix('-', nrow = length(fields), ncol = length(prop_type))
colnames(prop_chk) <- sort(prop_type)
rownames(prop_chk) <- sort(fields)

for (i in 2:length(fields)) {
  temp <- toString(data_dump[data_ext_beg[i]:data_ext_end[i]])

  for (i2 in 1:length(prop_type)) {
    if (grepl(prop_type[i2], temp)) {
      prop_chk[i, i2] = paste("Y")
    }
  }

  temp <- c()
}

# Excel
#---------------------------------------------------------------------------
prop_chk_wb <- createWorkbook(type = "xlsx")
prop_chk_sh <- createSheet(prop_chk_wb, sheetName = "Property Check")

cell_sty <- CellStyle(prop_chk_wb) + Alignment(horizontal = "ALIGN_CENTER")
clist <- rep(list(cell_sty), ncol(prop_chk))
names(clist) <- seq(1, ncol(prop_chk), by = 1)

col_nam_sty <- CellStyle(prop_chk_wb) + Font(prop_chk_wb, isBold = TRUE) + 
  Alignment(horizontal = "ALIGN_CENTER")
row_nam_sty <- CellStyle(prop_chk_wb) + Font(prop_chk_wb, isBold = TRUE)

addDataFrame(prop_chk, prop_chk_sh, colStyle = clist, 
             colnamesStyle = col_nam_sty, rownamesStyle = row_nam_sty)

setColumnWidth(prop_chk_sh, colIndex = 1, colWidth = max(nchar(fields) + 2))
setColumnWidth(prop_chk_sh, colIndex = 2:(ncol(prop_chk) + 1), 
               colWidth = max(nchar(prop_type) + 2))

saveWorkbook(prop_chk_wb, "property-check.xlsx")