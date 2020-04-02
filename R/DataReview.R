#' @title DataReviewFile
#'
#' @import openxlsx
#'
#' @description Exports the data review of all tables in lists to an excel file
#'
#' @param pathfile: the path of the xlsx file where you want your output
#' listTables: list of the names of all the tables you want to include in your review
#'
#' @return NULL
#'
#' @examples
#'
#' @export DataReviewFile

DataReviewFile <- function(pathfile, listTables){
  workbook <- createWorkbook()
  addWorksheet(workbook, sheetName = "General")
  addWorksheet(workbook, sheetName = "Category")
  addWorksheet(workbook, sheetName = "Numeric")
  rowgen = 2
  rowcat = 2
  rownum = 2
  for (elem in listTables){
    DRgeneral <- data.frame(diagnose(get(elem)))
    DRcategory <- data.frame(diagnose_category(get(elem)))
    DRnumeric <- data.frame(diagnose_numeric(get(elem)))
    writeData(workbook, sheet = "General", startCol = 2, startRow = rowgen , x = elem)
    writeDataTable(workbook, sheet = "General", x = DRgeneral, startCol = 2, startRow = (rowgen + 2))
    if (nrow(DRcategory)>0){
      writeData(workbook, sheet = "Category", startCol = 2, startRow = rowgen , x = elem)
      writeDataTable(workbook, sheet = "Category", x = DRcategory, startCol = 2, startRow = (rowcat + 2))
    }
    if (nrow(DRnumeric)>0){
      writeData(workbook, sheet = "Category", startCol = 2, startRow = rowgen , x = elem)
      writeDataTable(workbook, sheet = "Numeric", x = DRnumeric, startCol = 2, startRow = (rownum + 2))
    }
    rowgen = rowgen + nrow(DRgeneral) + 4
    rowcat = rowcat + nrow(DRcategory) + 4
    rownum = rownum + nrow(DRnumeric) + 4
  }
  saveWorkbook(workbook, file = pathfile)
}
