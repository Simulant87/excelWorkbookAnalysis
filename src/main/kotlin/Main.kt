import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream

fun main(args : Array<String>) {
    var max = 0
    var maxSheet = 0
    var maxRow = 0
    var maxCell = 0

    var path = "src\\main\\resources\\Mappe1.xlsx"
    if (args.isNotEmpty()) {
        path = args[0]
    }
    val inputStream = FileInputStream(path)
    val workbook = XSSFWorkbook(inputStream)
    for(sheetIndex in 0 until workbook.numberOfSheets) {
        val sheet = workbook.getSheetAt(sheetIndex)
        if (sheet != null) {
            for (rowIndex in 0 until sheet.lastRowNum + 1) {
                val row = sheet.getRow(rowIndex)
                if (row != null) {
                    for (cellIndex in 0 until row.lastCellNum + 1) {
                        val cell = row.getCell(cellIndex)
                        if(cell != null) {
                            try {
                                val value = cell.stringCellValue
                                if (value != null) {
                                    val length = value.length
                                    println("sheet: $sheetIndex, row: $rowIndex, cell: $cellIndex, length: $length")
                                    if (length > max) {
                                        max = length
                                        maxSheet = sheetIndex
                                        maxRow = rowIndex
                                        maxCell = cellIndex

                                    }
                                }
                            } catch (e: IllegalStateException) {
                                println("sheet: $sheetIndex, row: $rowIndex, cell: $cellIndex, ERROR: e.message")
                            }
                        }
                    }
                }
            }
        }
    }
    println("longest cell is $max characters long at sheet: $maxSheet, row: $maxRow, cell: $maxCell")
}