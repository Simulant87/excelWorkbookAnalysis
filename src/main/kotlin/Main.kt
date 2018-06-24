import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream

/**
 * reads a path to a file from the first command line argument, tries to read the file as excel workbook
 * and reports the longest String in a cell, highest row and column count per sheet.
 */
fun main(args: Array<String>) {

    val workbook = createWorkbook(args)

    val numberOfSheets = workbook.numberOfSheets
    val sheetResults = kotlin.arrayOfNulls<SheetResult>(numberOfSheets)
    val threads = ArrayList<Thread>()

    for(sheetIndex in 0 until numberOfSheets) {
        val sheet = workbook.getSheetAt(sheetIndex)
        if (sheet != null) {
            //starting a thread for each sheet to analyze the workbook in parallel
            val thread = Thread { analyzeSheet(sheet, sheetIndex, sheetResults) }
            thread.start()
            threads.add(thread)
        }
    }

    //joining all threads started for the analysis to the report is not empty
    for( thread in threads) {
        thread.join()
    }

    println()
    val result = analyzeSheetResults(sheetResults)
    println()
    println("Over all longest String result: $result")
}

private fun createWorkbook(args: Array<String>): Workbook {
    val path = if (args.isNotEmpty()) {
        //reading path to workbook to be analyzed from first command line argument
        args[0]
    } else {
        //default fall back path to the workbook in the resources directory
        """src\main\resources\Mappe1.xlsx"""
    }
    println("path to analyze: $path")
    val inputStream = FileInputStream(path)
    return if (path.endsWith(".xls")) {
        HSSFWorkbook(inputStream)
    } else {
        XSSFWorkbook(inputStream)
    }
}

private fun analyzeSheet(sheet: Sheet, sheetIndex: Int, sheetResults: Array<SheetResult?>) {
    var maxStringLength = 0
    var maxRow = 0
    var maxCell = 0
    var maxColumnCount = 0
    for (rowIndex in 0 until sheet.lastRowNum + 1) {
        val row = sheet.getRow(rowIndex)
        if (row != null) {
            if (maxColumnCount < row.lastCellNum) {
                maxColumnCount = row.lastCellNum.toInt() - 1
            }
            for (cellIndex in 0 until row.lastCellNum) {
                val cell = row.getCell(cellIndex)
                if (cell != null) {
                    try {
                        val value = cell.stringCellValue
                        if (value != null) {
                            val length = value.length
                            println("sheet: $sheetIndex, row: $rowIndex, cell: $cellIndex, length: $length")
                            if (length > maxStringLength) {
                                maxStringLength = length
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
    val cell = CellIdentification(sheetIndex, maxRow, maxCell)
    val sheetSize = SheetSize(sheet.lastRowNum, maxColumnCount)
    sheetResults[sheetIndex] = SheetResult(cell, maxStringLength, sheetSize)
}

fun analyzeSheetResults(sheetResults: Array<SheetResult?>): WorkbookResult {
    var maxStringLength = 0
    var maxStringLengthSheetIndex = 0
    for ((index, sheetResult) in sheetResults.withIndex()) {
        println(sheetResult)
        if (sheetResult != null) {
            if (sheetResult.maxStringLength > maxStringLength) {
                maxStringLength = sheetResult.maxStringLength
                maxStringLengthSheetIndex = index
            }
        }
    }
    return WorkbookResult(maxStringLength, sheetResults[maxStringLengthSheetIndex]!!.cell)
}
