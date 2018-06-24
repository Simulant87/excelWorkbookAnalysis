
data class SheetResult(
        val cell: CellIdentification,
        val maxStringLength: Int,
        val sheetSize: SheetSize
)


data class CellIdentification(
        val sheetIndex: Int,
        val rowIndex: Int,
        val columnIndex: Int
)

data class SheetSize(
        val maxRowIndex: Int,
        val maxColumnIndex: Int
)

data class WorkbookResult(
        val maxStringLength: Int,
        val longestStringCell: CellIdentification
//        val longestRowCell: CellIdentification,
//        val longestColumn: CellIdentification
)
