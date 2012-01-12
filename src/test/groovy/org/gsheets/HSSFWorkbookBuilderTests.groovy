package org.gsheets

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font

/**
 * @author me@andresteingress.com
 */
class HSSFWorkbookBuilderTests extends GroovyTestCase {

    File excel

    void setUp() {
        excel = new File("/Users/andre/Development/Projects/gsheets/temp/test.xls")
        if (!excel.exists()) excel.createNewFile()
    }

    void testCreateSimpleWorkbook()  {
        Workbook workbook = new HSSFWorkbookBuilder().workbook {

            // style definitions
            font("bold")  { Font font ->
                font.setBoldweight(Font.BOLDWEIGHT_BOLD)
            }

            cellStyle ("header")  { CellStyle cellStyle ->
                cellStyle.setAlignment(CellStyle.ALIGN_CENTER)
            }

            // data
            sheet ("Export")  {
                header(["Column1", "Column2", "Column3"])

                row(["a", "b", "c"])
            }

            // apply styles
            applyCellStyle(cellStyle: "header", font: "bold", rows: 1, columns: 1..3)
            mergeCells(rows: 1, columns: 1..3)
        }

        def excelOut = new FileOutputStream(excel)
        workbook.write(excelOut)
        excelOut.close()
    }
}
