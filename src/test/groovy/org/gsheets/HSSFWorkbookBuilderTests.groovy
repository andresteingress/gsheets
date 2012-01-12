package org.gsheets

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.CellStyle

/**
 * @author me@andresteingress.com
 */
class HSSFWorkbookBuilderTests extends GroovyTestCase {

    File excel

    void setUp() {
        excel = new File("/Users/andre/Development/Projects/Adternity/temp/test.xls")
        if (!excel.exists()) excel.createNewFile()
    }

    void testCreateSimpleWorkbook()  {
        Workbook workbook = new HSSFWorkbookBuilder().workbook {
            
            cellStyle ("header")  { CellStyle cellStyle ->
                cellStyle.setAlignment(CellStyle.ALIGN_CENTER)
            }
            
            sheet ("Export")  {
                header(["Column1", "Column2", "Column3"])

                row(["a", "b", "c"])
            }
            
            applyCellStyle(id: "header", rows: 1, columns: 1..3)
        }

        def excelOut = new FileOutputStream(excel)
        workbook.write(excelOut)
        excelOut.close()
    }
}
