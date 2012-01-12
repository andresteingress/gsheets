package org.gsheets

import org.apache.poi.ss.usermodel.Workbook

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
            sheet ("Export")  {
                row(["a", "b", "c"])
            }
        }

        def excelOut = new FileOutputStream(excel)
        workbook.write(excelOut)
        excelOut.close()
    }
}
