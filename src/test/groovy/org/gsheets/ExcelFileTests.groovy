/**
 Copyright (c) 2012, Andre Steingress
 All rights reserved.

 Redistribution and use in source and binary forms, with or without
 modification, are permitted provided that the following conditions are met:
 1. Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
 2. Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
 3. All advertising materials mentioning features or use of this software
    must display the following acknowledgement:
    This product includes software developed by the ASF.
 4. Neither the name of the ASF nor the
    names of its contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

 THIS SOFTWARE IS PROVIDED BY Andre Steingress ''AS IS'' AND ANY
 EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 DISCLAIMED. IN NO EVENT SHALL Andre Steingress BE LIABLE FOR ANY
 DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 GSheets is a Groovy builder based on Apache POI.
 */
package org.gsheets

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook

/**
 * @author me@andresteingress.com
 */
class ExcelFileTests extends GroovyTestCase {

    File excel

    void setUp() {
        excel = new File("test.xls")
        if (!excel.exists()) excel.createNewFile()
    }

    void testCreateSimpleWorkbook()  {
        Workbook workbook = new ExcelFile().workbook {

            styles {
                font("bold")  { Font font ->
                    font.setBoldweight(Font.BOLDWEIGHT_BOLD)
                }

                cellStyle ("header")  { CellStyle cellStyle ->
                    cellStyle.setAlignment(CellStyle.ALIGN_CENTER)

                }
            }

            data {
                // data
                sheet ("Export")  {
                    header(["Column1", "Column2", "Column3"])

                    row([10, 20, "=A2*B2"])
                    row(["", "", "=sum(C2)"])
                }
            }

            commands {
                applyCellStyle(cellStyle: "header", font: "bold", rows: 1, columns: 1..3)
                applyColumnWidth(columns: 1..2, width: 200)
                // mergeCells(rows: 1, columns: 1..3)
            }
        }

        def excelOut = new FileOutputStream(excel)
        workbook.write(excelOut)
        excelOut.close()
    }

    void testOpenWorkbookAndEditExistingSheet() {
        def file = createTestFile()
        def workbook = new ExcelFile(file).workbook {
            data {
                sheet("SheetA") {
                    onRow(4) {
                        row (['new row', 'on existing', 'workbook'])
                    }
                }
            }
        }

        def out = new FileOutputStream(file)
        workbook.write(out)
        out.close()
    }

    void testOpenWorkbookAndAddNewSheet() {
        def file = createTestFile()
        def workbook = new ExcelFile(file).workbook {
            data {
                sheet("SheetB") {
                    header (["New", "Header"])
                    row (["new", "row"])
                }
            }
        }

        def out = new FileOutputStream(file)
        workbook.write(out)
        out.close()
    }

    private File createTestFile() {
        def file = File.createTempFile("gsheets", ".xls")
        def workbook = new ExcelFile().workbook {
            data {
                sheet ("SheetA") {
                    header (['HA', 'HB', 'HC'])

                    3.times {
                        row ([it, 2*it, 3*it])
                    }
                }
            }
        }

        def out = new FileOutputStream(file)
        workbook.write(out)
        out.close()

        file
    }
}
