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

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.usermodel.DataFormat

/**
 * Groovy builder used to create XLS files based on Apache POI HSSF.
 *
 * <pre>
 *
 * Workbook workbook = new HSSFWorkbookBuilder().workbook {
 *
 * // style definitions
 * font("bold")  { Font font ->
 *     font.setBoldweight(Font.BOLDWEIGHT_BOLD)
 * }
 *
 * cellStyle ("header")  { CellStyle cellStyle ->
 *     cellStyle.setAlignment(CellStyle.ALIGN_CENTER)
 * }
 *
 * // data
 * sheet ("Export")  {
 *     header(["Column1", "Column2", "Column3"])
 *
 *     row(["a", "b", "c"])
 * }
 *
 * // apply styles
 * applyCellStyle(cellStyle: "header", font: "bold", rows: 1, columns: 1..3)
 * mergeCells(rows: 1, columns: 1..3)
 * }
 *
 * </pre>
 *
 * @author me@andresteingress.com
 */
class HSSFWorkbookBuilder extends NodeBuilder {

    private Workbook workbook = new HSSFWorkbook()
    private Sheet sheet
    private int rowsCounter

    private Map<String, CellStyle> cellStyles = [:]
    private Map<String, Font> fonts = [:]

    Workbook workbook(Closure closure) {
        assert closure

        closure.delegate = this
        closure.call()
        workbook
    }

    void styles(Closure closure) {
        assert closure

        closure.delegate = this
        closure.call()
    }

    void data(Closure closure) {
        assert closure

        closure.delegate = this
        closure.call()
    }

    void commands(Closure closure) {
        assert closure

        closure.delegate = this
        closure.call()
    }

    void sheet(String name, Closure closure) {
        assert name
        assert closure

        sheet = workbook.createSheet(name)
        rowsCounter = 0
        closure.delegate = this
        closure.call()
    }

    void cellStyle(String cellStyleId, Closure closure)  {
        assert cellStyleId
        assert !cellStyles.containsKey(cellStyleId)
        assert closure

        CellStyle cellStyle = workbook.createCellStyle()
        cellStyles.put(cellStyleId, cellStyle)

        closure.call(cellStyle)
    }

    void font(String fontId, Closure closure)  {
        assert fontId
        assert !fonts.containsKey(fontId)
        assert closure

        Font font = workbook.createFont()
        fonts.put(fontId, font)

        closure.call(font)
    }

    void applyCellStyle(Map<String, Object> args)  {
        def cellStyleId = args.cellStyle
        def fontId = args.font
        def dataFormat = args.dataFormat

        def sheetName = args.sheet

        def rows = args.rows
        def cells = args.columns

        assert cellStyleId || fontId || dataFormat

        assert rows && (rows instanceof Number || rows instanceof Range<Number>)
        assert cells && (cells instanceof Number || cells instanceof Range<Number>)

        if (cellStyleId && !cellStyles.containsKey(cellStyleId)) cellStyleId = null
        if (fontId && !fonts.containsKey(fontId)) fontId = null
        if (dataFormat && !(dataFormat instanceof String)) dataFormat = null
        if (sheetName && !(sheetName instanceof String)) sheetName = null

        if (rows instanceof  Number) rows = [rows]
        if (cells instanceof  Number) cells = [cells]

        rows.each { Number rowIndex ->
            assert rowIndex

            cells.each { Number cellIndex ->
                assert cellIndex

                def sheet = sheetName ? workbook.getSheet(sheetName as String) : workbook.getSheetAt(0)
                assert sheet

                Row row = sheet.getRow(rowIndex.intValue() - 1)
                if (!row) return

                Cell cell = row.getCell(cellIndex.intValue() - 1)
                if (!cell) return

                if (cellStyleId) cell.setCellStyle(cellStyles.get(cellStyleId))
                if (fontId) cell.getCellStyle().setFont(fonts.get(fontId))
                if (dataFormat) {
                    DataFormat df = workbook.createDataFormat()
                    cell.getCellStyle().setDataFormat(df.getFormat(dataFormat as String))
                }
            }
        }
    }

    void mergeCells(Map<String, Object> args)  {
        def rows = args.rows
        def cols = args.columns
        def sheetName = args.sheet

        assert rows && (rows instanceof Number || rows instanceof Range<Number>)
        assert cols && (cols instanceof Number || cols instanceof Range<Number>)

        if (rows instanceof Number) rows = [rows]
        if (cols instanceof Number) cols = [cols]
        if (sheetName && !(sheetName instanceof String)) sheetName = null

        def sheet = sheetName ? workbook.getSheet(sheetName as String) : workbook.getSheetAt(0)
        
        sheet.addMergedRegion(new CellRangeAddress(rows.first() - 1, rows.last() - 1, cols.first() - 1, cols.last() - 1))
    }

    void header(List<String> names)  {
        assert names

        Row row = sheet.createRow(rowsCounter++ as int)
        names.eachWithIndex { String value, col ->
            Cell cell = row.createCell(col)
            cell.setCellValue(value)
        }
    }

    void row(values) {
        assert values

        Row row = sheet.createRow(rowsCounter++ as int)
        values.eachWithIndex {value, col ->
            Cell cell = row.createCell(col)
            switch (value) {
                case Date: cell.setCellValue((Date) value); break
                case Double: cell.setCellValue((Double) value); break
                case BigDecimal: cell.setCellValue(((BigDecimal) value).doubleValue()); break
                default: cell.setCellValue(new HSSFRichTextString("" + value)); break
            }
        }
    }
}
