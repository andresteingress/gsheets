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
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.HSSFCell

/**
 * A Groovy builder that wraps Apache POI for generating binary Microsoft Excel sheets.
 *
 * <pre>
 *
 * Workbook workbook = new ExcelFile().workbook {
 *
 * // define cell styles and fonts
 * styles {
 *   font("bold")  { Font font ->
 *       font.setBoldweight(Font.BOLDWEIGHT_BOLD)
 *   }
 *
 *   cellStyle ("header")  { CellStyle cellStyle ->
 *       cellStyle.setAlignment(CellStyle.ALIGN_CENTER)
 *   }
 * }
 *
 * // declare the data to use
 * data {
 *   sheet ("Export")  {
 *       header(["Column1", "Column2", "Column3"])
 *
 *       row(["a", "b", "c"])
 *   }
 * }
 *
 * // apply link styles with data through 'commands'
 * commands {
 *     applyCellStyle(cellStyle: "header", font: "bold", rows: 1, columns: 1..3)
 *     mergeCells(rows: 1, columns: 1..3)
 *
 *     // autosize columns. useMergedCells is optional, defaults to false.
 *     applyAutoSizeColumn(columns: 1..3, useMergedCells: true)
 * }
 * }
 *
 * </pre>
 *
 * @author me@andresteingress.com
 */
class ExcelFile {

    Workbook workbook = new HSSFWorkbook()
    private Sheet sheet
    private int rowsCounter

    private Map<String, CellStyle> cellStyles = [:]
    private Map<String, Font> fonts = [:]

    /**
     * Creates a new workbook.
     *
     * @param the closure holds nested {@link ExcelFile} method calls
     * @return the created {@link Workbook}
     */
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
        assert workbook

        assert name
        assert closure

        sheet = workbook.createSheet(name)
        rowsCounter = 0

        closure.delegate = sheet
        closure.call()
    }

    void cellStyle(String cellStyleId, Closure closure)  {
        assert workbook

        assert cellStyleId
        assert !cellStyles.containsKey(cellStyleId)
        assert closure

        CellStyle cellStyle = workbook.createCellStyle()
        cellStyles.put(cellStyleId, cellStyle)

        closure.call(cellStyle)
    }

    void font(String fontId, Closure closure)  {
        assert workbook

        assert fontId
        assert !fonts.containsKey(fontId)
        assert closure

        Font font = workbook.createFont()
        fonts.put(fontId, font)

        closure.call(font)
    }

    void applyCellStyle(Map<String, Object> args)  {
        assert workbook

        def cellStyleId = args.cellStyle
        def fontId = args.font
        def dataFormat = args.dataFormat

        def sheetName = args.sheet

        def rows = args.rows ?: -1          // -1 denotes all rows
        def cells = args.columns ?: -1      // -1 denotes all cols

        def colName = args.columnName

        assert cellStyleId || fontId || dataFormat

        assert rows && (rows instanceof Number || rows instanceof Range<Number>)
        assert cells && (cells instanceof Number || cells instanceof Range<Number>)

        if (cellStyleId && !cellStyles.containsKey(cellStyleId)) cellStyleId = null
        if (fontId && !fonts.containsKey(fontId)) fontId = null
        if (dataFormat && !(dataFormat instanceof String)) dataFormat = null
        if (sheetName && !(sheetName instanceof String)) sheetName = null
        if (colName && !(colName instanceof String)) colName = null

        def sheet = sheetName ? workbook.getSheet(sheetName as String) : workbook.getSheetAt(0)
        assert sheet

        if (rows == -1)  rows  = [1..rowsCounter]
        if (rows instanceof  Number) rows = [rows]

        rows.each { Number rowIndex ->
            assert rowIndex

            Row row = sheet.getRow(rowIndex.intValue() - 1)
            if (!row) return

            if (cells == -1)  cells  = [row.firstCellNum..row.lastCellNum]
            if (rows instanceof  Number) rows = [rows]

            def applyStyleFunc = { Number cellIndex ->
                assert cellIndex

                Cell cell = row.getCell(cellIndex.intValue() - 1)
                if (!cell) return

                if (cellStyleId) cell.setCellStyle(cellStyles.get(cellStyleId))
                if (fontId) cell.getCellStyle().setFont(fonts.get(fontId))
                if (dataFormat) {
                    DataFormat df = workbook.createDataFormat()
                    cell.getCellStyle().setDataFormat(df.getFormat(dataFormat as String))
                }
            }

            cells.each applyStyleFunc
        }
    }

    void mergeCells(Map<String, Object> args)  {
        assert workbook

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

    void applyColumnWidth(Map<String, Object> args)  {
        assert workbook

        def cols = args.columns
        def sheetName = args.sheet
        def width = args.width

        assert cols && (cols instanceof Number || cols instanceof Range<Number>)
        assert width && width instanceof Number

        if (cols instanceof Number) cols = [cols]
        if (sheetName && !(sheetName instanceof String)) sheetName = null

        def sheet = sheetName ? workbook.getSheet(sheetName as String) : workbook.getSheetAt(0)

        cols.each {
            sheet.setColumnWidth(it - 1, width.intValue())
        }
    }

	void applyAutoSizeColumn(Map<String, Object> args) {
        assert sheet

        def cols = args.columns
        def useMergedCells = args.useMergedCells ?: false

        assert cols && (cols instanceof Number || cols instanceof Range<Number>)
        assert useMergedCells instanceof Boolean

        cols.each {
            sheet.autoSizeColumn(it - 1, useMergedCells);
        }
    }

    void header(List<String> names)  {
        assert sheet
        assert names

        Row row = sheet.createRow(rowsCounter++ as int)
        names.eachWithIndex { String value, col ->
            Cell cell = row.createCell(col)
            cell.setCellValue(value)
        }
    }

    void emptyRow()  {
        assert sheet

        sheet.createRow(rowsCounter++ as int)
    }

    void row(values) {
        assert sheet
        assert values

        Row row = sheet.createRow(rowsCounter++ as int)
        values.eachWithIndex {value, col ->
            Cell cell = row.createCell(col)
            switch (value) {
                case Date: cell.setCellValue((Date) value); break
                case Double: cell.setCellValue((Double) value); break
                case BigDecimal: cell.setCellValue(((BigDecimal) value).doubleValue()); break
                case Number: cell.setCellValue(((Number) value).doubleValue()); break
                default:
                    def stringValue = value?.toString() ?: ""
                    if (stringValue.startsWith('=')) {
                        cell.setCellType(Cell.CELL_TYPE_FORMULA)
                        cell.setCellFormula(stringValue.substring(1))
                    } else {
                        cell.setCellValue(new HSSFRichTextString(stringValue))
                    }
                    break
            }
        }
    }

    int getRowCount()  {
        assert sheet

        rowsCounter
    }
}
