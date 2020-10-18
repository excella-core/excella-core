/*-
 * #%L
 * excella-core
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.core.util;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.ParseException;
import java.util.Date;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.CellClone;
import org.bbreak.excella.core.CoreTestUtil;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.test.util.CheckException;
import org.bbreak.excella.core.test.util.TestUtil;
import org.junit.Test;

/**
 * PoiUtilテストクラス
 * 
 * @since 1.0
 */
public class PoiUtilTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public PoiUtilTest( String version) {
        super( version);
    }

    @Test
    public void testPoiUtil1() throws IOException, ParseException {

        Workbook workbook = getWorkbook();
        Sheet sheet_1 = workbook.getSheetAt( 0);

        Date expectedDate = DateFormat.getDateInstance().parse( "2009/4/16");
        String expectedString = "あああ";

        // ===============================================
        // isCellDateFormatted( Cell cell)
        // ===============================================
        assertEquals( Boolean.FALSE, PoiUtil.isCellDateFormatted( null));

        // ===============================================
        // getJavaDate( double excelDate)
        // ===============================================
        double excelDate = 39919; // 2009/4/16 --> 39919
        Date javaDate = PoiUtil.getJavaDate( excelDate);
        assertEquals( expectedDate, javaDate);

        // ===============================================
        // getCellValue(Cell cell)
        // ===============================================
        Object object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0));
        assertEquals( new Double( 10.0), object);

        // ===============================================
        // getCellValue(Cell cell, Class<?> propertyClass)
        // ===============================================
        // --> Short
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Short.class);
        assertEquals( Short.class, object.getClass());

        // --> Integer
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Integer.class);
        assertEquals( Integer.class, object.getClass());

        // --> Long
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Long.class);
        assertEquals( Long.class, object.getClass());

        // --> Float
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Float.class);
        assertEquals( Float.class, object.getClass());

        // --> Double
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Double.class);
        assertEquals( Double.class, object.getClass());

        // --> BigDecimal
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), BigDecimal.class);
        assertEquals( BigDecimal.class, object.getClass());

        // --> Byte
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Byte.class);
        assertEquals( Byte.class, object.getClass());

        // --> Date
        object = PoiUtil.getCellValue( sheet_1.getRow( 5).getCell( 0), Date.class);
        assertEquals( Date.class, object.getClass());

        // --> String
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), String.class);
        assertEquals( String.class, object.getClass());

        // --> Boolean
        object = PoiUtil.getCellValue( sheet_1.getRow( 3).getCell( 0), Boolean.class);
        assertEquals( Boolean.class, object.getClass());

        // --> boolean
        object = PoiUtil.getCellValue( sheet_1.getRow( 3).getCell( 0), boolean.class);
        assertEquals( Boolean.class, object.getClass());

        // --> byte
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), byte.class);
        assertEquals( Byte.class, object.getClass());

        // --> short
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), short.class);
        assertEquals( Short.class, object.getClass());

        // --> int
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), int.class);
        assertEquals( Integer.class, object.getClass());

        // --> long
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), long.class);
        assertEquals( Long.class, object.getClass());

        // --> float
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), float.class);
        assertEquals( Float.class, object.getClass());

        // --> double
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), double.class);
        assertEquals( Double.class, object.getClass());

        // --> other
        object = PoiUtil.getCellValue( sheet_1.getRow( 0).getCell( 0), Object.class);
        assertEquals( null, object);

        // ===============================================
        // getSheetName(Sheet sheet)
        // ===============================================
        String sheetname = PoiUtil.getSheetName( sheet_1.getRow( 0).getCell( 0));
        assertEquals( "Sheet1", sheetname);

        // ===============================================
        // getSheetName(Sheet sheet)
        // ===============================================
        sheetname = PoiUtil.getSheetName( workbook.getSheetAt( 0));
        assertEquals( "Sheet1", sheetname);

        // ===============================================
        // getCellValue(Sheet sheet, int rowIndex, int columnIndex)
        // ===============================================
        // CELL_TYPE_NUMERIC
        Object cellValue = PoiUtil.getCellValue( sheet_1, 0, 0);
        assertEquals( 10.0, cellValue);

        // CELL_TYPE_FORMULA
        cellValue = PoiUtil.getCellValue( sheet_1, 1, 0);
        assertEquals( 10000.0, cellValue);

        // CELL_TYPE_STRING
        cellValue = PoiUtil.getCellValue( sheet_1, 2, 0);
        assertEquals( expectedString, cellValue);

        // CELL_TYPE_BOOLEAN
        cellValue = PoiUtil.getCellValue( sheet_1, 3, 0);
        assertEquals( Boolean.TRUE, cellValue);

        // CELL_TYPE_ERROR
        cellValue = PoiUtil.getCellValue( sheet_1, 4, 0);
        if ( workbook instanceof HSSFWorkbook) {
            // #N/Aのエラーコード
            assertEquals( new Byte( "42"), cellValue);
        } else if ( workbook instanceof XSSFWorkbook) {
            // XSSF形式の場合はエラーの文字列を返却
            assertEquals( "#N/A", cellValue);
        }

        // CELL_TYPE_NUMERIC -> Date
        cellValue = PoiUtil.getCellValue( sheet_1, 5, 0);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_BLANK
        cellValue = PoiUtil.getCellValue( sheet_1, 6, 0);
        assertEquals( null, cellValue);

        // CELL_TYPE_FORMULA -> CELL_TYPE_NUMERIC
        cellValue = PoiUtil.getCellValue( sheet_1, 7, 0);
        assertEquals( 10.0, cellValue);

        // CELL_TYPE_FORMULA -> CELL_TYPE_STRING
        cellValue = PoiUtil.getCellValue( sheet_1, 8, 0);
        assertEquals( expectedString, cellValue);

        // CELL_TYPE_FORMULA -> CELL_TYPE_BOOLEAN
        cellValue = PoiUtil.getCellValue( sheet_1, 9, 0);
        assertEquals( Boolean.TRUE, cellValue);

        // CELL_TYPE_FORMULA -> CELL_TYPE_ERROR
        cellValue = PoiUtil.getCellValue( sheet_1, 10, 0);
        // #N/Aを返却
        if ( workbook instanceof HSSFWorkbook) {
            // #N/Aのエラーコード
            assertEquals( new Byte( "42"), cellValue);
        } else if ( workbook instanceof XSSFWorkbook) {
            // XSSF形式の場合はエラーの文字列を返却
            assertEquals( "#N/A", cellValue);
        }

        // CELL_TYPE_FORMULA -> Date
        cellValue = PoiUtil.getCellValue( sheet_1, 11, 0);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_FORMULA -> CELL_TYPE_BLANK
        cellValue = PoiUtil.getCellValue( sheet_1, 12, 0);
        assertEquals( new Double( 0.0), cellValue);

        // ===============================================
        // crossRangeAddress( CellRangeAddress baseAddress, CellRangeAddress targetAddress)
        // ===============================================
        // 基準範囲
        CellRangeAddress baseAddress = new CellRangeAddress( 2, 3, 2, 3);

        // ---重なっている場合（基準範囲と対象範囲が完全一致）---
        CellRangeAddress targetAddress1 = new CellRangeAddress( 2, 3, 2, 3);
        assertTrue( PoiUtil.crossRangeAddress( baseAddress, targetAddress1));

        // ---重なっている場合（基準範囲と対象範囲が部分一致）---
        // 対象範囲が上側にはみ出している場合
        CellRangeAddress targetAddress2 = new CellRangeAddress( 1, 2, 2, 3);
        assertTrue( PoiUtil.crossRangeAddress( baseAddress, targetAddress2));
        // 対象範囲が下側にはみ出している場合
        CellRangeAddress targetAddress3 = new CellRangeAddress( 3, 4, 2, 3);
        assertTrue( PoiUtil.crossRangeAddress( baseAddress, targetAddress3));
        // 対象範囲が左側にはみ出している場合
        CellRangeAddress targetAddress4 = new CellRangeAddress( 2, 3, 1, 2);
        assertTrue( PoiUtil.crossRangeAddress( baseAddress, targetAddress4));
        // 対象範囲が右側にはみ出している場合
        CellRangeAddress targetAddress5 = new CellRangeAddress( 2, 3, 3, 4);
        assertTrue( PoiUtil.crossRangeAddress( baseAddress, targetAddress5));

        // ---重なっていない場合（基準範囲と対象範囲の領域が別々に存在）---
        // 対象範囲が上側に存在する場合
        CellRangeAddress targetAddress6 = new CellRangeAddress( 0, 1, 2, 3);
        assertFalse( PoiUtil.crossRangeAddress( baseAddress, targetAddress6));
        // 対象範囲が下側に存在する場合
        CellRangeAddress targetAddress7 = new CellRangeAddress( 4, 5, 2, 3);
        assertFalse( PoiUtil.crossRangeAddress( baseAddress, targetAddress7));
        // 対象範囲が左側に存在する場合
        CellRangeAddress targetAddress8 = new CellRangeAddress( 2, 3, 0, 1);
        assertFalse( PoiUtil.crossRangeAddress( baseAddress, targetAddress8));
        // 対象範囲が右側に存在する場合
        CellRangeAddress targetAddress9 = new CellRangeAddress( 2, 3, 4, 5);
        assertFalse( PoiUtil.crossRangeAddress( baseAddress, targetAddress9));

        // ===============================================
        // containCellRangeAddress( CellRangeAddress baseAddress, CellRangeAddress targetAddress)
        // ===============================================
        CellRangeAddress rangeAddress1 = new CellRangeAddress( 0, 2, 0, 2);
        CellRangeAddress rangeAddress2 = new CellRangeAddress( 3, 4, 3, 4);
        CellRangeAddress rangeAddress3 = new CellRangeAddress( 0, 1, 0, 1);
        // 含まれていない
        assertFalse( PoiUtil.containCellRangeAddress( rangeAddress1, rangeAddress2));
        // 含まれてている
        assertTrue( PoiUtil.containCellRangeAddress( rangeAddress1, rangeAddress3));

        // ===============================================
        // writeBook( Workbook workbook, String filename)
        // ===============================================
        String extension = BookController.HSSF_SUFFIX;
        if ( workbook instanceof XSSFWorkbook) {
            extension = BookController.XSSF_SUFFIX;
        }
        PoiUtil.writeBook( workbook, CoreTestUtil.getTestOutputDir() + "PoiUtilTest" + System.currentTimeMillis() + extension);
    }

    @Test
    public void testPoiUtil2() throws ParseException, IOException {

        Workbook workbook = getWorkbook();
        Sheet sheet_1 = workbook.getSheetAt( 0);

        Date expectedDate = DateFormat.getDateInstance().parse( "2009/4/16");

        // ===============================================
        // 日付取得の調査用
        // ===============================================
        // CELL_TYPE_NUMERIC-Date:*yyyy年MM月dd日
        Object cellValue = PoiUtil.getCellValue( sheet_1, 0, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:yyyy年MM月dd日
        cellValue = PoiUtil.getCellValue( sheet_1, 1, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:yyyy年MM月
        cellValue = PoiUtil.getCellValue( sheet_1, 2, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:MM月dd日
        cellValue = PoiUtil.getCellValue( sheet_1, 3, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:yyyy/MM/dd
        cellValue = PoiUtil.getCellValue( sheet_1, 4, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:yyyy/MM/dd 12:00 AM
        cellValue = PoiUtil.getCellValue( sheet_1, 5, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:yyyy/MM/dd 0:00
        cellValue = PoiUtil.getCellValue( sheet_1, 6, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:MM/dd
        cellValue = PoiUtil.getCellValue( sheet_1, 7, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:MM/dd/yy（月日０埋め無）
        cellValue = PoiUtil.getCellValue( sheet_1, 8, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:MM/dd/yy（月日０埋め有）
        cellValue = PoiUtil.getCellValue( sheet_1, 9, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:dd-month
        cellValue = PoiUtil.getCellValue( sheet_1, 10, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:dd-month-yy
        cellValue = PoiUtil.getCellValue( sheet_1, 11, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:dd-month-yy
        cellValue = PoiUtil.getCellValue( sheet_1, 12, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:month-yy
        cellValue = PoiUtil.getCellValue( sheet_1, 13, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:month-yy
        cellValue = PoiUtil.getCellValue( sheet_1, 14, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:month
        cellValue = PoiUtil.getCellValue( sheet_1, 15, 5);
        assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:month-yy
        cellValue = PoiUtil.getCellValue( sheet_1, 16, 5);
        assertEquals( expectedDate, cellValue);

        // TODO : Localizeされたフォーマットに対応していない
        // CELL_TYPE_NUMERIC-Date:Hyy.MM.dd
        cellValue = PoiUtil.getCellValue( sheet_1, 17, 5);
        // assertEquals( expectedDate, cellValue);

        // CELL_TYPE_NUMERIC-Date:平成yy年MM月dd日
        cellValue = PoiUtil.getCellValue( sheet_1, 18, 5);
        // assertEquals( expectedDate, cellValue);
    }

    @Test
    public void testPoiUtil3() throws IOException, ParseException, CheckException {

        Workbook workbook = getWorkbook();
        Sheet sheet_1 = workbook.getSheetAt( 0);
        Sheet sheet_2 = workbook.getSheetAt( 1);
        Sheet sheet_3 = workbook.getSheetAt( 2);
        Sheet sheet_4 = workbook.getSheetAt( 3);
        Sheet sheet_5 = workbook.getSheetAt( 4);
        Sheet sheet_6 = workbook.getSheetAt( 5);
        Sheet sheet_7 = workbook.getSheetAt( 6);

        // ===============================================
        // copyCell( Cell fromCell, Cell toCell)
        // ===============================================
        // No.1 各セルタイプコピー
        Cell fromCellNumeric = sheet_1.getRow( 0).getCell( 0);
        Cell fromCellFormula = sheet_1.getRow( 1).getCell( 0);
        Cell fromCellString = sheet_1.getRow( 2).getCell( 0);
        Cell fromCellBoolean = sheet_1.getRow( 3).getCell( 0);
        Cell fromCellError = sheet_1.getRow( 4).getCell( 0);
        Cell fromCellDate = sheet_1.getRow( 5).getCell( 0);
        Cell fromCellBlank = sheet_1.getRow( 6).getCell( 0);

        Cell toCellNumeric = sheet_1.getRow( 0).createCell( 9);
        Cell toCellFormula = sheet_1.getRow( 1).createCell( 9);
        Cell toCellString = sheet_1.getRow( 2).createCell( 9);
        Cell toCellBoolean = sheet_1.getRow( 3).createCell( 9);
        Cell toCellError = sheet_1.getRow( 4).createCell( 9);
        Cell toCellDate = sheet_1.getRow( 5).createCell( 9);
        Cell toCellBlank = sheet_1.getRow( 6).createCell( 9);

        Cell fromCellNumericFrml = sheet_1.getRow( 7).getCell( 0);
        Cell fromCellStringFrml = sheet_1.getRow( 8).getCell( 0);
        Cell fromCellBooleanFrml = sheet_1.getRow( 9).getCell( 0);
        Cell fromCellErrorFrml = sheet_1.getRow( 10).getCell( 0);
        Cell fromCellDateFrml = sheet_1.getRow( 11).getCell( 0);
        Cell fromCellBlankFrml = sheet_1.getRow( 12).getCell( 0);

        Cell toCellNumericFrml = sheet_1.getRow( 7).createCell( 9);
        Cell toCellStringFrml = sheet_1.getRow( 8).createCell( 9);
        Cell toCellBooleanFrml = sheet_1.getRow( 9).createCell( 9);
        Cell toCellErrorFrml = sheet_1.getRow( 10).createCell( 9);
        Cell toCellDateFrml = sheet_1.getRow( 11).createCell( 9);
        Cell toCellBlankFrml = sheet_1.getRow( 12).createCell( 9);

        PoiUtil.copyCell( fromCellNumeric, toCellNumeric);
        PoiUtil.copyCell( fromCellFormula, toCellFormula);
        PoiUtil.copyCell( fromCellString, toCellString);
        PoiUtil.copyCell( fromCellBoolean, toCellBoolean);
        PoiUtil.copyCell( fromCellError, toCellError);
        PoiUtil.copyCell( fromCellDate, toCellDate);
        PoiUtil.copyCell( fromCellBlank, toCellBlank);

        PoiUtil.copyCell( fromCellNumericFrml, toCellNumericFrml);
        PoiUtil.copyCell( fromCellStringFrml, toCellStringFrml);
        PoiUtil.copyCell( fromCellBooleanFrml, toCellBooleanFrml);
        PoiUtil.copyCell( fromCellErrorFrml, toCellErrorFrml);
        PoiUtil.copyCell( fromCellDateFrml, toCellDateFrml);
        PoiUtil.copyCell( fromCellBlankFrml, toCellBlankFrml);

        // セルの検証
        TestUtil.checkCell( fromCellNumeric, toCellNumeric);
        TestUtil.checkCell( fromCellFormula, toCellFormula);
        TestUtil.checkCell( fromCellString, toCellString);
        TestUtil.checkCell( fromCellBoolean, toCellBoolean);
        TestUtil.checkCell( fromCellError, toCellError);
        TestUtil.checkCell( fromCellDate, toCellDate);
        TestUtil.checkCell( fromCellBlank, toCellBlank);

        TestUtil.checkCell( fromCellNumericFrml, toCellNumericFrml);
        TestUtil.checkCell( fromCellStringFrml, toCellStringFrml);
        TestUtil.checkCell( fromCellBooleanFrml, toCellBooleanFrml);
        TestUtil.checkCell( fromCellErrorFrml, toCellErrorFrml);
        TestUtil.checkCell( fromCellDateFrml, toCellDateFrml);
        TestUtil.checkCell( fromCellBlankFrml, toCellBlankFrml);

        // No.2 fromCellがnull
        Cell toCell = sheet_1.getRow( 0).createCell( 10);
        PoiUtil.copyCell( null, toCell);

        // No.3 toCellがnull
        try {
            PoiUtil.copyCell( fromCellNumeric, null);
            fail( "NullPointerException expected, but no exception thrown.");
        } catch ( NullPointerException ex) {
            // toCellがnullの場合は例外が発生
        }

        // No.4 コピー先が別シート
        Cell toCellNumeric2 = sheet_2.getRow( 0).createCell( 0);
        PoiUtil.copyCell( fromCellNumeric, toCellNumeric2);

        // ===============================================
        // copyRange( Sheet fromSheet, CellRangeAddress rangeAddress, Sheet toSheet, int toRowNum, int toColumnNum, boolean clearFromRange)
        // ===============================================
        // No.5 単一セル範囲コピー
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 0, 0, 0, 0), sheet_2, 0, 3, false);
        TestUtil.checkCell( sheet_1.getRow( 0).getCell( 0), sheet_2.getRow( 0).getCell( 3));

        // No.6 複数セル範囲コピー
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 1, 12, 0, 1), sheet_2, 9, 0, false);
        TestUtil.checkCell( sheet_1.getRow( 1).getCell( 0), sheet_2.getRow( 9).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 2).getCell( 0), sheet_2.getRow( 10).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 3).getCell( 0), sheet_2.getRow( 11).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 4).getCell( 0), sheet_2.getRow( 12).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 5).getCell( 0), sheet_2.getRow( 13).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 6).getCell( 0), sheet_2.getRow( 14).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 7).getCell( 0), sheet_2.getRow( 15).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 8).getCell( 0), sheet_2.getRow( 16).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 9).getCell( 0), sheet_2.getRow( 17).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 10).getCell( 0), sheet_2.getRow( 18).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 11).getCell( 0), sheet_2.getRow( 19).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 12).getCell( 0), sheet_2.getRow( 20).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 1).getCell( 1), sheet_2.getRow( 9).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 2).getCell( 1), sheet_2.getRow( 10).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 3).getCell( 1), sheet_2.getRow( 11).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 4).getCell( 1), sheet_2.getRow( 12).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 5).getCell( 1), sheet_2.getRow( 13).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 6).getCell( 1), sheet_2.getRow( 14).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 7).getCell( 1), sheet_2.getRow( 15).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 8).getCell( 1), sheet_2.getRow( 16).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 9).getCell( 1), sheet_2.getRow( 17).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 10).getCell( 1), sheet_2.getRow( 18).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 11).getCell( 1), sheet_2.getRow( 19).getCell( 1));
        TestUtil.checkCell( sheet_1.getRow( 12).getCell( 1), sheet_2.getRow( 20).getCell( 1));

        // No.7 引数にnullを設定
        PoiUtil.copyRange( null, new CellRangeAddress( 0, 0, 0, 0), sheet_2, 0, 0, false);
        PoiUtil.copyRange( sheet_1, null, sheet_2, 0, 0, false);
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 0, 0, 0, 0), null, 0, 0, false);

        // No.8 不正な範囲指定
        try {
            PoiUtil.copyRange( sheet_1, new CellRangeAddress( 1, 0, 0, 0), sheet_2, 0, 0, false);
            fail( "IllegalArgumentException expected, but no exception thrown.");
        } catch ( IllegalArgumentException ex) {
            // 不正な範囲を指定した場合、例外が発生
        }

        // No.9 結合セル範囲コピー
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 23, 23, 0, 1), sheet_2, 22, 0, false);
        TestUtil.checkCell( sheet_1.getRow( 23).getCell( 0), sheet_2.getRow( 22).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 23).getCell( 1), sheet_2.getRow( 22).getCell( 1));

        // No.10 結合セル範囲コピー
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 25, 26, 0, 0), sheet_2, 24, 0, false);
        TestUtil.checkCell( sheet_1.getRow( 25).getCell( 0), sheet_2.getRow( 24).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 26).getCell( 0), sheet_2.getRow( 25).getCell( 0));

        // No.11 nullセル範囲コピー
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 30, 30, 0, 1), sheet_2, 29, 0, false);
        TestUtil.checkCell( sheet_1.getRow( 30).getCell( 0), sheet_2.getRow( 29).getCell( 0));
        TestUtil.checkCell( sheet_1.getRow( 30).getCell( 1), sheet_2.getRow( 29).getCell( 1));

        // No.12 null行範囲コピー
        PoiUtil.copyRange( sheet_1, new CellRangeAddress( 34, 34, 0, 3), sheet_2, 33, 0, false);
        assertNull( sheet_2.getRow( 33));

        // No.13 コピー範囲重なる
        Cell copyFrom1 = new CellClone( sheet_2.getRow( 40).getCell( 0));
        Cell copyFrom2 = new CellClone( sheet_2.getRow( 40).getCell( 1));
        Cell copyFrom3 = new CellClone( sheet_2.getRow( 40).getCell( 2));
        Cell copyFrom4 = new CellClone( sheet_2.getRow( 41).getCell( 0));
        Cell copyFrom5 = new CellClone( sheet_2.getRow( 41).getCell( 1));
        Cell copyFrom6 = new CellClone( sheet_2.getRow( 41).getCell( 2));

        PoiUtil.copyRange( sheet_2, new CellRangeAddress( 40, 41, 0, 2), sheet_2, 41, 1, false);
        TestUtil.checkCell( copyFrom1, sheet_2.getRow( 41).getCell( 1));
        TestUtil.checkCell( copyFrom2, sheet_2.getRow( 41).getCell( 2));
        TestUtil.checkCell( copyFrom3, sheet_2.getRow( 41).getCell( 3));
        TestUtil.checkCell( copyFrom4, sheet_2.getRow( 42).getCell( 1));
        TestUtil.checkCell( copyFrom5, sheet_2.getRow( 42).getCell( 2));
        TestUtil.checkCell( copyFrom6, sheet_2.getRow( 42).getCell( 3));

        // No.14 コピー範囲を削除する（一時シートなし）
        copyFrom1 = new CellClone( sheet_2.getRow( 49).getCell( 0));
        PoiUtil.copyRange( sheet_2, new CellRangeAddress( 49, 49, 0, 0), sheet_2, 49, 2, true);
        assertNull( sheet_2.getRow( 49).getCell( 0));
        TestUtil.checkCell( copyFrom1, sheet_2.getRow( 49).getCell( 2));

        // No.15 コピー範囲を削除する（一時シートあり）
        copyFrom1 = new CellClone( sheet_2.getRow( 55).getCell( 0));
        copyFrom2 = new CellClone( sheet_2.getRow( 55).getCell( 1));
        copyFrom3 = new CellClone( sheet_2.getRow( 55).getCell( 2));
        copyFrom4 = new CellClone( sheet_2.getRow( 56).getCell( 0));
        copyFrom5 = new CellClone( sheet_2.getRow( 56).getCell( 1));
        copyFrom6 = new CellClone( sheet_2.getRow( 56).getCell( 2));

        PoiUtil.copyRange( sheet_2, new CellRangeAddress( 55, 56, 0, 2), sheet_2, 56, 1, true);
        assertNull( sheet_2.getRow( 55).getCell( 0));
        assertNull( sheet_2.getRow( 55).getCell( 1));
        assertNull( sheet_2.getRow( 55).getCell( 2));
        assertNull( sheet_2.getRow( 56).getCell( 0));
        TestUtil.checkCell( copyFrom1, sheet_2.getRow( 56).getCell( 1));
        TestUtil.checkCell( copyFrom2, sheet_2.getRow( 56).getCell( 2));
        TestUtil.checkCell( copyFrom3, sheet_2.getRow( 56).getCell( 3));
        TestUtil.checkCell( copyFrom4, sheet_2.getRow( 57).getCell( 1));
        TestUtil.checkCell( copyFrom5, sheet_2.getRow( 57).getCell( 2));
        TestUtil.checkCell( copyFrom6, sheet_2.getRow( 57).getCell( 3));

        // ===============================================
        // insertRangeDown( Sheet sheet, CellRangeAddress rangeAddress)
        // ===============================================
        // No.16 insertRangeDown
        copyFrom1 = new CellClone( sheet_3.getRow( 1).getCell( 1));
        copyFrom2 = new CellClone( sheet_3.getRow( 1).getCell( 2));
        copyFrom3 = new CellClone( sheet_3.getRow( 2).getCell( 1));
        copyFrom4 = new CellClone( sheet_3.getRow( 2).getCell( 2));
        PoiUtil.insertRangeDown( sheet_3, new CellRangeAddress( 1, 2, 1, 2));
        assertNull( sheet_3.getRow( 1).getCell( 1));
        assertNull( sheet_3.getRow( 1).getCell( 2));
        assertNull( sheet_3.getRow( 2).getCell( 1));
        assertNull( sheet_3.getRow( 2).getCell( 2));
        TestUtil.checkCell( copyFrom1, sheet_3.getRow( 3).getCell( 1));
        TestUtil.checkCell( copyFrom2, sheet_3.getRow( 3).getCell( 2));
        TestUtil.checkCell( copyFrom3, sheet_3.getRow( 4).getCell( 1));
        TestUtil.checkCell( copyFrom4, sheet_3.getRow( 4).getCell( 2));

        // ===============================================
        // insertRangeRight( Sheet sheet, CellRangeAddress rangeAddress)
        // ===============================================
        // No.17 insertRangeRight
        copyFrom1 = new CellClone( sheet_3.getRow( 6).getCell( 5));
        copyFrom2 = new CellClone( sheet_3.getRow( 6).getCell( 6));
        copyFrom3 = new CellClone( sheet_3.getRow( 7).getCell( 5));
        copyFrom4 = new CellClone( sheet_3.getRow( 7).getCell( 6));
        PoiUtil.insertRangeRight( sheet_3, new CellRangeAddress( 6, 7, 5, 6));
        assertNull( sheet_3.getRow( 6).getCell( 5));
        assertNull( sheet_3.getRow( 6).getCell( 6));
        assertNull( sheet_3.getRow( 7).getCell( 5));
        assertNull( sheet_3.getRow( 7).getCell( 6));
        TestUtil.checkCell( copyFrom1, sheet_3.getRow( 6).getCell( 7));
        TestUtil.checkCell( copyFrom2, sheet_3.getRow( 6).getCell( 8));
        TestUtil.checkCell( copyFrom3, sheet_3.getRow( 7).getCell( 7));
        TestUtil.checkCell( copyFrom4, sheet_3.getRow( 7).getCell( 8));

        // ===============================================
        // deleteRangeUp( Sheet sheet, CellRangeAddress rangeAddress)
        // ===============================================
        // No.18 deleteRangeUp
        copyFrom1 = new CellClone( sheet_3.getRow( 13).getCell( 9));
        copyFrom2 = new CellClone( sheet_3.getRow( 13).getCell( 10));
        copyFrom3 = new CellClone( sheet_3.getRow( 14).getCell( 9));
        copyFrom4 = new CellClone( sheet_3.getRow( 14).getCell( 10));
        PoiUtil.deleteRangeUp( sheet_3, new CellRangeAddress( 11, 12, 9, 10));
        assertNull( sheet_3.getRow( 13).getCell( 9));
        assertNull( sheet_3.getRow( 13).getCell( 10));
        assertNull( sheet_3.getRow( 14).getCell( 9));
        assertNull( sheet_3.getRow( 14).getCell( 10));
        TestUtil.checkCell( copyFrom1, sheet_3.getRow( 11).getCell( 9));
        TestUtil.checkCell( copyFrom2, sheet_3.getRow( 11).getCell( 10));
        TestUtil.checkCell( copyFrom3, sheet_3.getRow( 12).getCell( 9));
        TestUtil.checkCell( copyFrom4, sheet_3.getRow( 12).getCell( 10));

        // ===============================================
        // deleteRangeLeft( Sheet sheet, CellRangeAddress rangeAddress)
        // ===============================================
        // No.19 deleteRangeLeft
        copyFrom1 = new CellClone( sheet_3.getRow( 16).getCell( 15));
        copyFrom2 = new CellClone( sheet_3.getRow( 16).getCell( 14));
        copyFrom3 = new CellClone( sheet_3.getRow( 17).getCell( 15));
        copyFrom4 = new CellClone( sheet_3.getRow( 17).getCell( 14));
        PoiUtil.deleteRangeLeft( sheet_3, new CellRangeAddress( 16, 17, 13, 14));
        assertNull( sheet_3.getRow( 16).getCell( 15));
        assertNull( sheet_3.getRow( 16).getCell( 16));
        assertNull( sheet_3.getRow( 17).getCell( 15));
        assertNull( sheet_3.getRow( 17).getCell( 16));
        TestUtil.checkCell( copyFrom1, sheet_3.getRow( 16).getCell( 13));
        TestUtil.checkCell( copyFrom2, sheet_3.getRow( 16).getCell( 14));
        TestUtil.checkCell( copyFrom3, sheet_3.getRow( 17).getCell( 13));
        TestUtil.checkCell( copyFrom4, sheet_3.getRow( 17).getCell( 14));

        // ===============================================
        // clearRange( Sheet sheet, CellRangeAddress rangeAddress)
        // ===============================================
        // No.20 結合セルなし
        PoiUtil.clearRange( sheet_4, new CellRangeAddress( 0, 2, 0, 0));
        assertNull( sheet_4.getRow( 0).getCell( 0));
        assertNull( sheet_4.getRow( 1).getCell( 0));
        assertNull( sheet_4.getRow( 2).getCell( 0));
        assertEquals( "4", sheet_4.getRow( 3).getCell( 0).getStringCellValue());

        // No.21 結合セルなし
        PoiUtil.clearRange( sheet_4, new CellRangeAddress( 4, 5, 0, 1));
        assertNull( sheet_4.getRow( 4).getCell( 0));
        assertNull( sheet_4.getRow( 5).getCell( 0));
        assertNull( sheet_4.getRow( 4).getCell( 1));
        assertNull( sheet_4.getRow( 5).getCell( 1));
        assertEquals( "5C", sheet_4.getRow( 4).getCell( 2).getStringCellValue());
        assertEquals( "6C", sheet_4.getRow( 5).getCell( 2).getStringCellValue());

        // No.22 結合セルあり正常
        PoiUtil.clearRange( sheet_4, new CellRangeAddress( 8, 8, 0, 1));
        assertNull( null, sheet_4.getRow( 8).getCell( 0));

        // No.23 結合セルあり異常
        try {
            PoiUtil.clearRange( sheet_4, new CellRangeAddress( 10, 10, 0, 0));
            fail( "IllegalArgumentException expected, but no exception thrown.");
        } catch ( IllegalArgumentException ex) {
            // 横範囲内に結合部分が完全に入っていない場合、例外
        }
        // 消えていないことを確認
        assertEquals( "11", sheet_4.getRow( 10).getCell( 0).getStringCellValue());
        assertNotNull( sheet_4.getRow( 10).getCell( 1).getStringCellValue());

        // No.24 結合セルあり異常
        try {
            PoiUtil.clearRange( sheet_4, new CellRangeAddress( 12, 12, 0, 0));
            fail( "IllegalArgumentException expected, but no exception thrown.");
        } catch ( IllegalArgumentException ex) {
            // 縦範囲内に結合部分が完全に入っていない場合、例外
        }
        // 消えていないことを確認
        assertEquals( "13", sheet_4.getRow( 12).getCell( 0).getStringCellValue());
        assertNotNull( sheet_4.getRow( 13).getCell( 0).getStringCellValue());

        // ===============================================
        // clearCell( Sheet sheet, CellRangeAddress rangeAddress)
        // ===============================================
        // No.25 clearCell
        PoiUtil.clearCell( sheet_4, new CellRangeAddress( 15, 16, 0, 0));
        assertNull( sheet_4.getRow( 15).getCell( 0));
        assertNull( sheet_4.getRow( 15).getCell( 0));

        // ===============================================
        // setHyperlink( Cell cell, int type, String address)
        // ===============================================
        // No.26 setHyperlink
        Cell cellHyperlink = sheet_5.getRow( 0).getCell( 0);
        String address = "http://sourceforge.jp/projects/excella-core/";
        PoiUtil.setHyperlink( cellHyperlink, HyperlinkType.URL, address);
        Hyperlink hyperLink = cellHyperlink.getHyperlink();
        if ( hyperLink instanceof HSSFHyperlink) {
            assertEquals( address, (( HSSFHyperlink) hyperLink).getTextMark());
        } else if ( hyperLink instanceof XSSFHyperlink) {
            assertEquals( address, (( XSSFHyperlink) hyperLink).getAddress());
        }

        // ===============================================
        // setCellValue( Cell cell, Object value)
        // ===============================================
        // No.27 setCellValue
        Cell cellString = sheet_5.getRow( 1).getCell( 0);
        Cell cellNumber = sheet_5.getRow( 1).getCell( 1);
        Cell cellFloat = sheet_5.getRow( 1).getCell( 2);
        Cell cellDate = sheet_5.getRow( 1).getCell( 3);
        Cell cellBoolean = sheet_5.getRow( 1).getCell( 4);
        Cell cellNull = sheet_5.getRow( 1).getCell( 5);

        String stringValue = "aaa";
        Number numberValue = new Double( 10);
        Float floatValue = new Float( 10f);
        Date dateValue = new Date();
        Boolean booleanValue = Boolean.TRUE;

        PoiUtil.setCellValue( cellString, stringValue);
        PoiUtil.setCellValue( cellNumber, numberValue);
        PoiUtil.setCellValue( cellFloat, floatValue);
        PoiUtil.setCellValue( cellDate, dateValue);
        PoiUtil.setCellValue( cellBoolean, booleanValue);
        PoiUtil.setCellValue( cellNull, null);

        assertEquals( stringValue, cellString.getStringCellValue());
        assertEquals( numberValue, cellNumber.getNumericCellValue());
        assertEquals( new Double( String.valueOf( floatValue)), ( Double) cellFloat.getNumericCellValue());
        assertEquals( dateValue, cellDate.getDateCellValue());
        assertEquals( booleanValue, cellBoolean.getBooleanCellValue());
        assertNull( PoiUtil.getCellValue( cellNull));

        // No.28 セルがnull
        try {
            PoiUtil.setCellValue( null, stringValue);
            fail( "NullPointerException expected, but no exception thrown.");
        } catch ( NullPointerException ex) {
            // セルがnullの場合は例外が発生
        }

        // ===============================================
        // getLastColNum( Sheet sheet)
        // ===============================================
        // No.29 ブランクシート
        int lastColNum1 = PoiUtil.getLastColNum( sheet_6);
        assertEquals( -1, lastColNum1);

        // No.30 通常のシート
        int lastColNum2 = PoiUtil.getLastColNum( sheet_7);
        assertEquals( 10, lastColNum2);
    }
}
