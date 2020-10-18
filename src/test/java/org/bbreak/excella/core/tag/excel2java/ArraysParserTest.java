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

package org.bbreak.excella.core.tag.excel2java;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * ArraysParserテストクラス
 * 
 * @since 1.0
 */
public class ArraysParserTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public ArraysParserTest(String version) {
        super(version);
    }

    @BeforeClass
    public static void setUpBeforeClass() throws Exception {
    }

    @AfterClass
    public static void tearDownAfterClass() throws Exception {
    }

    @Before
    public void setUp() throws Exception {
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public final void testArraysParser() throws ParseException, IOException {
        Workbook wk = getWorkbook();
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        ArraysParser arraysParser = new ArraysParser( "@Arrays");
        String tag = arraysParser.getTag();
        arraysParser.setTag( tag);
        Cell tagCell = null;
        Object data = null;
        List<Object[]> list = null;

        // No.1 パラメータ無し
        tagCell = sheet1.getRow( 5).getCell( 0);
        list = arraysParser.parse( sheet1, tagCell, data);
        Object[] arrays = list.get( 0);
        assertEquals( "value1-1", arrays[0]);
        assertEquals( "value2-1", arrays[1]);
        assertEquals( "value3-1", arrays[2]);
        assertEquals( 3, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-2", arrays[0]);
        assertEquals( "value2-2", arrays[1]);
        assertEquals( "value3-2", arrays[2]);
        assertEquals( 3, arrays.length);
        arrays = list.get( 2);
        assertEquals( "value1-3", arrays[0]);
        assertEquals( "value2-3", arrays[1]);
        assertEquals( "value3-3", arrays[2]);
        assertEquals( 3, arrays.length);
        arrays = list.get( 3);
        assertEquals( "value1-4", arrays[0]);
        assertEquals( "value2-4", arrays[1]);
        assertEquals( "value3-4", arrays[2]);
        assertEquals( "value4-1", arrays[3]);
        assertEquals( 4, arrays.length);
        arrays = list.get( 4);
        assertEquals( "value1-5", arrays[0]);
        assertEquals( null, arrays[1]);
        assertEquals( "value3-5", arrays[2]);
        assertEquals( "value4-2", arrays[3]);
        assertEquals( 4, arrays.length);
        arrays = list.get( 5);
        assertEquals( "value1-6", arrays[0]);
        assertEquals( "value2-5", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 6, list.size());
        
        // No.2 パラメータ有り
        tagCell = sheet2.getRow( 5).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value1-1", arrays[0]);
        assertEquals( "value2-1", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-2", arrays[0]);
        assertEquals( "value2-2", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 2);
        assertEquals( "value1-3", arrays[0]);
        assertEquals( "value2-3", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 3, list.size());
        
        // No.3 開始列、終了列指定
        tagCell = sheet2.getRow( 13).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value1-4", arrays[0]);
        assertEquals( "value2-4", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-5", arrays[0]);
        assertEquals( "value2-5", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 2);
        assertEquals( "value1-6", arrays[0]);
        assertEquals( "value2-6", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 3, list.size());

        // No.4 データ行にnull行有り
        tagCell = sheet2.getRow( 22).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value1-7", arrays[0]);
        assertEquals( "value2-7", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-8", arrays[0]);
        assertEquals( "value2-8", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 2);
        assertEquals( "value1-9", arrays[0]);
        assertEquals( "value2-9", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 3, list.size());

        // No.5 データ行にnullセル有り
        tagCell = sheet2.getRow( 31).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( null, arrays[0]);
        assertEquals( "value2-10", arrays[1]);
        assertEquals( "value3-1", arrays[2]);
        assertEquals( 3, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-10", arrays[0]);
        assertEquals( null, arrays[1]);
        assertEquals( "value3-2", arrays[2]);
        assertEquals( 3, arrays.length);
        arrays = list.get( 2);
        assertEquals( "value1-11", arrays[0]);
        assertEquals( "value2-11", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 3, list.size());

        // No.6 開始列指定
        tagCell = sheet2.getRow( 39).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value2-12", arrays[0]);
        assertEquals( "value3-3", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value2-13", arrays[0]);
        assertEquals( "value3-4", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 2, list.size());

        // No.7 終了列指定
        tagCell = sheet2.getRow( 46).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value1-14", arrays[0]);
        assertEquals( "value2-14", arrays[1]);
        assertEquals( null, arrays[2]);
        assertEquals( null, arrays[3]);
        assertEquals( null, arrays[4]);
        assertEquals( 5, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-15", arrays[0]);
        assertEquals( "value2-15", arrays[1]);
        assertEquals( "value3-5", arrays[2]);
        assertEquals( null, arrays[3]);
        assertEquals( null, arrays[4]);
        assertEquals( 5, arrays.length);
        assertEquals( 2, list.size());

        // No.8 開始列・終了列指定
        tagCell = sheet2.getRow( 53).getCell( 1);
        list.clear();
        list = arraysParser.parse( sheet2, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value3-6", arrays[0]);
        assertEquals( "value4-1", arrays[1]);
        assertEquals( null, arrays[2]);
        assertEquals( null, arrays[3]);
        assertEquals( 4, arrays.length);
        arrays = list.get( 1);
        assertEquals( null, arrays[0]);
        assertEquals( "value4-2", arrays[1]);
        assertEquals( "value5-1", arrays[2]);
        assertEquals( null, arrays[3]);
        assertEquals( 4, arrays.length);
        arrays = list.get( 2);
        assertEquals( null, arrays[0]);
        assertEquals( null, arrays[1]);
        assertEquals( null, arrays[2]);
        assertEquals( null, arrays[3]);
        assertEquals( 4, arrays.length);
        assertEquals( 3, list.size());

        // No.9 マイナス範囲指定（データ行）
        tagCell = sheet3.getRow( 7).getCell( 0);
        list.clear();
        list = arraysParser.parse( sheet3, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value1-1", arrays[0]);
        assertEquals( "value2-1", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value1-2", arrays[0]);
        assertEquals( "value2-2", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 2);
        assertEquals( "value1-3", arrays[0]);
        assertEquals( "value2-3", arrays[1]);
        assertEquals( 2, arrays.length);
        arrays = list.get( 3);
        assertEquals( "value1-4", arrays[0]);
        assertEquals( "value2-4", arrays[1]);
        assertEquals( 2, arrays.length);
        assertEquals( 4, list.size());

        // No.10 マイナス範囲指定（データ列）
        tagCell = sheet3.getRow( 12).getCell( 2);
        list.clear();
        list = arraysParser.parse( sheet3, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value2-6", arrays[0]);
        assertEquals( "value3-1", arrays[1]);
        assertEquals( "value4-1", arrays[2]);
        assertEquals( "value5-1", arrays[3]);
        assertEquals( 4, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value2-7", arrays[0]);
        assertEquals( "value3-2", arrays[1]);
        assertEquals( "value4-2", arrays[2]);
        assertEquals( "value5-2", arrays[3]);
        assertEquals( 4, arrays.length);
        assertEquals( 2, list.size());

        // No.11 マイナス範囲指定（データ行・データ列）
        tagCell = sheet3.getRow( 21).getCell( 2);
        list.clear();
        list = arraysParser.parse( sheet3, tagCell, data);
        arrays = list.get( 0);
        assertEquals( "value2-8", arrays[0]);
        assertEquals( "value3-3", arrays[1]);
        assertEquals( "value4-3", arrays[2]);
        assertEquals( "value5-3", arrays[3]);
        assertEquals( 4, arrays.length);
        arrays = list.get( 1);
        assertEquals( "value2-9", arrays[0]);
        assertEquals( "value3-4", arrays[1]);
        assertEquals( "value4-4", arrays[2]);
        assertEquals( "value5-4", arrays[3]);
        assertEquals( 4, arrays.length);
        assertEquals( 2, list.size());

        // No.12 DataRowFrom > DataRowTo
        tagCell = sheet3.getRow( 30).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 30, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.12:" + pe);
        }

        // No.13 DataColumnFrom > DataColumnTo
        tagCell = sheet3.getRow( 40).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 40, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.13:" + pe);
        }

        // No.14 DataRowFrom不正（値が空）
        tagCell = sheet3.getRow( 48).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 48, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.14:" + pe);
        }

        // No.15 DataRowTo不正（値が空）
        tagCell = sheet3.getRow( 51).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 51, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.15:" + pe);
        }

        // No.16 DataColumnFrom不正（値が空）
        tagCell = sheet3.getRow( 54).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 54, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.16:" + pe);
        }

        // No.17 DataColumnTo不正（値が空）
        tagCell = sheet3.getRow( 57).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 57, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.17:" + pe);
        }

        // No.18 DataRowFrom不正（値が文字）
        tagCell = sheet3.getRow( 60).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 60, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.18:" + pe);
        }

        // No.19 DataRowTo不正（値が文字）
        tagCell = sheet3.getRow( 63).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 63, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.19:" + pe);
        }

        // No.20 DataColumnFrom不正（値が文字）
        tagCell = sheet3.getRow( 66).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 66, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.20:" + pe);
        }

        // No.21 DataColumnTo不正（値が文字）
        tagCell = sheet3.getRow( 69).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 69, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.21:" + pe);
        }
        
        // No.22 DataColumnFrom不正テスト（A列でデータ開始列にマイナスを指定）
        tagCell = sheet4.getRow( 2).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 2, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.22:" + pe);
        }
        
        // No.23 DataColumnTo不正テスト（A列でデータ終了列にマイナスを指定）
        tagCell = sheet4.getRow( 11).getCell( 0);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 11, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.23:" + pe);
        }
        
        // No.24 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        tagCell = sheet4.getRow( 0).getCell( 3);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 3, cell.getColumnIndex());
            System.out.println( "No.24:" + pe);
        }

        // No.25 DataRowTo不正テスト（1行目でデータ開始行にマイナスを指定）
        tagCell = sheet4.getRow( 0).getCell( 6);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 6, cell.getColumnIndex());
            System.out.println( "No.25:" + pe);
        }

        // No.26 DataColumnFrom不正テスト（最終列でキー列にプラスを指定）
        tagCell = sheet4.getRow( 2).getCell( 11);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 2, cell.getRow().getRowNum());
            assertEquals( 11, cell.getColumnIndex());
            System.out.println( "No.26:" + pe);
        }
        
        // No.27 DataColumnTo不正テスト（最終列で値列にプラスを指定）
        tagCell = sheet4.getRow( 11).getCell( 11);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 11, cell.getRow().getRowNum());
            assertEquals( 11, cell.getColumnIndex());
            System.out.println( "No.27:" + pe);
        }
        
        // No.28 DataRowFrom不正テスト（最終行でデータ開始行にプラスを指定）
        tagCell = sheet4.getRow( 18).getCell( 3);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 3, cell.getColumnIndex());
            System.out.println( "No.28:" + pe);
        }

        // No.29 DataRowTo不正テスト（最終行でデータ終了行にプラスを指定）
        tagCell = sheet4.getRow( 18).getCell( 6);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 6, cell.getColumnIndex());
            System.out.println( "No.29:" + pe);
        }
        
        // No.30 データ無し
        tagCell = sheet4.getRow( 18).getCell( 9);
        list.clear();
        try {
            list = arraysParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 9, cell.getColumnIndex());
            System.out.println( "No.30:" + pe);
        }
    }
}
