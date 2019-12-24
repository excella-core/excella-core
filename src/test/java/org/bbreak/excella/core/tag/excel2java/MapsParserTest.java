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

import java.util.List;
import java.util.Map;

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
 * MapsParserテストクラス
 * 
 * @since 1.0
 */
public class MapsParserTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public MapsParserTest( String version) {
        super( version);
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
    public final void testMapsParser() throws ParseException {
        Workbook wk = getWorkbook();
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        Sheet sheet5 = wk.getSheetAt( 4);
        MapsParser mapsParser = new MapsParser( "@Maps");
        Cell tagCell = null;
        Object data = null;
        List<Map<?, ?>> maps = null;

        // No.1 パラメータ無し
        tagCell = sheet1.getRow( 5).getCell( 0);
        maps = mapsParser.parse( sheet1, tagCell, data);
        assertEquals( "value1-1", maps.get( 0).get( "key1"));
        assertEquals( "value2-1", maps.get( 0).get( "key2"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value1-2", maps.get( 1).get( "key1"));
        assertEquals( "value2-2", maps.get( 1).get( "key2"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( "value1-3", maps.get( 2).get( "key1"));
        assertEquals( "value2-3", maps.get( 2).get( "key2"));
        assertEquals( 2, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.2 データ無し
        tagCell = sheet2.getRow( 5).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet2, tagCell, data);
        assertEquals( 0, maps.size());

        // No.3 パラメータ有り
        tagCell = sheet3.getRow( 5).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value1-1", maps.get( 0).get( "key1"));
        assertEquals( "value2-1", maps.get( 0).get( "key2"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value1-2", maps.get( 1).get( "key1"));
        assertEquals( "value2-2", maps.get( 1).get( "key2"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( 2, maps.size());

        // No.4 キー列、開始列、終了列指定
        tagCell = sheet3.getRow( 13).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value3-1", maps.get( 0).get( "key3"));
        assertEquals( "value4-1", maps.get( 0).get( "key4"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value3-2", maps.get( 1).get( "key3"));
        assertEquals( "value4-2", maps.get( 1).get( "key4"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( 2, maps.size());

        // No.5 コメント有り
        tagCell = sheet3.getRow( 23).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value5-1", maps.get( 0).get( "key5"));
        assertEquals( "value6-1", maps.get( 0).get( "key6"));
        assertEquals( null, maps.get( 2).get( "key7"));
        assertEquals( "value8-1", maps.get( 0).get( "key8"));
        assertEquals( 3, maps.get( 0).keySet().size());
        assertEquals( "value5-2", maps.get( 1).get( "key5"));
        assertEquals( "value6-2", maps.get( 1).get( "key6"));
        assertEquals( null, maps.get( 2).get( "key7"));
        assertEquals( "value8-2", maps.get( 1).get( "key8"));
        assertEquals( 3, maps.get( 1).keySet().size());
        assertEquals( "value5-3", maps.get( 2).get( "key5"));
        assertEquals( "value6-3", maps.get( 2).get( "key6"));
        assertEquals( null, maps.get( 2).get( "key7"));
        assertEquals( "value8-3", maps.get( 2).get( "key8"));
        assertEquals( 3, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.6 キー行がnull
        tagCell = sheet3.getRow( 31).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( 0, maps.size());

        // No.7 キーセルがnull
        tagCell = sheet3.getRow( 39).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value9-1", maps.get( 0).get( "key9"));
        assertEquals( "value10-1", maps.get( 0).get( "key10"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value9-2", maps.get( 1).get( "key9"));
        assertEquals( "value10-2", maps.get( 1).get( "key10"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( "value9-3", maps.get( 2).get( "key9"));
        assertEquals( "value10-3", maps.get( 2).get( "key10"));
        assertEquals( 2, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.8 キーセル値がnull
        tagCell = sheet3.getRow( 47).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value11-1", maps.get( 0).get( "key11"));
        assertEquals( "value12-1", maps.get( 0).get( "key12"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value11-2", maps.get( 1).get( "key11"));
        assertEquals( "value12-2", maps.get( 1).get( "key12"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( "value11-3", maps.get( 2).get( "key11"));
        assertEquals( "value12-3", maps.get( 2).get( "key12"));
        assertEquals( 2, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.9 値行がnull
        tagCell = sheet3.getRow( 55).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value13-1", maps.get( 0).get( "key13"));
        assertEquals( "value14-1", maps.get( 0).get( "key14"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value13-2", maps.get( 1).get( "key13"));
        assertEquals( "value14-2", maps.get( 1).get( "key14"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( 2, maps.size());

        // No.10 値セルがnull
        tagCell = sheet3.getRow( 63).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value15-1", maps.get( 0).get( "key15"));
        assertEquals( "value16-1", maps.get( 0).get( "key16"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( null, maps.get( 1).get( "key15"));
        assertEquals( null, maps.get( 1).get( "key16"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( "value15-2", maps.get( 2).get( "key15"));
        assertEquals( "value16-2", maps.get( 2).get( "key16"));
        assertEquals( 2, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.11 値セル値がnull
        tagCell = sheet3.getRow( 71).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value17-1", maps.get( 0).get( "key17"));
        assertEquals( "value18-1", maps.get( 0).get( "key18"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( null, maps.get( 1).get( "key17"));
        assertEquals( null, maps.get( 1).get( "key18"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( "value17-2", maps.get( 2).get( "key17"));
        assertEquals( "value18-2", maps.get( 2).get( "key18"));
        assertEquals( 2, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.12 重複キー有り
        tagCell = sheet3.getRow( 79).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet3, tagCell, data);
        assertEquals( "value19-4", maps.get( 0).get( "key19"));
        assertEquals( 1, maps.get( 0).keySet().size());
        assertEquals( null, maps.get( 1).get( "key19"));
        assertEquals( 1, maps.get( 1).keySet().size());
        assertEquals( "value19-5", maps.get( 2).get( "key19"));
        assertEquals( 1, maps.get( 2).keySet().size());
        assertEquals( 3, maps.size());

        // No.13 マイナス範囲指定
        tagCell = sheet4.getRow( 7).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet4, tagCell, data);
        assertEquals( "value1-1", maps.get( 0).get( "key1"));
        assertEquals( "value2-1", maps.get( 0).get( "key2"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( "value1-2", maps.get( 1).get( "key1"));
        assertEquals( "value2-2", maps.get( 1).get( "key2"));
        assertEquals( 2, maps.get( 1).keySet().size());
        assertEquals( "value1-3", maps.get( 2).get( "key1"));
        assertEquals( "value2-3", maps.get( 2).get( "key2"));
        assertEquals( 2, maps.get( 2).keySet().size());
        assertEquals( "value1-4", maps.get( 3).get( "key1"));
        assertEquals( "value2-4", maps.get( 3).get( "key2"));
        assertEquals( 2, maps.get( 3).keySet().size());
        assertEquals( 4, maps.size());

        // No.14 キー・データ同列
        tagCell = sheet4.getRow( 12).getCell( 0);
        maps.clear();
        maps = mapsParser.parse( sheet4, tagCell, data);
        assertEquals( "キー・値1", maps.get( 0).get( "キー・値1"));
        assertEquals( "キー・値2", maps.get( 0).get( "キー・値2"));
        assertEquals( 2, maps.get( 0).keySet().size());
        assertEquals( 1, maps.size());

        // No.15 DataRowFrom > DataRowTo
        tagCell = sheet4.getRow( 21).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 21, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.15:" + pe);
        }

        // No.16 KeyRow不正（値が空）
        tagCell = sheet4.getRow( 27).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 27, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.16:" + pe);
        }

        // No.17 DataRowFrom不正（値が空）
        tagCell = sheet4.getRow( 30).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 30, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.17:" + pe);
        }

        // No.18 DataRowTo不正（値が空）
        tagCell = sheet4.getRow( 33).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 33, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.18:" + pe);
        }

        // No.19 KeyRow不正（値が文字）
        tagCell = sheet4.getRow( 36).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 36, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.19:" + pe);
        }

        // No.20 DataRowFrom不正（値が文字）
        tagCell = sheet4.getRow( 39).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 39, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.20:" + pe);
        }

        // No.21 DataRowFrom不正（値が文字）
        tagCell = sheet4.getRow( 42).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet4, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 42, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.21:" + pe);
        }

        // No.22 KeyRow不正テスト（1行目でキー行にマイナスを指定）
        tagCell = sheet5.getRow( 0).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.22:" + pe);
        }

        // No.23 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        tagCell = sheet5.getRow( 0).getCell( 3);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 3, cell.getColumnIndex());
            System.out.println( "No.23:" + pe);
        }

        // No.24 DataRowTo不正テスト（1行目でデータ終了行にマイナスを指定）
        tagCell = sheet5.getRow( 0).getCell( 6);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 6, cell.getColumnIndex());
            System.out.println( "No.24:" + pe);
        }

        // No.25 KeyRow不正テスト（最終行でキー行にプラスを指定）
        tagCell = sheet5.getRow( 18).getCell( 0);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.25:" + pe);
        }

        // No.26 DataRowFrom不正テスト（最終行でデータ開始行にプラスを指定）
        tagCell = sheet5.getRow( 18).getCell( 3);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 3, cell.getColumnIndex());
            System.out.println( "No.26:" + pe);
        }

        // No.27 DataRowTo不正テスト（最終行でデータ終了行にプラスを指定）
        tagCell = sheet5.getRow( 18).getCell( 6);
        maps.clear();
        try {
            maps = mapsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 6, cell.getColumnIndex());
            System.out.println( "No.27:" + pe);
        }
    }

}
