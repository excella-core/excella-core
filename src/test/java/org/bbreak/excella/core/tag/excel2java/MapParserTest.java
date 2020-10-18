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
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.util.Map;
import java.util.Set;

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
 * MapParserテストクラス
 * 
 * @since 1.0
 */
public class MapParserTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public MapParserTest(String version) {
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
    public final void testMapParser() throws ParseException, IOException {
        Workbook wk = getWorkbook();
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        MapParser mapParser = new MapParser( "@Map");
        Cell tagCell = null;
        Object data = null;
        Map<?, ?> map = null;

        // No.1 パラメータ無し
        tagCell = sheet1.getRow( 3).getCell( 0);
        map = mapParser.parse( sheet1, tagCell, data);
        Set<?> keySet = map.keySet();
        assertEquals( 25, keySet.size());
        assertEquals( "値1", map.get( "キー1"));
        assertEquals( "値2", map.get( "キー2"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=5}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=5}"));
        assertEquals( "値3", map.get( "キー3"));
        assertEquals( "値4", map.get( "キー4"));
        assertEquals( "値5", map.get( "キー5"));
        assertEquals( "値6", map.get( "キー6"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,Key=キー}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=3,Key=キー}"));
        assertEquals( "値7", map.get( "キー7"));
        assertEquals( "値8", map.get( "キー8"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,Value=値}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=3,Value=値}"));
        assertEquals( "値9", map.get( "キー9"));
        assertEquals( "値10", map.get( "キー10"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,KeyColumn=2}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=3,KeyColumn=2}"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,ValueColumn=2}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=3,ValueColumn=2}"));
        assertEquals( "備考", map.get( "キー13"));
        assertEquals( "備考", map.get( "キー14"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=4,KeyCell=2:2}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=4,KeyCell=2:2}"));
        assertEquals( "値17", map.get( "備考"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=4,ValueCell=4:3}"));
        assertEquals( null, map.get( "@Map{DataRowFrom=2,DataRowTo=4,ValueCell=4:3}"));
        assertEquals( "備考", map.get( "キー18"));
        assertEquals( "備考", map.get( "キー19"));
        assertEquals( "備考", map.get( "キー20"));

        // No.2 パラメータ有り
        tagCell = sheet1.getRow( 7).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値3", map.get( "キー3"));
        assertEquals( "値4", map.get( "キー4"));
        assertEquals( "値5", map.get( "キー5"));
        assertEquals( "値6", map.get( "キー6"));

        // No.3 キー指定
        tagCell = sheet1.getRow( 15).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "値8", map.get( "キー"));

        // No.4 値指定
        tagCell = sheet1.getRow( 24).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値", map.get( "キー9"));
        assertEquals( "値", map.get( "キー10"));

        // No.5 キー列指定
        tagCell = sheet1.getRow( 32).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値11", map.get( "キー11"));
        assertEquals( "値12", map.get( "キー12"));

        // No.6 値列指定
        tagCell = sheet1.getRow( 41).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値13", map.get( "キー13"));
        assertEquals( "値14", map.get( "キー14"));

        // No.7 キーセル指定
        tagCell = sheet1.getRow( 49).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "値17", map.get( "キー15"));

        // No.8 値セル指定
        tagCell = sheet1.getRow( 58).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値20", map.get( "キー18"));
        assertEquals( "値20", map.get( "キー19"));
        assertEquals( "値20", map.get( "キー20"));

        // No.9 DataRowFrom > DataRowTo
        tagCell = sheet1.getRow( 3).getCell( 7);
        map.clear();
        try {
            map = mapParser.parse( sheet1, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 3, cell.getRow().getRowNum());
            assertEquals( 7, cell.getColumnIndex());
            System.out.println( "No.9:" + pe);
        }

        // No.10 マイナス範囲指定
        tagCell = sheet1.getRow( 18).getCell( 7);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値5", map.get( "キー5"));
        assertEquals( "値6", map.get( "キー6"));
        assertEquals( "値7", map.get( "キー7"));
        assertEquals( "値8", map.get( "キー8"));

        // No.11 キー列マイナス指定
        tagCell = sheet1.getRow( 23).getCell( 8);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値9", map.get( "キー9"));
        assertEquals( "値10", map.get( "キー10"));
        assertEquals( "値11", map.get( "キー11"));
        assertEquals( "値12", map.get( "キー12"));

        // No.12 値列マイナス指定
        tagCell = sheet1.getRow( 33).getCell( 8);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値13", map.get( "キー13"));
        assertEquals( "値14", map.get( "キー14"));
        assertEquals( "値15", map.get( "キー15"));
        assertEquals( "値16", map.get( "キー16"));

        // No.13 キーセルマイナス指定
        tagCell = sheet1.getRow( 43).getCell( 8);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "値20", map.get( "キー"));

        // No.14 値セルマイナス指定
        tagCell = sheet1.getRow( 53).getCell( 8);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値", map.get( "キー21"));
        assertEquals( "値", map.get( "キー22"));
        assertEquals( "値", map.get( "キー23"));
        assertEquals( "値", map.get( "キー24"));

        // No.15 キー・値を同列に指定
        tagCell = sheet1.getRow( 63).getCell( 7);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "キー・値25", map.get( "キー・値25"));
        assertEquals( "キー・値26", map.get( "キー・値26"));

        // No.16 キー・値にタグと同行・同列を指定
        tagCell = sheet1.getRow( 71).getCell( 7);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "@Map{DataRowFrom=0,DataRowTo=0,ValueColumn=0}", map.get( "@Map{DataRowFrom=0,DataRowTo=0,ValueColumn=0}"));

        // No.17 キー・値セルにタグセルを指定
        tagCell = sheet1.getRow( 76).getCell( 7);
        map.clear();
        map = mapParser.parse( sheet1, tagCell, data);
        keySet = map.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "@Map{DataRowFrom=0,DataRowTo=0,KeyCell=0:0,ValueCell=0:0}", map.get( "@Map{DataRowFrom=0,DataRowTo=0,KeyCell=0:0,ValueCell=0:0}"));

        // No.18 キーセルがnull
        tagCell = sheet2.getRow( 3).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値1", map.get( "キー1"));
        assertTrue( keySet.contains( null));
        assertEquals( "値2", map.get( null));
        assertEquals( "値3", map.get( "キー2"));

        // No.19 値セルがnull
        tagCell = sheet2.getRow( 11).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値4", map.get( "キー3"));
        assertEquals( null, map.get( "キー4"));
        assertEquals( "値5", map.get( "キー5"));

        // No.20 データ行がnull
        tagCell = sheet2.getRow( 19).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値6", map.get( "キー6"));
        assertEquals( "値7", map.get( "キー7"));

        // No.21 キーセル値がnull
        tagCell = sheet2.getRow( 27).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値8", map.get( "キー8"));
        assertEquals( "値9", map.get( null));
        assertEquals( "値10", map.get( "キー9"));

        // No.22 値セルの値がnull
        tagCell = sheet2.getRow( 35).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値11", map.get( "キー10"));
        assertTrue( keySet.contains( "キー11"));
        assertEquals( null, map.get( "キー11"));
        assertEquals( "値12", map.get( "キー12"));

        // No.23 キーセルの値・値セルの値がnull
        tagCell = sheet2.getRow( 43).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値13", map.get( "キー13"));
        assertTrue( keySet.contains( null));
        assertEquals( null, map.get( null));
        assertEquals( "値14", map.get( "キー14"));

        // No.24 値セルの行がnull
        tagCell = sheet2.getRow( 51).getCell( 0);
        map.clear();
        map = mapParser.parse( sheet2, tagCell, data);
        keySet = map.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( null, map.get( "キー15"));
        assertEquals( null, map.get( "キー16"));

        // No.25 重複定義：Key, KeyColumn, KeyCell
        tagCell = sheet3.getRow( 4).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 4, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.25:" + pe);
        }

        // No.26 重複定義：Key, KeyColumn
        tagCell = sheet3.getRow( 8).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 8, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.26:" + pe);
        }

        // No.27 重複定義：Key, KeyCell
        tagCell = sheet3.getRow( 12).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 12, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.27:" + pe);
        }

        // No.28 重複定義：KeyColumn, KeyCell
        tagCell = sheet3.getRow( 16).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet1, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 16, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.28:" + pe);
        }

        // No.29 重複定義：Value, ValueColumn, ValueCell
        tagCell = sheet3.getRow( 20).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 20, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.29:" + pe);
        }

        // No.30 重複定義：Value, ValueColumn
        tagCell = sheet3.getRow( 24).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 24, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.30:" + pe);
        }

        // No.31 重複定義：Value, ValueCell
        tagCell = sheet3.getRow( 28).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 28, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.31:" + pe);
        }

        // No.32 重複定義：ValueColumn, ValueCell
        tagCell = sheet3.getRow( 32).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 32, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.32:" + pe);
        }

        // No.33 データ無し
        tagCell = sheet3.getRow( 38).getCell( 1);
        map.clear();
        map = mapParser.parse( sheet3, tagCell, data);

        // No.34 DataRowFrom不正（値が空）
        tagCell = sheet3.getRow( 41).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 41, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.33:" + pe);
        }

        // No.35 DataRowTo不正（値が空）
        tagCell = sheet3.getRow( 44).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 44, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.35:" + pe);
        }

        // No.36 Key不正（値が空）
        tagCell = sheet3.getRow( 47).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 47, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.36:" + pe);
        }

        // No.37 Value不正（値が空）
        tagCell = sheet3.getRow( 50).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 50, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.37:" + pe);
        }

        // No.38 KeyColumn不正（値が空）
        tagCell = sheet3.getRow( 53).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 53, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.38:" + pe);
        }

        // No.39 ValueColumn不正（値が空）
        tagCell = sheet3.getRow( 56).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 56, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.39:" + pe);
        }

        // No.40 KeyCell不正（値が空）
        tagCell = sheet3.getRow( 59).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 59, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.40:" + pe);
        }

        // No.41 ValueCell不正（値が空）
        tagCell = sheet3.getRow( 62).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 62, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.41:" + pe);
        }

        // No.42 DataRowFrom不正（値が文字）
        tagCell = sheet3.getRow( 65).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 65, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.42:" + pe);
        }

        // No.43 DataRowTo不正（値が文字）
        tagCell = sheet3.getRow( 68).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 68, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.43:" + pe);
        }

        // No.44 KeyColumen不正（値が文字）
        tagCell = sheet3.getRow( 71).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 71, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.44:" + pe);
        }

        // No.45 ValueColumn不正（値が文字）
        tagCell = sheet3.getRow( 74).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 74, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.45:" + pe);
        }

        // No.46 KeyCell不正（値が文字）
        tagCell = sheet3.getRow( 77).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 77, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.46:" + pe);
        }

        // No.47 ValueCell不正（値が文字）
        tagCell = sheet3.getRow( 80).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet3, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 80, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.47:" + pe);
        }

        // No.48 KeyColumn不正テスト（A列でキー列にマイナスを指定）
        tagCell = sheet4.getRow( 2).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 2, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.48:" + pe);
        }

        // No.49 ValueColumn不正テスト（A列で値列にマイナスを指定）
        tagCell = sheet4.getRow( 11).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 11, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.49:" + pe);
        }

        // No.50 KeyCell不正テスト（A列でキーセル列にマイナスを指定）
        tagCell = sheet4.getRow( 20).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 20, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.50:" + pe);
        }

        // No.51 ValueCell不正テスト（A列で値セル列にマイナスを指定）
        tagCell = sheet4.getRow( 29).getCell( 0);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 29, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.51:" + pe);
        }

        // No.52 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        tagCell = sheet4.getRow( 0).getCell( 3);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 3, cell.getColumnIndex());
            System.out.println( "No.52:" + pe);
        }

        // No.53 DataRowTo不正テスト（1行目でデータ終了行にマイナスを指定）
        tagCell = sheet4.getRow( 0).getCell( 6);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 6, cell.getColumnIndex());
            System.out.println( "No.53:" + pe);
        }

        // No.54 KeyCell不正テスト（1行目でキーセル行にマイナスを指定）
        tagCell = sheet4.getRow( 0).getCell( 9);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 9, cell.getColumnIndex());
            System.out.println( "No.54:" + pe);
        }

        // No.55 ValueCell不正テスト（1行目で値セル行にマイナスを指定）
        tagCell = sheet4.getRow( 0).getCell( 12);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 12, cell.getColumnIndex());
            System.out.println( "No.55:" + pe);
        }

        // No.56 KeyColumn不正テスト（最終列でキー列にプラスを設定）
        tagCell = sheet4.getRow( 3).getCell( 16);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 3, cell.getRow().getRowNum());
            assertEquals( 16, cell.getColumnIndex());
            System.out.println( "No.56:" + pe);
        }

        // No.57 ValueColumn不正テスト（最終列で値列にプラスを設定）
        tagCell = sheet4.getRow( 12).getCell( 16);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 12, cell.getRow().getRowNum());
            assertEquals( 16, cell.getColumnIndex());
            System.out.println( "No.57:" + pe);
        }

        // No.58 KeyCell不正テスト（最終列でキーセル列にプラスを設定）
        tagCell = sheet4.getRow( 21).getCell( 16);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 21, cell.getRow().getRowNum());
            assertEquals( 16, cell.getColumnIndex());
            System.out.println( "No.58:" + pe);
        }

        // No.59 ValueCell不正テスト（最終列で値セル列にプラスを設定）
        tagCell = sheet4.getRow( 30).getCell( 16);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 30, cell.getRow().getRowNum());
            assertEquals( 16, cell.getColumnIndex());
            System.out.println( "No.59:" + pe);
        }

        // No.60 DataRowFrom不正テスト（最終行でデータ開始行にプラスを設定）
        tagCell = sheet4.getRow( 34).getCell( 3);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 34, cell.getRow().getRowNum());
            assertEquals( 3, cell.getColumnIndex());
            System.out.println( "No.60:" + pe);
        }

        // No.61 DataRowTo不正テスト（最終行でデータ終了行にプラスを設定）
        tagCell = sheet4.getRow( 34).getCell( 6);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 34, cell.getRow().getRowNum());
            assertEquals( 6, cell.getColumnIndex());
            System.out.println( "No.61:" + pe);
        }

        // No.62 KeyCell不正テスト（最終行でキーセル行にプラスを設定）
        tagCell = sheet4.getRow( 34).getCell( 9);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 34, cell.getRow().getRowNum());
            assertEquals( 9, cell.getColumnIndex());
            System.out.println( "No.62:" + pe);
        }

        // No.63 ValueCell不正テスト（最終行で値セル行にプラスを設定）
        tagCell = sheet4.getRow( 34).getCell( 12);
        map.clear();
        try {
            map = mapParser.parse( sheet4, tagCell, data);
            fail( "ParseException expected, but no exception thrown,");
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 34, cell.getRow().getRowNum());
            assertEquals( 12, cell.getColumnIndex());
            System.out.println( "No.63:" + pe);
        }
    }
}
