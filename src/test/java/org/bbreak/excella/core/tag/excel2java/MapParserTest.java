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

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * MapParserテストクラス
 * 
 * @since 1.0
 */
public class MapParserTest extends WorkbookTest {

    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public final void testMapParser( String version) throws ParseException, IOException {
        Workbook wk = getWorkbook( version);
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        MapParser mapParser = new MapParser( "@Map");
        Object data = null;

        // No.1 パラメータ無し
        Cell tagCell1 = sheet1.getRow( 3).getCell( 0);
        Map<?, ?> map1 = mapParser.parse( sheet1, tagCell1, data);
        Set<?> keySet = map1.keySet();
        assertEquals( 25, keySet.size());
        assertEquals( "値1", map1.get( "キー1"));
        assertEquals( "値2", map1.get( "キー2"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=5}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=5}"));
        assertEquals( "値3", map1.get( "キー3"));
        assertEquals( "値4", map1.get( "キー4"));
        assertEquals( "値5", map1.get( "キー5"));
        assertEquals( "値6", map1.get( "キー6"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,Key=キー}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=3,Key=キー}"));
        assertEquals( "値7", map1.get( "キー7"));
        assertEquals( "値8", map1.get( "キー8"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,Value=値}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=3,Value=値}"));
        assertEquals( "値9", map1.get( "キー9"));
        assertEquals( "値10", map1.get( "キー10"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,KeyColumn=2}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=3,KeyColumn=2}"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=3,ValueColumn=2}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=3,ValueColumn=2}"));
        assertEquals( "備考", map1.get( "キー13"));
        assertEquals( "備考", map1.get( "キー14"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=4,KeyCell=2:2}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=4,KeyCell=2:2}"));
        assertEquals( "値17", map1.get( "備考"));
        assertTrue( keySet.contains( "@Map{DataRowFrom=2,DataRowTo=4,ValueCell=4:3}"));
        assertEquals( null, map1.get( "@Map{DataRowFrom=2,DataRowTo=4,ValueCell=4:3}"));
        assertEquals( "備考", map1.get( "キー18"));
        assertEquals( "備考", map1.get( "キー19"));
        assertEquals( "備考", map1.get( "キー20"));

        // No.2 パラメータ有り
        Cell tagCell2 = sheet1.getRow( 7).getCell( 0);
        Map<?, ?> map2 = mapParser.parse( sheet1, tagCell2, data);
        keySet = map2.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値3", map2.get( "キー3"));
        assertEquals( "値4", map2.get( "キー4"));
        assertEquals( "値5", map2.get( "キー5"));
        assertEquals( "値6", map2.get( "キー6"));

        // No.3 キー指定
        Cell tagCell3 = sheet1.getRow( 15).getCell( 0);
        Map<?, ?> map3 = mapParser.parse( sheet1, tagCell3, data);
        keySet = map3.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "値8", map3.get( "キー"));

        // No.4 値指定
        Cell tagCell4 = sheet1.getRow( 24).getCell( 0);
        Map<?, ?> map4 = mapParser.parse( sheet1, tagCell4, data);
        keySet = map4.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値", map4.get( "キー9"));
        assertEquals( "値", map4.get( "キー10"));

        // No.5 キー列指定
        Cell tagCell5 = sheet1.getRow( 32).getCell( 0);
        Map<?, ?> map5 = mapParser.parse( sheet1, tagCell5, data);
        keySet = map5.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値11", map5.get( "キー11"));
        assertEquals( "値12", map5.get( "キー12"));

        // No.6 値列指定
        Cell tagCell6 = sheet1.getRow( 41).getCell( 0);
        Map<?, ?> map6 = mapParser.parse( sheet1, tagCell6, data);
        keySet = map6.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値13", map6.get( "キー13"));
        assertEquals( "値14", map6.get( "キー14"));

        // No.7 キーセル指定
        Cell tagCell7 = sheet1.getRow( 49).getCell( 0);
        Map<?, ?> map7 = mapParser.parse( sheet1, tagCell7, data);
        keySet = map7.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "値17", map7.get( "キー15"));

        // No.8 値セル指定
        Cell tagCell8 = sheet1.getRow( 58).getCell( 0);
        Map<?, ?> map8 = mapParser.parse( sheet1, tagCell8, data);
        keySet = map8.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値20", map8.get( "キー18"));
        assertEquals( "値20", map8.get( "キー19"));
        assertEquals( "値20", map8.get( "キー20"));

        // No.9 DataRowFrom > DataRowTo
        Cell tagCell9 = sheet1.getRow( 3).getCell( 7);
        ParseException pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet1, tagCell9, data));
        Cell cell = pe.getCell();
        assertEquals( 3, cell.getRow().getRowNum());
        assertEquals( 7, cell.getColumnIndex());
        System.out.println( "No.9:" + pe);

        // No.10 マイナス範囲指定
        Cell tagCell10 = sheet1.getRow( 18).getCell( 7);
        Map<?, ?> map10 = mapParser.parse( sheet1, tagCell10, data);
        keySet = map10.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値5", map10.get( "キー5"));
        assertEquals( "値6", map10.get( "キー6"));
        assertEquals( "値7", map10.get( "キー7"));
        assertEquals( "値8", map10.get( "キー8"));

        // No.11 キー列マイナス指定
        Cell tagCell11 = sheet1.getRow( 23).getCell( 8);
        Map<?, ?> map11 = mapParser.parse( sheet1, tagCell11, data);
        keySet = map11.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値9", map11.get( "キー9"));
        assertEquals( "値10", map11.get( "キー10"));
        assertEquals( "値11", map11.get( "キー11"));
        assertEquals( "値12", map11.get( "キー12"));

        // No.12 値列マイナス指定
        Cell tagCell12 = sheet1.getRow( 33).getCell( 8);
        Map<?, ?> map12 = mapParser.parse( sheet1, tagCell12, data);
        keySet = map12.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値13", map12.get( "キー13"));
        assertEquals( "値14", map12.get( "キー14"));
        assertEquals( "値15", map12.get( "キー15"));
        assertEquals( "値16", map12.get( "キー16"));

        // No.13 キーセルマイナス指定
        Cell tagCell13 = sheet1.getRow( 43).getCell( 8);
        Map<?, ?> map13 = mapParser.parse( sheet1, tagCell13, data);
        keySet = map13.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "値20", map13.get( "キー"));

        // No.14 値セルマイナス指定
        Cell tagCell14 = sheet1.getRow( 53).getCell( 8);
        Map<?, ?> map14 = mapParser.parse( sheet1, tagCell14, data);
        keySet = map14.keySet();
        assertEquals( 4, keySet.size());
        assertEquals( "値", map14.get( "キー21"));
        assertEquals( "値", map14.get( "キー22"));
        assertEquals( "値", map14.get( "キー23"));
        assertEquals( "値", map14.get( "キー24"));

        // No.15 キー・値を同列に指定
        Cell tagCell15 = sheet1.getRow( 63).getCell( 7);
        Map<?, ?> map15 = mapParser.parse( sheet1, tagCell15, data);
        keySet = map15.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "キー・値25", map15.get( "キー・値25"));
        assertEquals( "キー・値26", map15.get( "キー・値26"));

        // No.16 キー・値にタグと同行・同列を指定
        Cell tagCell16 = sheet1.getRow( 71).getCell( 7);
        Map<?, ?> map16 = mapParser.parse( sheet1, tagCell16, data);
        keySet = map16.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "@Map{DataRowFrom=0,DataRowTo=0,ValueColumn=0}", map16.get( "@Map{DataRowFrom=0,DataRowTo=0,ValueColumn=0}"));

        // No.17 キー・値セルにタグセルを指定
        Cell tagCell17 = sheet1.getRow( 76).getCell( 7);
        Map<?, ?> map17 = mapParser.parse( sheet1, tagCell17, data);
        keySet = map17.keySet();
        assertEquals( 1, keySet.size());
        assertEquals( "@Map{DataRowFrom=0,DataRowTo=0,KeyCell=0:0,ValueCell=0:0}", map17.get( "@Map{DataRowFrom=0,DataRowTo=0,KeyCell=0:0,ValueCell=0:0}"));

        // No.18 キーセルがnull
        Cell tagCell18 = sheet2.getRow( 3).getCell( 0);
        Map<?, ?> map18 = mapParser.parse( sheet2, tagCell18, data);
        keySet = map18.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値1", map18.get( "キー1"));
        assertTrue( keySet.contains( null));
        assertEquals( "値2", map18.get( null));
        assertEquals( "値3", map18.get( "キー2"));

        // No.19 値セルがnull
        Cell tagCell19 = sheet2.getRow( 11).getCell( 0);
        Map<?, ?> map19 = mapParser.parse( sheet2, tagCell19, data);
        keySet = map19.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値4", map19.get( "キー3"));
        assertEquals( null, map19.get( "キー4"));
        assertEquals( "値5", map19.get( "キー5"));

        // No.20 データ行がnull
        Cell tagCell20 = sheet2.getRow( 19).getCell( 0);
        Map<?, ?> map20 = mapParser.parse( sheet2, tagCell20, data);
        keySet = map20.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( "値6", map20.get( "キー6"));
        assertEquals( "値7", map20.get( "キー7"));

        // No.21 キーセル値がnull
        Cell tagCell21 = sheet2.getRow( 27).getCell( 0);
        Map<?, ?> map21 = mapParser.parse( sheet2, tagCell21, data);
        keySet = map21.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値8", map21.get( "キー8"));
        assertEquals( "値9", map21.get( null));
        assertEquals( "値10", map21.get( "キー9"));

        // No.22 値セルの値がnull
        Cell tagCell22 = sheet2.getRow( 35).getCell( 0);
        Map<?, ?> map22 = mapParser.parse( sheet2, tagCell22, data);
        keySet = map22.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値11", map22.get( "キー10"));
        assertTrue( keySet.contains( "キー11"));
        assertEquals( null, map22.get( "キー11"));
        assertEquals( "値12", map22.get( "キー12"));

        // No.23 キーセルの値・値セルの値がnull
        Cell tagCell23 = sheet2.getRow( 43).getCell( 0);
        Map<?, ?> map23 = mapParser.parse( sheet2, tagCell23, data);
        keySet = map23.keySet();
        assertEquals( 3, keySet.size());
        assertEquals( "値13", map23.get( "キー13"));
        assertTrue( keySet.contains( null));
        assertEquals( null, map23.get( null));
        assertEquals( "値14", map23.get( "キー14"));

        // No.24 値セルの行がnull
        Cell tagCell24 = sheet2.getRow( 51).getCell( 0);
        Map<?, ?> map24 = mapParser.parse( sheet2, tagCell24, data);
        keySet = map24.keySet();
        assertEquals( 2, keySet.size());
        assertEquals( null, map24.get( "キー15"));
        assertEquals( null, map24.get( "キー16"));

        // No.25 重複定義：Key, KeyColumn, KeyCell
        Cell tagCell25 = sheet3.getRow( 4).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell25, data));
        cell = pe.getCell();
        assertEquals( 4, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.25:" + pe);

        // No.26 重複定義：Key, KeyColumn
        Cell tagCell26 = sheet3.getRow( 8).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell26, data));
        cell = pe.getCell();
        assertEquals( 8, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.26:" + pe);

        // No.27 重複定義：Key, KeyCell
        Cell tagCell27 = sheet3.getRow( 12).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell27, data));
        cell = pe.getCell();
        assertEquals( 12, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.27:" + pe);

        // No.28 重複定義：KeyColumn, KeyCell
        Cell tagCell28 = sheet3.getRow( 16).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet1, tagCell28, data));
        cell = pe.getCell();
        assertEquals( 16, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.28:" + pe);

        // No.29 重複定義：Value, ValueColumn, ValueCell
        Cell tagCell29 = sheet3.getRow( 20).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell29, data));
        cell = pe.getCell();
        assertEquals( 20, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.29:" + pe);

        // No.30 重複定義：Value, ValueColumn
        Cell tagCell30 = sheet3.getRow( 24).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell30, data));
        cell = pe.getCell();
        assertEquals( 24, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.30:" + pe);

        // No.31 重複定義：Value, ValueCell
        Cell tagCell31 = sheet3.getRow( 28).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell31, data));
        cell = pe.getCell();
        assertEquals( 28, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.31:" + pe);

        // No.32 重複定義：ValueColumn, ValueCell
        Cell tagCell32 = sheet3.getRow( 32).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell32, data));
        cell = pe.getCell();
        assertEquals( 32, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.32:" + pe);

        // No.33 データ無し
        Cell tagCell33 = sheet3.getRow( 38).getCell( 1);
        assertDoesNotThrow( () -> mapParser.parse( sheet3, tagCell33, data));

        // No.34 DataRowFrom不正（値が空）
        Cell tagCell34 = sheet3.getRow( 41).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell34, data));
        cell = pe.getCell();
        assertEquals( 41, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.33:" + pe);

        // No.35 DataRowTo不正（値が空）
        Cell tagCell35 = sheet3.getRow( 44).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell35, data));
        cell = pe.getCell();
        assertEquals( 44, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.35:" + pe);

        // No.36 Key不正（値が空）
        Cell tagCell36 = sheet3.getRow( 47).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell36, data));
        cell = pe.getCell();
        assertEquals( 47, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.36:" + pe);

        // No.37 Value不正（値が空）
        Cell tagCell37 = sheet3.getRow( 50).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell37, data));
        cell = pe.getCell();
        assertEquals( 50, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.37:" + pe);

        // No.38 KeyColumn不正（値が空）
        Cell tagCell38 = sheet3.getRow( 53).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell38, data));
        cell = pe.getCell();
        assertEquals( 53, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.38:" + pe);

        // No.39 ValueColumn不正（値が空）
        Cell tagCell39 = sheet3.getRow( 56).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell39, data));
        cell = pe.getCell();
        assertEquals( 56, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.39:" + pe);

        // No.40 KeyCell不正（値が空）
        Cell tagCell40 = sheet3.getRow( 59).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell40, data));
        cell = pe.getCell();
        assertEquals( 59, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.40:" + pe);

        // No.41 ValueCell不正（値が空）
        Cell tagCell41 = sheet3.getRow( 62).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell41, data));
        cell = pe.getCell();
        assertEquals( 62, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.41:" + pe);

        // No.42 DataRowFrom不正（値が文字）
        Cell tagCell42 = sheet3.getRow( 65).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell42, data));
        cell = pe.getCell();
        assertEquals( 65, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.42:" + pe);

        // No.43 DataRowTo不正（値が文字）
        Cell tagCell43 = sheet3.getRow( 68).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell43, data));
        cell = pe.getCell();
        assertEquals( 68, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.43:" + pe);

        // No.44 KeyColumen不正（値が文字）
        Cell tagCell44 = sheet3.getRow( 71).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell44, data));
        cell = pe.getCell();
        assertEquals( 71, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.44:" + pe);

        // No.45 ValueColumn不正（値が文字）
        Cell tagCell45 = sheet3.getRow( 74).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell45, data));
        cell = pe.getCell();
        assertEquals( 74, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.45:" + pe);

        // No.46 KeyCell不正（値が文字）
        Cell tagCell46 = sheet3.getRow( 77).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell46, data));
        cell = pe.getCell();
        assertEquals( 77, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.46:" + pe);

        // No.47 ValueCell不正（値が文字）
        Cell tagCell47 = sheet3.getRow( 80).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet3, tagCell47, data));
        cell = pe.getCell();
        assertEquals( 80, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.47:" + pe);

        // No.48 KeyColumn不正テスト（A列でキー列にマイナスを指定）
        Cell tagCell48 = sheet4.getRow( 2).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell48, data));
        cell = pe.getCell();
        assertEquals( 2, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.48:" + pe);

        // No.49 ValueColumn不正テスト（A列で値列にマイナスを指定）
        Cell tagCell49 = sheet4.getRow( 11).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell49, data));
        cell = pe.getCell();
        assertEquals( 11, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.49:" + pe);

        // No.50 KeyCell不正テスト（A列でキーセル列にマイナスを指定）
        Cell tagCell50 = sheet4.getRow( 20).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell50, data));
        cell = pe.getCell();
        assertEquals( 20, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.50:" + pe);

        // No.51 ValueCell不正テスト（A列で値セル列にマイナスを指定）
        Cell tagCell51 = sheet4.getRow( 29).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell51, data));
        cell = pe.getCell();
        assertEquals( 29, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.51:" + pe);

        // No.52 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        Cell tagCell52 = sheet4.getRow( 0).getCell( 3);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell52, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 3, cell.getColumnIndex());
        System.out.println( "No.52:" + pe);

        // No.53 DataRowTo不正テスト（1行目でデータ終了行にマイナスを指定）
        Cell tagCell53 = sheet4.getRow( 0).getCell( 6);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell53, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 6, cell.getColumnIndex());
        System.out.println( "No.53:" + pe);

        // No.54 KeyCell不正テスト（1行目でキーセル行にマイナスを指定）
        Cell tagCell54 = sheet4.getRow( 0).getCell( 9);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell54, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 9, cell.getColumnIndex());
        System.out.println( "No.54:" + pe);

        // No.55 ValueCell不正テスト（1行目で値セル行にマイナスを指定）
        Cell tagCell55 = sheet4.getRow( 0).getCell( 12);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell55, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 12, cell.getColumnIndex());
        System.out.println( "No.55:" + pe);

        // No.56 KeyColumn不正テスト（最終列でキー列にプラスを設定）
        Cell tagCell56 = sheet4.getRow( 3).getCell( 16);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell56, data));
        cell = pe.getCell();
        assertEquals( 3, cell.getRow().getRowNum());
        assertEquals( 16, cell.getColumnIndex());
        System.out.println( "No.56:" + pe);

        // No.57 ValueColumn不正テスト（最終列で値列にプラスを設定）
        Cell tagCell57 = sheet4.getRow( 12).getCell( 16);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell57, data));
        cell = pe.getCell();
        assertEquals( 12, cell.getRow().getRowNum());
        assertEquals( 16, cell.getColumnIndex());
        System.out.println( "No.57:" + pe);

        // No.58 KeyCell不正テスト（最終列でキーセル列にプラスを設定）
        Cell tagCell58 = sheet4.getRow( 21).getCell( 16);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell58, data));
        cell = pe.getCell();
        assertEquals( 21, cell.getRow().getRowNum());
        assertEquals( 16, cell.getColumnIndex());
        System.out.println( "No.58:" + pe);

        // No.59 ValueCell不正テスト（最終列で値セル列にプラスを設定）
        Cell tagCell59 = sheet4.getRow( 30).getCell( 16);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell59, data));
        cell = pe.getCell();
        assertEquals( 30, cell.getRow().getRowNum());
        assertEquals( 16, cell.getColumnIndex());
        System.out.println( "No.59:" + pe);

        // No.60 DataRowFrom不正テスト（最終行でデータ開始行にプラスを設定）
        Cell tagCell60 = sheet4.getRow( 34).getCell( 3);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell60, data));
        cell = pe.getCell();
        assertEquals( 34, cell.getRow().getRowNum());
        assertEquals( 3, cell.getColumnIndex());
        System.out.println( "No.60:" + pe);

        // No.61 DataRowTo不正テスト（最終行でデータ終了行にプラスを設定）
        Cell tagCell61 = sheet4.getRow( 34).getCell( 6);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell61, data));
        cell = pe.getCell();
        assertEquals( 34, cell.getRow().getRowNum());
        assertEquals( 6, cell.getColumnIndex());
        System.out.println( "No.61:" + pe);

        // No.62 KeyCell不正テスト（最終行でキーセル行にプラスを設定）
        Cell tagCell62 = sheet4.getRow( 34).getCell( 9);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell62, data));
        cell = pe.getCell();
        assertEquals( 34, cell.getRow().getRowNum());
        assertEquals( 9, cell.getColumnIndex());
        System.out.println( "No.62:" + pe);

        // No.63 ValueCell不正テスト（最終行で値セル行にプラスを設定）
        Cell tagCell63 = sheet4.getRow( 34).getCell( 12);
        pe = assertThrows( ParseException.class, () -> mapParser.parse( sheet4, tagCell63, data));
        cell = pe.getCell();
        assertEquals( 34, cell.getRow().getRowNum());
        assertEquals( 12, cell.getColumnIndex());
        System.out.println( "No.63:" + pe);
    }
}
