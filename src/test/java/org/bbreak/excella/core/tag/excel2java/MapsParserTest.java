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

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * MapsParserテストクラス
 * 
 * @since 1.0
 */
public class MapsParserTest extends WorkbookTest {

    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public final void testMapsParser( String version) throws ParseException, IOException {
        Workbook wk = getWorkbook( version);
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        Sheet sheet5 = wk.getSheetAt( 4);
        MapsParser mapsParser = new MapsParser( "@Maps");
        Object data = null;

        // No.1 パラメータ無し
        Cell tagCell1 = sheet1.getRow( 5).getCell( 0);
        List<Map<?, ?>> maps1 = mapsParser.parse( sheet1, tagCell1, data);
        assertEquals( "value1-1", maps1.get( 0).get( "key1"));
        assertEquals( "value2-1", maps1.get( 0).get( "key2"));
        assertEquals( 2, maps1.get( 0).keySet().size());
        assertEquals( "value1-2", maps1.get( 1).get( "key1"));
        assertEquals( "value2-2", maps1.get( 1).get( "key2"));
        assertEquals( 2, maps1.get( 1).keySet().size());
        assertEquals( "value1-3", maps1.get( 2).get( "key1"));
        assertEquals( "value2-3", maps1.get( 2).get( "key2"));
        assertEquals( 2, maps1.get( 2).keySet().size());
        assertEquals( 3, maps1.size());

        // No.2 データ無し
        Cell tagCell2 = sheet2.getRow( 5).getCell( 0);
        List<Map<?, ?>> maps2 = mapsParser.parse( sheet2, tagCell2, data);
        assertEquals( 0, maps2.size());

        // No.3 パラメータ有り
        Cell tagCell3 = sheet3.getRow( 5).getCell( 0);
        List<Map<?, ?>> maps3 = mapsParser.parse( sheet3, tagCell3, data);
        assertEquals( "value1-1", maps3.get( 0).get( "key1"));
        assertEquals( "value2-1", maps3.get( 0).get( "key2"));
        assertEquals( 2, maps3.get( 0).keySet().size());
        assertEquals( "value1-2", maps3.get( 1).get( "key1"));
        assertEquals( "value2-2", maps3.get( 1).get( "key2"));
        assertEquals( 2, maps3.get( 1).keySet().size());
        assertEquals( 2, maps3.size());

        // No.4 キー列、開始列、終了列指定
        Cell tagCell4 = sheet3.getRow( 13).getCell( 0);
        List<Map<?, ?>> maps4 = mapsParser.parse( sheet3, tagCell4, data);
        assertEquals( "value3-1", maps4.get( 0).get( "key3"));
        assertEquals( "value4-1", maps4.get( 0).get( "key4"));
        assertEquals( 2, maps4.get( 0).keySet().size());
        assertEquals( "value3-2", maps4.get( 1).get( "key3"));
        assertEquals( "value4-2", maps4.get( 1).get( "key4"));
        assertEquals( 2, maps4.get( 1).keySet().size());
        assertEquals( 2, maps4.size());

        // No.5 コメント有り
        Cell tagCell5 = sheet3.getRow( 23).getCell( 0);
        List<Map<?, ?>> maps5 = mapsParser.parse( sheet3, tagCell5, data);
        assertEquals( "value5-1", maps5.get( 0).get( "key5"));
        assertEquals( "value6-1", maps5.get( 0).get( "key6"));
        assertEquals( null, maps5.get( 2).get( "key7"));
        assertEquals( "value8-1", maps5.get( 0).get( "key8"));
        assertEquals( 3, maps5.get( 0).keySet().size());
        assertEquals( "value5-2", maps5.get( 1).get( "key5"));
        assertEquals( "value6-2", maps5.get( 1).get( "key6"));
        assertEquals( null, maps5.get( 2).get( "key7"));
        assertEquals( "value8-2", maps5.get( 1).get( "key8"));
        assertEquals( 3, maps5.get( 1).keySet().size());
        assertEquals( "value5-3", maps5.get( 2).get( "key5"));
        assertEquals( "value6-3", maps5.get( 2).get( "key6"));
        assertEquals( null, maps5.get( 2).get( "key7"));
        assertEquals( "value8-3", maps5.get( 2).get( "key8"));
        assertEquals( 3, maps5.get( 2).keySet().size());
        assertEquals( 3, maps5.size());

        // No.6 キー行がnull
        Cell tagCell6 = sheet3.getRow( 31).getCell( 0);
        List<Map<?, ?>> maps6 = mapsParser.parse( sheet3, tagCell6, data);
        assertEquals( 0, maps6.size());

        // No.7 キーセルがnull
        Cell tagCell7 = sheet3.getRow( 39).getCell( 0);
        List<Map<?, ?>> maps7 = mapsParser.parse( sheet3, tagCell7, data);
        assertEquals( "value9-1", maps7.get( 0).get( "key9"));
        assertEquals( "value10-1", maps7.get( 0).get( "key10"));
        assertEquals( 2, maps7.get( 0).keySet().size());
        assertEquals( "value9-2", maps7.get( 1).get( "key9"));
        assertEquals( "value10-2", maps7.get( 1).get( "key10"));
        assertEquals( 2, maps7.get( 1).keySet().size());
        assertEquals( "value9-3", maps7.get( 2).get( "key9"));
        assertEquals( "value10-3", maps7.get( 2).get( "key10"));
        assertEquals( 2, maps7.get( 2).keySet().size());
        assertEquals( 3, maps7.size());

        // No.8 キーセル値がnull
        Cell tagCell8 = sheet3.getRow( 47).getCell( 0);
        List<Map<?, ?>> maps8 = mapsParser.parse( sheet3, tagCell8, data);
        assertEquals( "value11-1", maps8.get( 0).get( "key11"));
        assertEquals( "value12-1", maps8.get( 0).get( "key12"));
        assertEquals( 2, maps8.get( 0).keySet().size());
        assertEquals( "value11-2", maps8.get( 1).get( "key11"));
        assertEquals( "value12-2", maps8.get( 1).get( "key12"));
        assertEquals( 2, maps8.get( 1).keySet().size());
        assertEquals( "value11-3", maps8.get( 2).get( "key11"));
        assertEquals( "value12-3", maps8.get( 2).get( "key12"));
        assertEquals( 2, maps8.get( 2).keySet().size());
        assertEquals( 3, maps8.size());

        // No.9 値行がnull
        Cell tagCell9 = sheet3.getRow( 55).getCell( 0);
        List<Map<?, ?>> maps9 = mapsParser.parse( sheet3, tagCell9, data);
        assertEquals( "value13-1", maps9.get( 0).get( "key13"));
        assertEquals( "value14-1", maps9.get( 0).get( "key14"));
        assertEquals( 2, maps9.get( 0).keySet().size());
        assertEquals( "value13-2", maps9.get( 1).get( "key13"));
        assertEquals( "value14-2", maps9.get( 1).get( "key14"));
        assertEquals( 2, maps9.get( 1).keySet().size());
        assertEquals( 2, maps9.size());

        // No.10 値セルがnull
        Cell tagCell10 = sheet3.getRow( 63).getCell( 0);
        List<Map<?, ?>> maps10 = mapsParser.parse( sheet3, tagCell10, data);
        assertEquals( "value15-1", maps10.get( 0).get( "key15"));
        assertEquals( "value16-1", maps10.get( 0).get( "key16"));
        assertEquals( 2, maps10.get( 0).keySet().size());
        assertEquals( null, maps10.get( 1).get( "key15"));
        assertEquals( null, maps10.get( 1).get( "key16"));
        assertEquals( 2, maps10.get( 1).keySet().size());
        assertEquals( "value15-2", maps10.get( 2).get( "key15"));
        assertEquals( "value16-2", maps10.get( 2).get( "key16"));
        assertEquals( 2, maps10.get( 2).keySet().size());
        assertEquals( 3, maps10.size());

        // No.11 値セル値がnull
        Cell tagCell11 = sheet3.getRow( 71).getCell( 0);
        List<Map<?, ?>> maps11 = mapsParser.parse( sheet3, tagCell11, data);
        assertEquals( "value17-1", maps11.get( 0).get( "key17"));
        assertEquals( "value18-1", maps11.get( 0).get( "key18"));
        assertEquals( 2, maps11.get( 0).keySet().size());
        assertEquals( null, maps11.get( 1).get( "key17"));
        assertEquals( null, maps11.get( 1).get( "key18"));
        assertEquals( 2, maps11.get( 1).keySet().size());
        assertEquals( "value17-2", maps11.get( 2).get( "key17"));
        assertEquals( "value18-2", maps11.get( 2).get( "key18"));
        assertEquals( 2, maps11.get( 2).keySet().size());
        assertEquals( 3, maps11.size());

        // No.12 重複キー有り
        Cell tagCell12 = sheet3.getRow( 79).getCell( 0);
        List<Map<?, ?>> maps12 = mapsParser.parse( sheet3, tagCell12, data);
        assertEquals( "value19-4", maps12.get( 0).get( "key19"));
        assertEquals( 1, maps12.get( 0).keySet().size());
        assertEquals( null, maps12.get( 1).get( "key19"));
        assertEquals( 1, maps12.get( 1).keySet().size());
        assertEquals( "value19-5", maps12.get( 2).get( "key19"));
        assertEquals( 1, maps12.get( 2).keySet().size());
        assertEquals( 3, maps12.size());

        // No.13 マイナス範囲指定
        Cell tagCell13 = sheet4.getRow( 7).getCell( 0);
        List<Map<?, ?>> maps13 = mapsParser.parse( sheet4, tagCell13, data);
        assertEquals( "value1-1", maps13.get( 0).get( "key1"));
        assertEquals( "value2-1", maps13.get( 0).get( "key2"));
        assertEquals( 2, maps13.get( 0).keySet().size());
        assertEquals( "value1-2", maps13.get( 1).get( "key1"));
        assertEquals( "value2-2", maps13.get( 1).get( "key2"));
        assertEquals( 2, maps13.get( 1).keySet().size());
        assertEquals( "value1-3", maps13.get( 2).get( "key1"));
        assertEquals( "value2-3", maps13.get( 2).get( "key2"));
        assertEquals( 2, maps13.get( 2).keySet().size());
        assertEquals( "value1-4", maps13.get( 3).get( "key1"));
        assertEquals( "value2-4", maps13.get( 3).get( "key2"));
        assertEquals( 2, maps13.get( 3).keySet().size());
        assertEquals( 4, maps13.size());

        // No.14 キー・データ同列
        Cell tagCell14 = sheet4.getRow( 12).getCell( 0);
        List<Map<?, ?>> maps14 = mapsParser.parse( sheet4, tagCell14, data);
        assertEquals( "キー・値1", maps14.get( 0).get( "キー・値1"));
        assertEquals( "キー・値2", maps14.get( 0).get( "キー・値2"));
        assertEquals( 2, maps14.get( 0).keySet().size());
        assertEquals( 1, maps14.size());

        // No.15 DataRowFrom > DataRowTo
        Cell tagCell15 = sheet4.getRow( 21).getCell( 0);
        ParseException pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell15, data));
        Cell cell = pe.getCell();
        assertEquals( 21, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.15:" + pe);

        // No.16 KeyRow不正（値が空）
        Cell tagCell16 = sheet4.getRow( 27).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell16, data));
        cell = pe.getCell();
        assertEquals( 27, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.16:" + pe);

        // No.17 DataRowFrom不正（値が空）
        Cell tagCell17 = sheet4.getRow( 30).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell17, data));
        cell = pe.getCell();
        assertEquals( 30, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.17:" + pe);

        // No.18 DataRowTo不正（値が空）
        Cell tagCell18 = sheet4.getRow( 33).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell18, data));
        cell = pe.getCell();
        assertEquals( 33, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.18:" + pe);

        // No.19 KeyRow不正（値が文字）
        Cell tagCell19 = sheet4.getRow( 36).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell19, data));
        cell = pe.getCell();
        assertEquals( 36, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.19:" + pe);

        // No.20 DataRowFrom不正（値が文字）
        Cell tagCell20 = sheet4.getRow( 39).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell20, data));
        cell = pe.getCell();
        assertEquals( 39, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.20:" + pe);

        // No.21 DataRowFrom不正（値が文字）
        Cell tagCell21 = sheet4.getRow( 42).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet4, tagCell21, data));
        cell = pe.getCell();
        assertEquals( 42, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.21:" + pe);

        // No.22 KeyRow不正テスト（1行目でキー行にマイナスを指定）
        Cell tagCell22 = sheet5.getRow( 0).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet5, tagCell22, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.22:" + pe);

        // No.23 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        Cell tagCell23 = sheet5.getRow( 0).getCell( 3);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet5, tagCell23, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 3, cell.getColumnIndex());
        System.out.println( "No.23:" + pe);

        // No.24 DataRowTo不正テスト（1行目でデータ終了行にマイナスを指定）
        Cell tagCell24 = sheet5.getRow( 0).getCell( 6);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet5, tagCell24, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 6, cell.getColumnIndex());
        System.out.println( "No.24:" + pe);

        // No.25 KeyRow不正テスト（最終行でキー行にプラスを指定）
        Cell tagCell25 = sheet5.getRow( 18).getCell( 0);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet5, tagCell25, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.25:" + pe);

        // No.26 DataRowFrom不正テスト（最終行でデータ開始行にプラスを指定）
        Cell tagCell26 = sheet5.getRow( 18).getCell( 3);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet5, tagCell26, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 3, cell.getColumnIndex());
        System.out.println( "No.26:" + pe);

        // No.27 DataRowTo不正テスト（最終行でデータ終了行にプラスを指定）
        Cell tagCell27 = sheet5.getRow( 18).getCell( 6);
        pe = assertThrows( ParseException.class, () -> mapsParser.parse( sheet5, tagCell27, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 6, cell.getColumnIndex());
        System.out.println( "No.27:" + pe);
    }

}
