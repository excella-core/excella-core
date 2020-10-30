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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * ListParserテストクラス
 * 
 * @since 1.0
 */
public class ListParserTest extends WorkbookTest {

    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public final void testListParser( String version) throws ParseException, IOException {
        Workbook wk = getWorkbook( version);
        Sheet sheet1 = wk.getSheetAt( 0);
        ListParser listParser = new ListParser( "@List");
        Object data = null;

        // No.1 @List
        Cell tagCell1 = sheet1.getRow( 2).getCell( 0);
        List<?> list1 = listParser.parse( sheet1, tagCell1, data);
        assertEquals( "要素1", list1.get( 0));
        assertEquals( "要素2", list1.get( 1));
        assertEquals( "@List{DataRowFrom=2,DataRowTo=5}", list1.get( 2));
        assertEquals( "リスト", list1.get( 3));
        assertEquals( "要素3", list1.get( 4));
        assertEquals( "要素4", list1.get( 5));
        assertEquals( "要素5", list1.get( 6));
        assertEquals( "@List{DataRowFrom=2,DataRowTo=3,ValueColumn=1}", list1.get( 7));
        assertEquals( null, list1.get( 8));
        assertEquals( null, list1.get( 9));
        assertEquals( null, list1.get( 10));
        assertEquals( "@List{ValueColumn=-1}", list1.get( 11));
        assertEquals( "要素8", list1.get( 12));
        assertEquals( 13, list1.size());

        // No.2 @List{DataRowFrom=2,DataRowTo=5}
        Cell tagCell2 = sheet1.getRow( 6).getCell( 0);
        List<?> list2 = listParser.parse( sheet1, tagCell2, data);
        assertEquals( "要素3", list2.get( 0));
        assertEquals( "要素4", list2.get( 1));
        assertEquals( "要素5", list2.get( 2));
        assertEquals( 3, list2.size());

        // No.3 @List{DataRowFrom=2,DataRowTo=3,ValueColumn=1}
        Cell tagCell3 = sheet1.getRow( 14).getCell( 0);
        List<?> list3 = listParser.parse( sheet1, tagCell3, data);
        assertEquals( "要素6", list3.get( 0));
        assertEquals( "要素7", list3.get( 1));
        assertEquals( 2, list3.size());

        // No.4 @List{ValueColumn=-1}(1列目に書かれていた場合)
        Cell tagCell4 = sheet1.getRow( 24).getCell( 0);
        ParseException pe = assertThrows( ParseException.class, () -> listParser.parse( sheet1, tagCell4, data));
        Cell cell = pe.getCell();
        assertEquals( 24, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.4:" + pe);

        // No.5 @List{DataRowFrom=-1}(1行目に書かれている場合)
        Cell tagCell5 = sheet1.getRow( 0).getCell( 4);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet1, tagCell5, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 4, cell.getColumnIndex());
        System.out.println( "No.5:" + pe);

        // No.6 @List{DataRowFrom=1}(指定シートの最終行に書かれていた場合)
        Cell tagCell6 = sheet1.getRow( 25).getCell( 4);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet1, tagCell6, data));
        cell = pe.getCell();
        assertEquals( 25, cell.getRow().getRowNum());
        assertEquals( 4, cell.getColumnIndex());
        System.out.println( "No.6:" + pe);

        // No.7 @List{DataRowTo=-1}(1行目に書かれている場合)
        Cell tagCell7 = sheet1.getRow( 0).getCell( 7);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet1, tagCell7, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 7, cell.getColumnIndex());
        System.out.println( "No.7:" + pe);

        // No.8 @List{DataRowTo=1}(指定シートの最終行に書かれていた場合)
        Cell tagCell8 = sheet1.getRow( 25).getCell( 7);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet1, tagCell8, data));
        cell = pe.getCell();
        assertEquals( 25, cell.getRow().getRowNum());
        assertEquals( 7, cell.getColumnIndex());
        System.out.println( "No.8:" + pe);

        // No.9 @List{ValueColumn=1}(最終列に書かれていた場合)
        Cell tagCell9 = sheet1.getRow( 2).getCell( 10);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet1, tagCell9, data));
        cell = pe.getCell();
        assertEquals( 2, cell.getRow().getRowNum());
        assertEquals( 10, cell.getColumnIndex());
        System.out.println( "No.9:" + pe);

        // シート変更
        Sheet sheet2 = wk.getSheetAt( 1);

        // No.10 @List{DataRowFrom=3,DataRowTo=1}(DataRowFrom > DataRowTo の場合)
        Cell tagCell10 = sheet2.getRow( 0).getCell( 0);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell10, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.10:" + pe);

        // No.11 @List{DataRowFrom=-1,DataRowTo=-3}(パラメータがマイナス値)
        Cell tagCell11 = sheet2.getRow( 13).getCell( 0);
        List<?> list11 = listParser.parse( sheet2, tagCell11, data);
        assertEquals( "要素2-4", list11.get( 0));
        assertEquals( "要素2-5", list11.get( 1));
        assertEquals( "要素2-6", list11.get( 2));
        assertEquals( 3, list11.size());

        // No.12 @List{ValueColumn=-1}(パラメータがマイナス値)
        Cell tagCell12 = sheet2.getRow( 29).getCell( 1);
        List<?> list12 = listParser.parse( sheet2, tagCell12, data);
        assertEquals( "要素2-7", list12.get( 0));
        assertEquals( "要素2-8", list12.get( 1));
        assertEquals( "要素2-9", list12.get( 2));
        assertEquals( 3, list12.size());

        // No.13 @List{DataRowFrom=1}(DataRowFromのみを設定)
        Cell tagCell13 = sheet2.getRow( 27).getCell( 4);
        List<?> list13 = listParser.parse( sheet2, tagCell13, data);
        assertEquals( "要素2-10", list13.get( 0));
        assertEquals( "要素2-11", list13.get( 1));
        assertEquals( "要素2-12", list13.get( 2));
        assertEquals( null, list13.get( 3));
        assertEquals( "要素2-13", list13.get( 4));
        assertEquals( 5, list13.size());

        // No.14 @List{DataRowTo=2}(DataRowToのみを設定)
        Cell tagCell14 = sheet2.getRow( 27).getCell( 8);
        List<?> list14 = listParser.parse( sheet2, tagCell14, data);
        assertEquals( "要素2-14", list14.get( 0));
        assertEquals( "要素2-15", list14.get( 1));
        assertEquals( 2, list14.size());

        // No.15 @List{DataRowFrom=}(DataRowFrom値空)
        Cell tagCell15 = sheet2.getRow( 0).getCell( 4);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell15, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 4, cell.getColumnIndex());
        System.out.println( "No.15:" + pe);

        // No.16 @List{DataRowTo=}(DataRowTo値空)
        Cell tagCell16 = sheet2.getRow( 9).getCell( 4);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell16, data));
        cell = pe.getCell();
        assertEquals( 9, cell.getRow().getRowNum());
        assertEquals( 4, cell.getColumnIndex());
        System.out.println( "No.16:" + pe);

        // No.17 @List{ValueColumn=}(ValueColumn値空)
        Cell tagCell17 = sheet2.getRow( 18).getCell( 4);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell17, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 4, cell.getColumnIndex());
        System.out.println( "No.17:" + pe);

        // No.18 @List{DataRowFrom=a}(DataRowFrom値文字)
        Cell tagCell18 = sheet2.getRow( 0).getCell( 7);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell18, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 7, cell.getColumnIndex());
        System.out.println( "No.18:" + pe);

        // No.19 @List{DataRowTo=a}(DataRowTo値文字)
        Cell tagCell19 = sheet2.getRow( 9).getCell( 7);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell19, data));
        cell = pe.getCell();
        assertEquals( 9, cell.getRow().getRowNum());
        assertEquals( 7, cell.getColumnIndex());
        System.out.println( "No.19:" + pe);

        // No.20 @List{ValueColumn=a}(ValueColumn値文字)
        Cell tagCell20 = sheet2.getRow( 18).getCell( 7);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell20, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 7, cell.getColumnIndex());
        System.out.println( "No.20:" + pe);

        // No.21 @List(要素なし)
        Cell tagCell21 = sheet2.getRow( 0).getCell( 0);
        pe = assertThrows( ParseException.class, () -> listParser.parse( sheet2, tagCell21, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.21:" + pe);
    }
}
