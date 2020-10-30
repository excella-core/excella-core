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
import java.math.BigDecimal;
import java.text.DateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.excel2java.entity.TargetEntity;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * ObjectsParserテストクラス
 * 
 * @since 1.0
 */
public class ObjectsParserTest extends WorkbookTest {

    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public final void testObjectsParser( String version) throws ParseException, java.text.ParseException, IOException {
        Workbook wk = getWorkbook( version);
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        Sheet sheet5 = wk.getSheetAt( 4);
        ObjectsParser objectsParser = new ObjectsParser( "@Objects");
        Object data = null;

        // No.1 パラメータ無し
        Cell tagCell1 = sheet1.getRow( 5).getCell( 0);
        ParseException pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet1, tagCell1, data));
        Cell cell = pe.getCell();
        assertEquals( 5, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.1:" + pe);

        // No.2 クラス指定
        Cell tagCell2 = sheet1.getRow( 13).getCell( 0);
        List<Object> list2 = objectsParser.parse( sheet1, tagCell2, data);
        TargetEntity targetEntity = ( TargetEntity) list2.get( 0);
        assertEquals( Integer.valueOf( "1"), targetEntity.getNumberInteger());
        assertEquals( 2, targetEntity.getNumberInt());
        assertEquals( Long.valueOf( "3"), targetEntity.getNumberLong());
        assertEquals( 4, targetEntity.getNumberlong());
        assertEquals( Float.valueOf( "5.5"), targetEntity.getNumberFloat());
        assertEquals( "6.6", String.valueOf( targetEntity.getNumberfloat()));
        assertEquals( Double.valueOf( "8.8"), targetEntity.getNumberDouble());
        assertEquals( "9.9", String.valueOf( targetEntity.getNumberdouble()));
        assertEquals( new BigDecimal( "10.1"), targetEntity.getNumberDecimal());
        assertEquals( Byte.valueOf( "11"), targetEntity.getNumberByte());
        assertEquals( 12, targetEntity.getNumberbyte());
        assertEquals( Boolean.TRUE, targetEntity.getValueBoolean());
        assertEquals( Boolean.FALSE, targetEntity.isValueboolean());
        assertEquals( DateFormat.getDateInstance().parse( "2009/4/13"), targetEntity.getDate());
        assertEquals( "文字列", targetEntity.getString());
        assertEquals( 1, list2.size());

        // No.3 キー列、開始列、終了列指定
        Cell tagCell3 = sheet2.getRow( 4).getCell( 0);
        List<Object> list3 = objectsParser.parse( sheet2, tagCell3, data);
        targetEntity = ( TargetEntity) list3.get( 0);
        assertEquals( Integer.valueOf( "1"), targetEntity.getNumberInteger());
        assertEquals( 3, targetEntity.getNumberInt());
        targetEntity = ( TargetEntity) list3.get( 1);
        assertEquals( Integer.valueOf( "2"), targetEntity.getNumberInteger());
        assertEquals( 4, targetEntity.getNumberInt());
        assertEquals( 2, list3.size());

        // No.4 コメント有り
        Cell tagCell4 = sheet2.getRow( 13).getCell( 0);
        List<Object> list4 = objectsParser.parse( sheet2, tagCell4, data);
        targetEntity = ( TargetEntity) list4.get( 0);
        assertEquals( Integer.valueOf( "5"), targetEntity.getNumberInteger());
        assertEquals( 9, targetEntity.getNumberInt());
        targetEntity = ( TargetEntity) list4.get( 1);
        assertEquals( Integer.valueOf( "6"), targetEntity.getNumberInteger());
        assertEquals( 10, targetEntity.getNumberInt());
        assertEquals( 2, list4.size());

        // No.5 キー行がnull
        Cell tagCell5 = sheet2.getRow( 20).getCell( 0);
        List<Object> list5 = objectsParser.parse( sheet2, tagCell5, data);
        assertEquals( 0, list5.size());

        // No.6 キーセルがnull
        Cell tagCell6 = sheet2.getRow( 28).getCell( 0);
        List<Object> list6 = objectsParser.parse( sheet2, tagCell6, data);
        targetEntity = ( TargetEntity) list6.get( 0);
        assertEquals( 20000000, targetEntity.getNumberlong());
        assertEquals( Long.valueOf( "26000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list6.get( 1);
        assertEquals( 21000000, targetEntity.getNumberlong());
        assertEquals( Long.valueOf( "27000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list6.get( 2);
        assertEquals( 22000000, targetEntity.getNumberlong());
        assertEquals( Long.valueOf( "28000000"), targetEntity.getNumberLong());
        assertEquals( 3, list6.size());

        // No.7 キーセル値がnull
        Cell tagCell7 = sheet2.getRow( 36).getCell( 0);
        List<Object> list7 = objectsParser.parse( sheet2, tagCell7, data);
        targetEntity = ( TargetEntity) list7.get( 0);
        assertEquals( 29000000, targetEntity.getNumberlong());
        assertEquals( Long.valueOf( "35000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list7.get( 1);
        assertEquals( 30000000, targetEntity.getNumberlong());
        assertEquals( Long.valueOf( "36000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list7.get( 2);
        assertEquals( 31000000, targetEntity.getNumberlong());
        assertEquals( Long.valueOf( "37000000"), targetEntity.getNumberLong());
        assertEquals( 3, list7.size());

        // No.8 値行がnull
        Cell tagCell8 = sheet2.getRow( 44).getCell( 0);
        List<Object> list8 = objectsParser.parse( sheet2, tagCell8, data);
        targetEntity = ( TargetEntity) list8.get( 0);
        assertEquals( Float.valueOf( "38000000"), targetEntity.getNumberFloat());
        assertEquals( 40000000, targetEntity.getNumberfloat(), 0.01);
        targetEntity = ( TargetEntity) list8.get( 1);
        assertEquals( Float.valueOf( "39000000"), targetEntity.getNumberFloat());
        assertEquals( 41000000, targetEntity.getNumberfloat(), 0.01);
        assertEquals( 2, list8.size());

        // No.9 値セルがnull
        Cell tagCell9 = sheet2.getRow( 52).getCell( 0);
        List<Object> list9 = objectsParser.parse( sheet2, tagCell9, data);
        targetEntity = ( TargetEntity) list9.get( 0);
        assertEquals( Float.valueOf( "42000000"), targetEntity.getNumberFloat());
        assertEquals( Double.valueOf( "44000000"), targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list9.get( 1);
        assertEquals( null, targetEntity.getNumberFloat());
        assertEquals( null, targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list9.get( 2);
        assertEquals( Float.valueOf( "43000000"), targetEntity.getNumberFloat());
        assertEquals( Double.valueOf( "45000000"), targetEntity.getNumberDouble());
        assertEquals( 3, list9.size());

        // No.10 値セル値がnull
        Cell tagCell10 = sheet2.getRow( 60).getCell( 0);
        List<Object> list10 = objectsParser.parse( sheet2, tagCell10, data);
        targetEntity = ( TargetEntity) list10.get( 0);
        assertEquals( Float.valueOf( "46000000"), targetEntity.getNumberFloat());
        assertEquals( Double.valueOf( "48000000"), targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list10.get( 1);
        assertEquals( null, targetEntity.getNumberFloat());
        assertEquals( null, targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list10.get( 2);
        assertEquals( Float.valueOf( "47000000"), targetEntity.getNumberFloat());
        assertEquals( Double.valueOf( "49000000"), targetEntity.getNumberDouble());
        assertEquals( 3, list10.size());

        // No.11 重複キー有り
        Cell tagCell11 = sheet2.getRow( 68).getCell( 0);
        List<Object> list11 = objectsParser.parse( sheet2, tagCell11, data);
        targetEntity = ( TargetEntity) list11.get( 0);
        assertEquals( Double.valueOf( "8.8888888"), targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list11.get( 1);
        assertEquals( null, targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list11.get( 2);
        assertEquals( Double.valueOf( "9.9999999"), targetEntity.getNumberDouble());
        assertEquals( 3, list11.size());

        // No.12 マイナス範囲指定
        Cell tagCell12 = sheet3.getRow( 7).getCell( 0);
        List<Object> list12 = objectsParser.parse( sheet3, tagCell12, data);
        targetEntity = ( TargetEntity) list12.get( 0);
        assertEquals( Boolean.TRUE, targetEntity.getValueBoolean());
        assertEquals( "値1", targetEntity.getString());
        targetEntity = ( TargetEntity) list12.get( 1);
        assertEquals( Boolean.TRUE, targetEntity.getValueBoolean());
        assertEquals( "値2", targetEntity.getString());
        targetEntity = ( TargetEntity) list12.get( 2);
        assertEquals( Boolean.FALSE, targetEntity.getValueBoolean());
        assertEquals( "値3", targetEntity.getString());
        targetEntity = ( TargetEntity) list12.get( 3);
        assertEquals( Boolean.FALSE, targetEntity.getValueBoolean());
        assertEquals( "値4", targetEntity.getString());
        assertEquals( 4, list12.size());

        // No.13 プロパティ列とデータ列を同じ列に指定
        Cell tagCell13 = sheet3.getRow( 12).getCell( 0);
        List<Object> list13 = objectsParser.parse( sheet3, tagCell13, data);
        targetEntity = ( TargetEntity) list13.get( 0);
        assertEquals( "string", targetEntity.getString());
        assertEquals( 1, list13.size());

        // No.14 DataRowFrom > DataRowTo
        Cell tagCell14 = sheet3.getRow( 21).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell14, data));
        cell = pe.getCell();
        assertEquals( 21, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.14:" + pe);

        // No.15 PropertyRow不正（値が空）
        Cell tagCell15 = sheet3.getRow( 27).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell15, data));
        cell = pe.getCell();
        assertEquals( 27, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.15:" + pe);

        // No.16 DataRowFrom不正（値が空）
        Cell tagCell16 = sheet3.getRow( 30).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell16, data));
        cell = pe.getCell();
        assertEquals( 30, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.16:" + pe);

        // No.17 DataRowTo不正（値が空）
        Cell tagCell17 = sheet3.getRow( 33).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell17, data));
        cell = pe.getCell();
        assertEquals( 33, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.17:" + pe);

        // No.18 PropertyRow不正（値が文字）
        Cell tagCell18 = sheet3.getRow( 36).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell18, data));
        cell = pe.getCell();
        assertEquals( 36, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.18:" + pe);

        // No.19 DataRowFrom不正（値が文字）
        Cell tagCell19 = sheet3.getRow( 39).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell19, data));
        cell = pe.getCell();
        assertEquals( 39, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.19:" + pe);

        // No.20 DataRowTo不正（値が文字）
        Cell tagCell20 = sheet3.getRow( 42).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell20, data));
        cell = pe.getCell();
        assertEquals( 42, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.20:" + pe);

        // No.21 Class不正（存在しないクラスを指定）
        Cell tagCell21 = sheet3.getRow( 46).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell21, data));
        cell = pe.getCell();
        assertEquals( 46, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.21:" + pe);

        // No.22 存在しないプロパティを指定
        Cell tagCell22 = sheet3.getRow( 50).getCell( 0);
        List<Object> list22 = objectsParser.parse( sheet3, tagCell22, data);
        assertEquals( 0, list22.size());

        // No.23 データ型不正
        Cell tagCell23 = sheet3.getRow( 55).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell23, data));
        cell = pe.getCell();
        assertEquals( 57, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.23:" + pe);

        // No.24 クラスに抽象クラスを指定
        Cell tagCell24 = sheet3.getRow( 61).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell24, data));
        cell = pe.getCell();
        assertEquals( 62, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.24:" + pe);

        // No.25 クラスのコンストラクタがprivate
        Cell tagCell25 = sheet3.getRow( 67).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell25, data));
        cell = pe.getCell();
        assertEquals( 68, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.25:" + pe);

        // No.26 子エンティティのプロパティを指定
        Cell tagCell26 = sheet3.getRow( 73).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet3, tagCell26, data));
        cell = pe.getCell();
        assertEquals( 74, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.26:" + pe);

        // No.27 カスタムパーサテスト
        Cell tagCell27 = sheet4.getRow( 1).getCell( 0);
        ChildNameParser childNameParser = new ChildNameParser( "");
        childNameParser.setTag( "@childName");
        assertEquals( "@childName", childNameParser.getTag());
        assertEquals( Boolean.FALSE, childNameParser.isParse( sheet4, null));

        objectsParser.addPropertyParser( childNameParser);
        List<Object> list27 = objectsParser.parse( sheet4, tagCell27, data);
        targetEntity = ( TargetEntity) list27.get( 0);
        assertEquals( "childName1", targetEntity.getChildEntity().getChildName());
        targetEntity = ( TargetEntity) list27.get( 1);
        assertEquals( "childName2", targetEntity.getChildEntity().getChildName());
        assertEquals( 2, list27.size());

        // No.28 カスタムパーサを削除する
        Cell tagCell28 = sheet4.getRow( 8).getCell( 0);
        objectsParser.removePropertyParser( childNameParser);
        List<Object> list28 = objectsParser.parse( sheet4, tagCell28, data);
        assertEquals( 0, list28.size());

        // No.29 複数のカスタムパーサを指定
        Cell tagCell29 = sheet4.getRow( 15).getCell( 0);
        ChildNoParser childNoParser = new ChildNoParser( "@childNo");
        objectsParser.addPropertyParser( childNameParser);
        objectsParser.addPropertyParser( childNoParser);
        List<Object> list29 = objectsParser.parse( sheet4, tagCell29, data);
        targetEntity = ( TargetEntity) list29.get( 0);
        assertEquals( Integer.valueOf( "55"), targetEntity.getNumberInteger());
        assertEquals( "childName5", targetEntity.getChildEntity().getChildName());
        assertEquals( 5, targetEntity.getChildEntity().getChildNo());
        targetEntity = ( TargetEntity) list29.get( 1);
        assertEquals( Integer.valueOf( "66"), targetEntity.getNumberInteger());
        assertEquals( "childName6", targetEntity.getChildEntity().getChildName());
        assertEquals( 6, targetEntity.getChildEntity().getChildNo());
        assertEquals( 2, list29.size());

        // No.30 カスタムパーサを全削除
        Cell tagCell30 = sheet4.getRow( 22).getCell( 0);
        objectsParser.clearPropertyParsers();
        List<Object> list30 = objectsParser.parse( sheet4, tagCell30, data);
        targetEntity = ( TargetEntity) list30.get( 0);
        assertEquals( Integer.valueOf( "77"), targetEntity.getNumberInteger());
        assertEquals( null, targetEntity.getChildEntity());
        targetEntity = ( TargetEntity) list30.get( 1);
        assertEquals( Integer.valueOf( "88"), targetEntity.getNumberInteger());
        assertEquals( null, targetEntity.getChildEntity());
        assertEquals( 2, list30.size());

        // No.31 PropertyRow不正テスト（1行目でプロパティ行にマイナスを指定）
        Cell tagCell31 = sheet5.getRow( 0).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet5, tagCell31, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.31:" + pe);

        // No.32 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        Cell tagCell32 = sheet5.getRow( 0).getCell( 8);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet5, tagCell32, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 8, cell.getColumnIndex());
        System.out.println( "No.32:" + pe);

        // No.33 DataRowTo不正テスト（1行目でデータ終了行にマイナスを指定）
        Cell tagCell33 = sheet5.getRow( 0).getCell( 16);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet5, tagCell33, data));
        cell = pe.getCell();
        assertEquals( 0, cell.getRow().getRowNum());
        assertEquals( 16, cell.getColumnIndex());
        System.out.println( "No.33:" + pe);

        // No.34 PropertyRow不正テスト（最終行でプロパティ行にプラスを指定）
        Cell tagCell34 = sheet5.getRow( 18).getCell( 0);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet5, tagCell34, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 0, cell.getColumnIndex());
        System.out.println( "No.34:" + pe);

        // No.35 DataRowFrom不正テスト（最終行でデータ開始行にプラスを指定）
        Cell tagCell35 = sheet5.getRow( 18).getCell( 8);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet5, tagCell35, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 8, cell.getColumnIndex());
        System.out.println( "No.35:" + pe);

        // No.36 DataRowTo不正テスト（最終行でデータ終了行にプラスを指定）
        Cell tagCell36 = sheet5.getRow( 18).getCell( 16);
        pe = assertThrows( ParseException.class, () -> objectsParser.parse( sheet5, tagCell36, data));
        cell = pe.getCell();
        assertEquals( 18, cell.getRow().getRowNum());
        assertEquals( 16, cell.getColumnIndex());
        System.out.println( "No.36:" + pe);
    }

}
