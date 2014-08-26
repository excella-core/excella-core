/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ObjectsParserTest.java 128 2009-07-02 06:32:17Z yuta-takahashi $
 * $Revision: 128 $
 *
 * This file is part of ExCella Core.
 *
 * ExCella Core is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License version 3
 * only, as published by the Free Software Foundation.
 *
 * ExCella Core is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License version 3 for more details
 * (a copy is included in the COPYING.LESSER file that accompanied this code).
 *
 * You should have received a copy of the GNU Lesser General Public License
 * version 3 along with ExCella Core.  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
package org.bbreak.excella.core.tag.excel2java;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.math.BigDecimal;
import java.text.DateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.excel2java.entity.TargetEntity;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * ObjectsParserテストクラス
 * 
 * @since 1.0
 */
public class ObjectsParserTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public ObjectsParserTest( String version) {
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
    public final void testObjectsParser() throws ParseException, java.text.ParseException {
        Workbook wk = getWorkbook();
        Sheet sheet1 = wk.getSheetAt( 0);
        Sheet sheet2 = wk.getSheetAt( 1);
        Sheet sheet3 = wk.getSheetAt( 2);
        Sheet sheet4 = wk.getSheetAt( 3);
        Sheet sheet5 = wk.getSheetAt( 4);
        ObjectsParser objectsParser = new ObjectsParser( "@Objects");
        Cell tagCell = null;
        Object data = null;
        List<Object> list = null;

        // No.1 パラメータ無し
        tagCell = sheet1.getRow( 5).getCell( 0);
        try {
            list = objectsParser.parse( sheet1, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 5, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.1:" + pe);
        }

        // No.2 クラス指定
        tagCell = sheet1.getRow( 13).getCell( 0);
        list = objectsParser.parse( sheet1, tagCell, data);
        TargetEntity targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Integer( "1"), targetEntity.getNumberInteger());
        assertEquals( 2, targetEntity.getNumberInt());
        assertEquals( new Long( "3"), targetEntity.getNumberLong());
        assertEquals( 4, targetEntity.getNumberlong());
        assertEquals( new Float( "5.5"), targetEntity.getNumberFloat());
        assertEquals( "6.6", String.valueOf( targetEntity.getNumberfloat()));
        assertEquals( new Double( "8.8"), targetEntity.getNumberDouble());
        assertEquals( "9.9", String.valueOf( targetEntity.getNumberdouble()));
        assertEquals( new BigDecimal( 10.1), targetEntity.getNumberDecimal());
        assertEquals( new Byte( "11"), targetEntity.getNumberByte());
        assertEquals( 12, targetEntity.getNumberbyte());
        assertEquals( Boolean.TRUE, targetEntity.getValueBoolean());
        assertEquals( Boolean.FALSE, targetEntity.isValueboolean());
        assertEquals( DateFormat.getDateInstance().parse( "2009/4/13"), targetEntity.getDate());
        assertEquals( "文字列", targetEntity.getString());
        assertEquals( 1, list.size());

        // No.3 キー列、開始列、終了列指定
        tagCell = sheet2.getRow( 4).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Integer( "1"), targetEntity.getNumberInteger());
        assertEquals( 3, targetEntity.getNumberInt());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( new Integer( "2"), targetEntity.getNumberInteger());
        assertEquals( 4, targetEntity.getNumberInt());
        assertEquals( 2, list.size());

        // No.4 コメント有り
        tagCell = sheet2.getRow( 13).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Integer( "5"), targetEntity.getNumberInteger());
        assertEquals( 9, targetEntity.getNumberInt());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( new Integer( "6"), targetEntity.getNumberInteger());
        assertEquals( 10, targetEntity.getNumberInt());
        assertEquals( 2, list.size());

        // No.5 キー行がnull
        tagCell = sheet2.getRow( 20).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        assertEquals( 0, list.size());

        // No.6 キーセルがnull
        tagCell = sheet2.getRow( 28).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( 20000000, targetEntity.getNumberlong());
        assertEquals( new Long( "26000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( 21000000, targetEntity.getNumberlong());
        assertEquals( new Long( "27000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list.get( 2);
        assertEquals( 22000000, targetEntity.getNumberlong());
        assertEquals( new Long( "28000000"), targetEntity.getNumberLong());
        assertEquals( 3, list.size());

        // No.7 キーセル値がnull
        tagCell = sheet2.getRow( 36).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( 29000000, targetEntity.getNumberlong());
        assertEquals( new Long( "35000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( 30000000, targetEntity.getNumberlong());
        assertEquals( new Long( "36000000"), targetEntity.getNumberLong());
        targetEntity = ( TargetEntity) list.get( 2);
        assertEquals( 31000000, targetEntity.getNumberlong());
        assertEquals( new Long( "37000000"), targetEntity.getNumberLong());
        assertEquals( 3, list.size());

        // No.8 値行がnull
        tagCell = sheet2.getRow( 44).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Float( "38000000"), targetEntity.getNumberFloat());
        assertEquals( 40000000, targetEntity.getNumberfloat(), 0.01);
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( new Float( "39000000"), targetEntity.getNumberFloat());
        assertEquals( 41000000, targetEntity.getNumberfloat(), 0.01);
        assertEquals( 2, list.size());

        // No.9 値セルがnull
        tagCell = sheet2.getRow( 52).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Float( "42000000"), targetEntity.getNumberFloat());
        assertEquals( new Double( "44000000"), targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( null, targetEntity.getNumberFloat());
        assertEquals( null, targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list.get( 2);
        assertEquals( new Float( "43000000"), targetEntity.getNumberFloat());
        assertEquals( new Double( "45000000"), targetEntity.getNumberDouble());
        assertEquals( 3, list.size());

        // No.10 値セル値がnull
        tagCell = sheet2.getRow( 60).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Float( "46000000"), targetEntity.getNumberFloat());
        assertEquals( new Double( "48000000"), targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( null, targetEntity.getNumberFloat());
        assertEquals( null, targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list.get( 2);
        assertEquals( new Float( "47000000"), targetEntity.getNumberFloat());
        assertEquals( new Double( "49000000"), targetEntity.getNumberDouble());
        assertEquals( 3, list.size());

        // No.11 重複キー有り
        tagCell = sheet2.getRow( 68).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet2, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Double( "8.8888888"), targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( null, targetEntity.getNumberDouble());
        targetEntity = ( TargetEntity) list.get( 2);
        assertEquals( new Double( "9.9999999"), targetEntity.getNumberDouble());
        assertEquals( 3, list.size());

        // No.12 マイナス範囲指定
        tagCell = sheet3.getRow( 7).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet3, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( Boolean.TRUE, targetEntity.getValueBoolean());
        assertEquals( "値1", targetEntity.getString());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( Boolean.TRUE, targetEntity.getValueBoolean());
        assertEquals( "値2", targetEntity.getString());
        targetEntity = ( TargetEntity) list.get( 2);
        assertEquals( Boolean.FALSE, targetEntity.getValueBoolean());
        assertEquals( "値3", targetEntity.getString());
        targetEntity = ( TargetEntity) list.get( 3);
        assertEquals( Boolean.FALSE, targetEntity.getValueBoolean());
        assertEquals( "値4", targetEntity.getString());
        assertEquals( 4, list.size());

        // No.13 プロパティ列とデータ列を同じ列に指定
        tagCell = sheet3.getRow( 12).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet3, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( "string", targetEntity.getString());
        assertEquals( 1, list.size());

        // No.14 DataRowFrom > DataRowTo
        tagCell = sheet3.getRow( 21).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 21, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.14:" + pe);
        }

        // No.15 PropertyRow不正（値が空）
        tagCell = sheet3.getRow( 27).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 27, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.15:" + pe);
        }

        // No.16 DataRowFrom不正（値が空）
        tagCell = sheet3.getRow( 30).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 30, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.16:" + pe);
        }

        // No.17 DataRowTo不正（値が空）
        tagCell = sheet3.getRow( 33).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 33, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.17:" + pe);
        }

        // No.18 PropertyRow不正（値が文字）
        tagCell = sheet3.getRow( 36).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 36, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.18:" + pe);
        }

        // No.19 DataRowFrom不正（値が文字）
        tagCell = sheet3.getRow( 39).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 39, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.19:" + pe);
        }

        // No.20 DataRowTo不正（値が文字）
        tagCell = sheet3.getRow( 42).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 42, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.20:" + pe);
        }

        // No.21 Class不正（存在しないクラスを指定）
        tagCell = sheet3.getRow( 46).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 46, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.21:" + pe);
        }

        // No.22 存在しないプロパティを指定
        tagCell = sheet3.getRow( 50).getCell( 0);
        list.clear();
        list = objectsParser.parse( sheet3, tagCell, data);
        assertEquals( 0, list.size());

        // No.23 データ型不正
        tagCell = sheet3.getRow( 55).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 57, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.23:" + pe);
        }

        // No.24 クラスに抽象クラスを指定
        tagCell = sheet3.getRow( 61).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 62, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.24:" + pe);
        }

        // No.25 クラスのコンストラクタがprivate
        tagCell = sheet3.getRow( 67).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 68, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.25:" + pe);
        }

        // No.26 子エンティティのプロパティを指定
        tagCell = sheet3.getRow( 73).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet3, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 74, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.26:" + pe);
        }

        // No.27 カスタムパーサテスト
        tagCell = sheet4.getRow( 1).getCell( 0);
        ChildNameParser childNameParser = new ChildNameParser( "");
        childNameParser.setTag( "@childName");
        assertEquals( "@childName", childNameParser.getTag());
        assertEquals( Boolean.FALSE, childNameParser.isParse( sheet4, null));

        objectsParser.addPropertyParser( childNameParser);
        list.clear();
        list = objectsParser.parse( sheet4, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( "childName1", targetEntity.getChildEntity().getChildName());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( "childName2", targetEntity.getChildEntity().getChildName());
        assertEquals( 2, list.size());

        // No.28 カスタムパーサを削除する
        tagCell = sheet4.getRow( 8).getCell( 0);
        list.clear();
        objectsParser.removePropertyParser( childNameParser);
        list = objectsParser.parse( sheet4, tagCell, data);
        assertEquals( 0, list.size());

        // No.29 複数のカスタムパーサを指定
        tagCell = sheet4.getRow( 15).getCell( 0);
        list.clear();
        ChildNoParser childNoParser = new ChildNoParser( "@childNo");
        objectsParser.addPropertyParser( childNameParser);
        objectsParser.addPropertyParser( childNoParser);
        list = objectsParser.parse( sheet4, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Integer( "55"), targetEntity.getNumberInteger());
        assertEquals( "childName5", targetEntity.getChildEntity().getChildName());
        assertEquals( 5, targetEntity.getChildEntity().getChildNo());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( new Integer( "66"), targetEntity.getNumberInteger());
        assertEquals( "childName6", targetEntity.getChildEntity().getChildName());
        assertEquals( 6, targetEntity.getChildEntity().getChildNo());
        assertEquals( 2, list.size());

        // No.30 カスタムパーサを全削除
        tagCell = sheet4.getRow( 22).getCell( 0);
        list.clear();
        objectsParser.clearPropertyParsers();
        list = objectsParser.parse( sheet4, tagCell, data);
        targetEntity = ( TargetEntity) list.get( 0);
        assertEquals( new Integer( "77"), targetEntity.getNumberInteger());
        assertEquals( null, targetEntity.getChildEntity());
        targetEntity = ( TargetEntity) list.get( 1);
        assertEquals( new Integer( "88"), targetEntity.getNumberInteger());
        assertEquals( null, targetEntity.getChildEntity());
        assertEquals( 2, list.size());

        // No.31 PropertyRow不正テスト（1行目でプロパティ行にマイナスを指定）
        tagCell = sheet5.getRow( 0).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.31:" + pe);
        }

        // No.32 DataRowFrom不正テスト（1行目でデータ開始行にマイナスを指定）
        tagCell = sheet5.getRow( 0).getCell( 8);
        list.clear();
        try {
            list = objectsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 8, cell.getColumnIndex());
            System.out.println( "No.32:" + pe);
        }

        // No.33 DataRowTo不正テスト（1行目でデータ終了行にマイナスを指定）
        tagCell = sheet5.getRow( 0).getCell( 16);
        list.clear();
        try {
            list = objectsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 16, cell.getColumnIndex());
            System.out.println( "No.33:" + pe);
        }

        // No.34 PropertyRow不正テスト（最終行でプロパティ行にプラスを指定）
        tagCell = sheet5.getRow( 18).getCell( 0);
        list.clear();
        try {
            list = objectsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.34:" + pe);
        }

        // No.35 DataRowFrom不正テスト（最終行でデータ開始行にプラスを指定）
        tagCell = sheet5.getRow( 18).getCell( 8);
        list.clear();
        try {
            list = objectsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 8, cell.getColumnIndex());
            System.out.println( "No.35:" + pe);
        }

        // No.36 DataRowTo不正テスト（最終行でデータ終了行にプラスを指定）
        tagCell = sheet5.getRow( 18).getCell( 16);
        list.clear();
        try {
            list = objectsParser.parse( sheet5, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 16, cell.getColumnIndex());
            System.out.println( "No.36:" + pe);
        }
    }

}
