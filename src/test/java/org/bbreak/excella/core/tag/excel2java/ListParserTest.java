/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ListParserTest.java 128 2009-07-02 06:32:17Z yuta-takahashi $
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

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.Test;

/**
 * ListParserテストクラス
 * 
 * @since 1.0
 */
public class ListParserTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public ListParserTest( String version) {
        super( version);
    }

    @Test
    public final void testListParser() throws ParseException {
        Workbook wk = getWorkbook();
        Sheet sheet = wk.getSheetAt( 0);
        ListParser listParser = new ListParser( "@List");
        Cell tagCell = null;
        Object data = null;
        List<?> list = null;

        // No.1 @List
        tagCell = sheet.getRow( 2).getCell( 0);
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素1", list.get( 0));
        assertEquals( "要素2", list.get( 1));
        assertEquals( "@List{DataRowFrom=2,DataRowTo=5}", list.get( 2));
        assertEquals( "リスト", list.get( 3));
        assertEquals( "要素3", list.get( 4));
        assertEquals( "要素4", list.get( 5));
        assertEquals( "要素5", list.get( 6));
        assertEquals( "@List{DataRowFrom=2,DataRowTo=3,ValueColumn=1}", list.get( 7));
        assertEquals( null, list.get( 8));
        assertEquals( null, list.get( 9));
        assertEquals( null, list.get( 10));
        assertEquals( "@List{ValueColumn=-1}", list.get( 11));
        assertEquals( "要素8", list.get( 12));
        assertEquals( 13, list.size());

        // No.2 @List{DataRowFrom=2,DataRowTo=5}
        tagCell = sheet.getRow( 6).getCell( 0);
        list.clear();
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素3", list.get( 0));
        assertEquals( "要素4", list.get( 1));
        assertEquals( "要素5", list.get( 2));
        assertEquals( 3, list.size());

        // No.3 @List{DataRowFrom=2,DataRowTo=3,ValueColumn=1}
        tagCell = sheet.getRow( 14).getCell( 0);
        list.clear();
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素6", list.get( 0));
        assertEquals( "要素7", list.get( 1));
        assertEquals( 2, list.size());

        // No.4 @List{ValueColumn=-1}(1列目に書かれていた場合)
        tagCell = sheet.getRow( 24).getCell( 0);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 24, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.4:" + pe);
        }

        // No.5 @List{DataRowFrom=-1}(1行目に書かれている場合)
        tagCell = sheet.getRow( 0).getCell( 4);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 4, cell.getColumnIndex());
            System.out.println( "No.5:" + pe);
        }

        // No.6 @List{DataRowFrom=1}(指定シートの最終行に書かれていた場合)
        tagCell = sheet.getRow( 25).getCell( 4);
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 25, cell.getRow().getRowNum());
            assertEquals( 4, cell.getColumnIndex());
            System.out.println( "No.6:" + pe);
        }

        // No.7 @List{DataRowTo=-1}(1行目に書かれている場合)
        tagCell = sheet.getRow( 0).getCell( 7);
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 7, cell.getColumnIndex());
            System.out.println( "No.7:" + pe);
        }

        // No.8 @List{DataRowTo=1}(指定シートの最終行に書かれていた場合)
        tagCell = sheet.getRow( 25).getCell( 7);
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 25, cell.getRow().getRowNum());
            assertEquals( 7, cell.getColumnIndex());
            System.out.println( "No.8:" + pe);
        }

        // No.9 @List{ValueColumn=1}(最終列に書かれていた場合)
        tagCell = sheet.getRow( 2).getCell( 10);
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 2, cell.getRow().getRowNum());
            assertEquals( 10, cell.getColumnIndex());
            System.out.println( "No.9:" + pe);
        }

        // シート変更
        sheet = wk.getSheetAt( 1);

        // No.10 @List{DataRowFrom=3,DataRowTo=1}(DataRowFrom > DataRowTo の場合)
        tagCell = sheet.getRow( 0).getCell( 0);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.10:" + pe);
        }

        // No.11 @List{DataRowFrom=-1,DataRowTo=-3}(パラメータがマイナス値)
        tagCell = sheet.getRow( 13).getCell( 0);
        list.clear();
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素2-4", list.get( 0));
        assertEquals( "要素2-5", list.get( 1));
        assertEquals( "要素2-6", list.get( 2));
        assertEquals( 3, list.size());

        // No.12 @List{ValueColumn=-1}(パラメータがマイナス値)
        tagCell = sheet.getRow( 29).getCell( 1);
        list.clear();
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素2-7", list.get( 0));
        assertEquals( "要素2-8", list.get( 1));
        assertEquals( "要素2-9", list.get( 2));
        assertEquals( 3, list.size());

        // No.13 @List{DataRowFrom=1}(DataRowFromのみを設定)
        tagCell = sheet.getRow( 27).getCell( 4);
        list.clear();
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素2-10", list.get( 0));
        assertEquals( "要素2-11", list.get( 1));
        assertEquals( "要素2-12", list.get( 2));
        assertEquals( null, list.get( 3));
        assertEquals( "要素2-13", list.get( 4));
        assertEquals( 5, list.size());

        // No.14 @List{DataRowTo=2}(DataRowToのみを設定)
        tagCell = sheet.getRow( 27).getCell( 8);
        list.clear();
        list = listParser.parse( sheet, tagCell, data);
        assertEquals( "要素2-14", list.get( 0));
        assertEquals( "要素2-15", list.get( 1));
        assertEquals( 2, list.size());

        // No.15 @List{DataRowFrom=}(DataRowFrom値空)
        tagCell = sheet.getRow( 0).getCell( 4);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 4, cell.getColumnIndex());
            System.out.println( "No.15:" + pe);
        }

        // No.16 @List{DataRowTo=}(DataRowTo値空)
        tagCell = sheet.getRow( 9).getCell( 4);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 9, cell.getRow().getRowNum());
            assertEquals( 4, cell.getColumnIndex());
            System.out.println( "No.16:" + pe);
        }

        // No.17 @List{ValueColumn=}(ValueColumn値空)
        tagCell = sheet.getRow( 18).getCell( 4);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 4, cell.getColumnIndex());
            System.out.println( "No.17:" + pe);
        }

        // No.18 @List{DataRowFrom=a}(DataRowFrom値文字)
        tagCell = sheet.getRow( 0).getCell( 7);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 7, cell.getColumnIndex());
            System.out.println( "No.18:" + pe);
        }

        // No.19 @List{DataRowTo=a}(DataRowTo値文字)
        tagCell = sheet.getRow( 9).getCell( 7);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 9, cell.getRow().getRowNum());
            assertEquals( 7, cell.getColumnIndex());
            System.out.println( "No.19:" + pe);
        }

        // No.20 @List{ValueColumn=a}(ValueColumn値文字)
        tagCell = sheet.getRow( 18).getCell( 7);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 18, cell.getRow().getRowNum());
            assertEquals( 7, cell.getColumnIndex());
            System.out.println( "No.20:" + pe);
        }

        // シート変更
        sheet = wk.getSheetAt( 1);

        // No.21 @List(要素なし)
        tagCell = sheet.getRow( 0).getCell( 0);
        list.clear();
        try {
            list = listParser.parse( sheet, tagCell, data);
            fail();
        } catch ( ParseException pe) {
            Cell cell = pe.getCell();
            assertEquals( 0, cell.getRow().getRowNum());
            assertEquals( 0, cell.getColumnIndex());
            System.out.println( "No.21:" + pe);
        }
    }
}
