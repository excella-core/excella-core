/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: DebugErrorHandlerTest.java 144 2009-11-13 06:01:34Z akira-yokoi $
 * $Revision: 144 $
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
package org.bbreak.excella.core.handler;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.Assert;
import org.junit.Test;

/**
 *  DebugErrorHandlerrテストクラス
 * 
 * @since 1.0
 */
public class DebugErrorHandlerTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public DebugErrorHandlerTest( String version) {
        super( version);
    }

    @Test
    public void testDebugErrorHandler() {

        Workbook workbook = getWorkbook();
        Sheet sheet = workbook.getSheetAt( 0);

        String errorFilePath = "DebugErrorHandlerTest" + System.currentTimeMillis();
        if ( workbook instanceof XSSFWorkbook) {
            errorFilePath += BookController.XSSF_SUFFIX;
        } else {
            errorFilePath += BookController.HSSF_SUFFIX;
        }

        DebugErrorHandler debugErrorHandler = new DebugErrorHandler();

        // ===============================================
        // setErrorFilePath(String errorFilePath)
        // ===============================================
        debugErrorHandler.setErrorFilePath( errorFilePath);

        // ===============================================
        // getErrorFilePath()
        // ===============================================
        String actualPath = debugErrorHandler.getErrorFilePath();
        Assert.assertEquals( errorFilePath, actualPath);

        // ===============================================
        // notifyException( Workbook workbook, Sheet sheet, ParseException exception)
        // ===============================================
        debugErrorHandler.notifyException( workbook, sheet, new ParseException( sheet.getRow( 0).getCell( 0), "message", new NullPointerException()));
        debugErrorHandler.notifyException( workbook, sheet, new ParseException( "message"));
    }

    @Test
    public void errorTest1() {

        Workbook workbook = getWorkbook();
        Sheet sheet = workbook.getSheetAt( 0);

        // ===============================================
        // notifyException( Workbook workbook, Sheet sheet, ParseException exception) --> NG
        // ===============================================
        DebugErrorHandler debugErrorHandler = new DebugErrorHandler();
        debugErrorHandler.setErrorFilePath( null);
        debugErrorHandler.notifyException( workbook, sheet, new ParseException( sheet.getRow( 0).getCell( 0)));
    }

    // 存在しないファイルパスを指定するテストだが、環境によっては存在する可能性があるため、コメントアウト
//    @Test
//    public void errorTest2() {
//
//        Workbook workbook = getWorkbook();
//        Sheet sheet = workbook.getSheetAt( 0);
//
//        // ===============================================
//        // notifyException( Workbook workbook, Sheet sheet, ParseException exception) --> NG
//        // ===============================================
//        DebugErrorHandler debugErrorHandler = new DebugErrorHandler();
//        debugErrorHandler.setErrorFilePath( "g:\\result");
//        debugErrorHandler.notifyException( workbook, sheet, new ParseException( sheet.getRow( 0).getCell( 0)));
//
//    }
}
