/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ParseExceptionTest.java 2 2009-05-08 07:39:20Z yuta-takahashi $
 * $Revision: 2 $
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
package org.bbreak.excella.core.exception;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.junit.Assert;
import org.junit.Test;

/**
 * ExportExceptionテストクラス
 * 
 * @since 1.0
 */
public class ParseExceptionTest extends WorkbookTest {
    
    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public ParseExceptionTest( String version) {
        super( version);
    }

    @Test
    public void testParseException() {
        Workbook workbook = getWorkbook();

        String message = "test message";
        Cell cell = workbook.getSheetAt( 0).getRow( 0).getCell( 0);
        
        Throwable throwable = new NullPointerException();

        
        // ===============================================
        // ParseException( Cell cell)
        // ===============================================
        ParseException ex = new ParseException( cell);

        
        // ===============================================
        // ParseException( Cell cell, String message)
        // ===============================================
        ex = new ParseException( cell, message);

        
        // ===============================================
        // ParseException( Cell cell, Throwable throwable)
        // ===============================================
        ex = new ParseException( cell, throwable);
        
        
        // ===============================================
        // ParseException( Cell cell, String message, Throwable throwable)
        // ===============================================
        ex = new ParseException( cell, message, throwable);

        
        // ===============================================
        // toString()
        // ===============================================
        ex.toString();
        
        
        // ===============================================
        // ParseException( String message)
        // ===============================================
        ex = new ParseException( message);
        
        
        // ===============================================
        // setCell(Cell cell)
        // ===============================================
        ex.setCell( cell);
        Assert.assertEquals(ex.getCell(), cell);
        
    }

}
