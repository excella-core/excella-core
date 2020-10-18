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

package org.bbreak.excella.core.exception;

import java.io.IOException;

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
    public ParseExceptionTest(String version) {
        super(version);
    }

    @Test
    public void testParseException() throws IOException {
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
