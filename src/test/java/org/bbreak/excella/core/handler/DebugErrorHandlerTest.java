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

package org.bbreak.excella.core.handler;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.CoreTestUtil;
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

        String errorFilePath = CoreTestUtil.getTestOutputDir() + "DebugErrorHandlerTest" + System.currentTimeMillis();
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
