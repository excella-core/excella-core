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

package org.bbreak.excella.core;

import java.io.IOException;
import java.util.List;

import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.exporter.book.BookExporter;
import org.bbreak.excella.core.exporter.book.ConsoleExporter;
import org.bbreak.excella.core.exporter.book.TextFileExporter;
import org.bbreak.excella.core.handler.DebugErrorHandler;
import org.bbreak.excella.core.listener.SheetParseListener;
import org.bbreak.excella.core.tag.excel2java.ListParser;
import org.bbreak.excella.core.tag.excel2java.MapParser;
import org.junit.Assert;
import org.junit.Test;

/**
 * BookControllerテストクラス
 * 
 * @since 1.0
 */
public class BookControllerTest extends WorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version バージョン
     */
    public BookControllerTest( String version) {
        super( version);
     }
    
    @Test
    public void testBookController() throws IOException, ParseException, ExportException {
        
        Workbook workbook = getWorkbook();

        String filePath = getFilepath();
        
        BookController controller = new BookController(filePath);
        workbook = controller.getBook();
        
        
        // ===============================================
        // getSheetNames()
        // ===============================================
        List<String> sheetNames = controller.getSheetNames();

        String sheetName1 = sheetNames.get( 0);
        String sheetName2 = sheetNames.get( 1);
        
        controller.setErrorHandler( new DebugErrorHandler());
        // ===============================================
        // getErrorHandler()
        // ===============================================
        LogFactory.getLog(BookControllerTest.class).info( "====ErrorHandler:" + controller.getErrorHandler());
        
        // ===============================================
        // addTagParser( String sheetName, TagParser<?> parser)
        // ===============================================
        controller.addTagParser( sheetName1, new MapParser( "@Map"));
        
        // ===============================================
        // addSheetParseListener( SheetParseListener listener)
        // ===============================================
        controller.addSheetParseListener( new TestListener());
        
        // ===============================================
        // addSheetParseListener( String sheetName, SheetParseListener listener)
        // ===============================================
        controller.addSheetParseListener( sheetName2, new TestListener());
        
        // ===============================================
        // addSheetExporter( String sheetName, SheetExporter exporter)
        // ===============================================
        controller.addSheetExporter( sheetName1, new org.bbreak.excella.core.exporter.sheet.ConsoleExporter());

        // ===============================================
        // parseSheet( String sheetName)
        // ===============================================
        controller.parseSheet( workbook.getSheetName( 0));

        
//        BookData bookData = BookController.getBookData();
//        Workbook newWorkbook = BookController.getBook();
        

        controller = new BookController( workbook);
        // ===============================================
        // addTagParser( TagParser<?> parser)
        // ===============================================
        controller.addTagParser( new ListParser("@List"));
        // ===============================================
        // addSheetExporter( SheetExporter exporter)
        // ===============================================
        controller.addSheetExporter( new org.bbreak.excella.core.exporter.sheet.ConsoleExporter());
        // ===============================================
        // addBookExporter( BookExporter exporter)
        // ===============================================
        controller.addBookExporter( new ConsoleExporter());
        // ===============================================
        // getExporter()
        // ===============================================
        List<BookExporter> exporterList = controller.getExporter();
        for (BookExporter exporter : exporterList) {
            LogFactory.getLog(BookControllerTest.class).info( "====BookExporter:" + exporter);
        }
        // ===============================================
        // parseBook()
        // ===============================================
        controller.parseBook();

        
    }

    @Test
    @SuppressWarnings("unchecked")
    public void testBookController2() throws ParseException, ExportException, IOException {

        Workbook workbook = getWorkbook();
        BookController controller = new BookController( workbook);
        
        // ===============================================
        // removeTagParser( String tag)
        // ===============================================
        String sheetName = "Sheet3";

        // タグパーサが空の状態でremove
        controller.removeTagParser( "@List");

        controller.addTagParser( sheetName, new ListParser("@List1"));
        controller.addTagParser( sheetName, new ListParser("@List2"));
        controller.addTagParser( sheetName, new ListParser("@List3"));
        // "@List2"をremove
        controller.removeTagParser( "@List2");
        
        SheetData sheetData = controller.parseSheet( sheetName);
        List<String> list1 = (List<String>) sheetData.get( "@List1");
        List<String> list2 = (List<String>) sheetData.get( "@List2");
        List<String> list3 = (List<String>) sheetData.get( "@List3");
        
        Assert.assertEquals( "result1", list1.get( 0));
        Assert.assertEquals( "result2", list1.get( 1));
        Assert.assertNull( list2);
        Assert.assertEquals( "result5", list3.get( 0));
        Assert.assertEquals( "result6", list3.get( 1));
        
        // ===============================================
        // clearTagParsers()
        // ===============================================
        controller.clearTagParsers();
        
        sheetData = controller.parseSheet( sheetName);
        list1 = (List<String>) sheetData.get( "@List1");
        list2 = (List<String>) sheetData.get( "@List2");
        list3 = (List<String>) sheetData.get( "@List3");

        Assert.assertNull( list1);
        Assert.assertNull( list2);
        Assert.assertNull( list3);
        
        // ===============================================
        // clearBookExporters()
        // ===============================================
        controller.addBookExporter( new ConsoleExporter());
        controller.addBookExporter( new TextFileExporter());
        controller.clearBookExporters();
        controller.parseBook();
        
        // ===============================================
        // clearSheetExporters()
        // ===============================================
        controller.addSheetExporter( new org.bbreak.excella.core.exporter.sheet.ConsoleExporter());
        controller.addSheetExporter( new org.bbreak.excella.core.exporter.sheet.TextFileExporter());
        controller.clearSheetExporters();
        controller.parseBook();
        
        // ===============================================
        // clearSheetParseListeners()
        // ===============================================
        controller.addSheetParseListener( new TestListener());
        controller.clearSheetParseListeners();
        controller.parseBook();
    }
    
    /**
     * テスト用リスナ
     *
     * @since 1.0
     */
    private class TestListener implements SheetParseListener {
    
        /**
         * preParse処理
         */
        public void preParse( Sheet sheet, SheetParser sheetParser) throws ParseException {
            System.out.println( "TestListener : preParse実行");
        }
    
        /**
         * postParse処理
         */
        public void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException {
            System.out.println( "TestListener : postParse実行");
        }
    
    }

}
