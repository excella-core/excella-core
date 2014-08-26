/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: BookControllerTest.java 95 2009-05-29 06:08:04Z yuta-takahashi $
 * $Revision: 95 $
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
    public void testBookController2() throws ParseException, ExportException {

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
