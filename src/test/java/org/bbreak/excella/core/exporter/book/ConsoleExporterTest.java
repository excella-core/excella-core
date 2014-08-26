/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ConsoleExporterTest.java 2 2009-05-08 07:39:20Z yuta-takahashi $
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
package org.bbreak.excella.core.exporter.book;


import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.SheetData;
import org.junit.Test;

/**
 * ConsoleExporterテストクラス
 * 
 * @since 1.0
 */
public class ConsoleExporterTest {

    @Test
    public void testConsoleExporter() throws Exception {
        Workbook book = null;
        
        // sheetdata1作成
        String sheetName1 = "sheetName1";
        SheetData sheetdata1 = new SheetData( sheetName1);
        String sheet1Tag = "sheet1Tag";
        List<Object> result1 = new ArrayList<Object>();
        result1.add( "要素１");
        result1.add( "要素２");
        result1.add( "要素３");
        sheetdata1.put( sheet1Tag, result1);

        // sheetdata2作成
        String sheetName2 = "sheetName2";
        SheetData sheetdata2 = new SheetData( sheetName2);
        String sheet2Tag = "sheet2Tag";
        List<Object> result2 = new ArrayList<Object>();
        result2.add( "要素４");
        result2.add( "要素５");
        result2.add( "要素６");
        sheetdata2.put( sheet2Tag, result2);

        // bookdata作成
        BookData bookdata = new BookData();
        bookdata.putSheetData( sheetName1, sheetdata1);
        bookdata.putSheetData( sheetName2, sheetdata2);

        ConsoleExporter expoter = new ConsoleExporter();
        expoter.setup();
        expoter.export( book, bookdata);
        expoter.tearDown();
    }

}
