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
