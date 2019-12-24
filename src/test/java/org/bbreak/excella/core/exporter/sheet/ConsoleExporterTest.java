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

package org.bbreak.excella.core.exporter.sheet;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
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
        String sheetName = "sheetName";
        Sheet sheet = null;
        
        // sheetdata作成
        SheetData sheetdata = new SheetData(sheetName);
        String tagName = "testTag";
        List<Object> result = new ArrayList<Object>();
        result.add( "要素１");
        result.add( "要素２");
        result.add( "要素３");
        sheetdata.put( tagName, result);
        
        ConsoleExporter exporter = new ConsoleExporter();
        exporter.setup();
        exporter.export( sheet, sheetdata);
        exporter.tearDown();
        
    }

}
