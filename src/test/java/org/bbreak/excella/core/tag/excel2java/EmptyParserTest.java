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

package org.bbreak.excella.core.tag.excel2java;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * EmptyParserテストクラス
 * 
 * @since 1.0
 */
public class EmptyParserTest extends WorkbookTest {

    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public final void testEmptyParser( String version) throws ParseException, IOException {
        Workbook wk = getWorkbook( version);
        Sheet sheet1 = wk.getSheetAt( 0);
        EmptyParser emptyParser = new EmptyParser( "@Empty");
        Cell tagCell = null;
        Object data = null;
        Object object = null;

        // No.1 パラメータ無し
        tagCell = sheet1.getRow( 5).getCell( 0);
        object = emptyParser.parse( sheet1, tagCell, data);
        assertEquals( null, object);

        // No.2 パラメータ無し
        tagCell = sheet1.getRow( 13).getCell( 0);
        object = emptyParser.parse( sheet1, tagCell, data);
        assertEquals( null, object);

    }
}
