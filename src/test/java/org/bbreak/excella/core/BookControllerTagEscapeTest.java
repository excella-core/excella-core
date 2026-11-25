/*-
 * #%L
 * excella-core
 * %%
 * Copyright (C) 2009 - 2026 bBreak Systems and contributors
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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.exporter.sheet.SheetExporter;
import org.bbreak.excella.core.tag.TagParser;
import org.junit.Test;

public class BookControllerTagEscapeTest {

    @Test
    public void testPostListenerAndExporterObserveCorrectEscapeStates() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet( "data");
            Cell cell = sheet.createRow( 0).createCell( 0);
            cell.setCellValue( "\\\\@Map{}");
            BookController controller = new BookController( workbook);
            RecordingParser parser = new RecordingParser();
            controller.addTagParser( parser);
            controller.addTagParser( parser);
            String[] postParseValue = new String[1];
            String[] exportedValue = new String[1];
            controller.addPostSheetParseListener( (parsedSheet, sheetParser, sheetData) -> {
                postParseValue[0] = cell.getStringCellValue();
                assertTrue( parser.isEscaped( parsedSheet, cell));
            });
            controller.addSheetExporter( new SheetExporter() {
                @Override
                public void setup() {
                    // This exporter does not require setup.
                }

                @Override
                public void export( Sheet exportedSheet, SheetData sheetData) throws ExportException {
                    exportedValue[0] = cell.getStringCellValue();
                }

                @Override
                public void tearDown() {
                    // This exporter does not require teardown.
                }
            });

            controller.parseSheet( "data");

            assertEquals( "\\\\@Map{}", postParseValue[0]);
            assertEquals( "\\@Map{}", exportedValue[0]);
            assertEquals( "\\@Map{}", cell.getStringCellValue());
            assertEquals( 0, parser.getParseCount());
        }
    }

    private static class RecordingParser extends TagParser<Object> {

        private int parseCount;

        private RecordingParser() {
            super( "@Map");
        }

        @Override
        public Object parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            parseCount++;
            return null;
        }

        private int getParseCount() {
            return parseCount;
        }
    }
}
