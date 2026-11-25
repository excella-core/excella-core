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

package org.bbreak.excella.core.tag;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.exception.ParseException;
import org.junit.Test;

public class TagParserTest {

    @Test
    public void testEscapeAndUnescapeStringValues() {
        TestParser parser = new TestParser( "@Map");

        assertEquals( "\\@Map{}", parser.escape( "@Map{}"));
        assertEquals( "\\\\@Map{}", parser.escape( "\\@Map{}"));
        assertEquals( "@Map{}", parser.unescape( "\\@Map{}"));
        assertEquals( "\\@Map{}", parser.unescape( "\\\\@Map{}"));
    }

    @Test
    public void testIsEscapedAndUnescapeCellValues() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow( 0);
            TestParser parser = new TestParser( "@Map");
            Cell singleEscapedCell = row.createCell( 0);
            singleEscapedCell.setCellValue( "\\@Map{}");
            Cell doubleEscapedCell = row.createCell( 1);
            doubleEscapedCell.setCellValue( "\\\\@Map{}");

            assertTrue( parser.isEscaped( sheet, singleEscapedCell));
            assertTrue( parser.isEscaped( sheet, doubleEscapedCell));
            parser.unescape( sheet, singleEscapedCell);
            parser.unescape( sheet, doubleEscapedCell);
            assertEquals( "@Map{}", singleEscapedCell.getStringCellValue());
            assertEquals( "\\@Map{}", doubleEscapedCell.getStringCellValue());
        }
    }

    @Test
    public void testNonStringAndNullCellsAreNotEscaped() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow( 0);
            TestParser parser = new TestParser( "@Map");
            Cell numericCell = row.createCell( 0);
            numericCell.setCellValue( 100);
            Cell formulaCell = row.createCell( 1);
            formulaCell.setCellFormula( "\"\\@Map{}\"");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell( formulaCell);
            Cell blankCell = row.createCell( 2, CellType.BLANK);

            assertFalse( parser.isEscaped( sheet, numericCell));
            assertFalse( parser.isEscaped( sheet, formulaCell));
            assertFalse( parser.isEscaped( sheet, blankCell));
            assertFalse( parser.isEscaped( sheet, null));
            parser.unescape( sheet, numericCell);
            parser.unescape( sheet, formulaCell);
            parser.unescape( sheet, blankCell);
            parser.unescape( sheet, null);
            assertEquals( CellType.NUMERIC, numericCell.getCellType());
            assertEquals( CellType.FORMULA, formulaCell.getCellType());
            assertEquals( CellType.STRING, formulaCell.getCachedFormulaResultType());
            assertEquals( "\\@Map{}", formulaCell.getStringCellValue());
            assertEquals( CellType.BLANK, blankCell.getCellType());
        }
    }

    @Test(expected = ParseException.class)
    public void testUnescapeCellPropagatesParseException() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Cell cell = sheet.createRow( 0).createCell( 0);
            cell.setCellValue( "\\@Map{}");
            TestParser parser = new TestParser( "@Map") {
                @Override
                public boolean isEscaped( Sheet targetSheet, Cell tagCell) throws ParseException {
                    throw new ParseException( tagCell);
                }
            };

            parser.unescape( sheet, cell);
        }
    }

    @Test
    public void testNullOrEmptyTagDoesNotEscapeValues() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Cell cell = sheet.createRow( 0).createCell( 0);
            cell.setCellValue( "\\@Map{}");

            for ( String tag : new String[] { null, "" }) {
                TestParser parser = new TestParser( tag);
                assertFalse( parser.isEscaped( sheet, cell));
                assertEquals( "\\@Map{}", parser.escape( "\\@Map{}"));
                assertEquals( "\\@Map{}", parser.unescape( "\\@Map{}"));
                parser.unescape( sheet, cell);
                assertEquals( "\\@Map{}", cell.getStringCellValue());
            }
        }
    }

    private static class TestParser extends TagParser<Object> {

        private TestParser( String tag) {
            super( tag);
        }

        @Override
        public Object parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            return null;
        }
    }
}
