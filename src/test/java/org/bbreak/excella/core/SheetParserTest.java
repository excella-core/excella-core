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

import static org.junit.Assert.assertEquals;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.junit.Test;

public class SheetParserTest {

    @Test
    public void testParseSheetUnescapesOneLeadingBackslash() throws Exception {
        String suffix = "{Key=key,Value=value}";

        for ( int backslashCount = 0; backslashCount <= 4; backslashCount++) {
            try ( Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet();
                Cell cell = sheet.createRow( 0).createCell( 0);
                String value = repeatBackslashes( backslashCount) + "@Map" + suffix;
                cell.setCellValue( value);
                RecordingParser parser = new RecordingParser( "@Map");
                SheetParser sheetParser = new SheetParser();
                sheetParser.addTagParser( parser);

                sheetParser.parseSheet( sheet, null);

                assertEquals( repeatBackslashes( Math.max( 0, backslashCount - 1)) + "@Map" + suffix,
                        cell.getStringCellValue());
                assertEquals( backslashCount == 0 ? 1 : 0, parser.getParseCount());
            }
        }
    }

    @Test
    public void testUnmatchedValuesAndNonStringCellsRemainUnchanged() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow( 0);
            String[] values = { "\\1,000", "C:\\temp\\a.txt", "\\", "\\\\" };
            for ( int i = 0; i < values.length; i++) {
                row.createCell( i).setCellValue( values[i]);
            }
            Cell numericCell = row.createCell( 4);
            numericCell.setCellValue( 100);
            Cell formulaCell = row.createCell( 5);
            formulaCell.setCellFormula( "1+1");
            Cell blankCell = row.createCell( 6, CellType.BLANK);
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( new RecordingParser( "@Map"));

            sheetParser.parseSheet( sheet, null);

            for ( int i = 0; i < values.length; i++) {
                assertEquals( values[i], row.getCell( i).getStringCellValue());
            }
            assertEquals( CellType.NUMERIC, numericCell.getCellType());
            assertEquals( CellType.FORMULA, formulaCell.getCellType());
            assertEquals( CellType.BLANK, blankCell.getCellType());
        }
    }

    @Test
    public void testEscapedReplacementIsNotReparsed() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Cell replacementCell = sheet.createRow( 0).createCell( 0);
            replacementCell.setCellValue( "@Replace{}");
            Cell laterRowCell = sheet.createRow( 1).createCell( 0);
            laterRowCell.setCellValue( "\\@Other{}");
            RecordingParser replaceParser = new RecordingParser( "@Replace", "\\@Other{}");
            RecordingParser otherParser = new RecordingParser( "@Other");
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( replaceParser);
            sheetParser.addTagParser( otherParser);

            sheetParser.parseSheet( sheet, null);

            assertEquals( "@Other{}", replacementCell.getStringCellValue());
            assertEquals( "@Other{}", laterRowCell.getStringCellValue());
            assertEquals( 1, replaceParser.getParseCount());
            assertEquals( 0, otherParser.getParseCount());
        }
    }

    @Test
    public void testDuplicateParserRegistrationUnescapesOnlyOnce() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Cell cell = sheet.createRow( 0).createCell( 0);
            cell.setCellValue( "\\\\@Map{}");
            RecordingParser parser = new RecordingParser( "@Map");
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( parser);
            sheetParser.addTagParser( parser);

            sheetParser.parseSheet( sheet, null);

            assertEquals( "\\@Map{}", cell.getStringCellValue());
        }
    }

    @Test
    public void testLastTagLimitsUnescapingToItsPosition() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Row firstRow = sheet.createRow( 0);
            Cell beforeLastTag = firstRow.createCell( 0);
            beforeLastTag.setCellValue( "\\@Map{}");
            Cell lastTag = firstRow.createCell( 1);
            lastTag.setCellValue( "@Stop{LastTag=true}");
            Cell afterLastTag = firstRow.createCell( 2);
            afterLastTag.setCellValue( "\\@Map{}");
            Cell laterRowCell = sheet.createRow( 1).createCell( 0);
            laterRowCell.setCellValue( "\\@Map{}");
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( new RecordingParser( "@Map"));
            sheetParser.addTagParser( new RecordingParser( "@Stop"));

            sheetParser.parseSheet( sheet, null);

            assertEquals( "@Map{}", beforeLastTag.getStringCellValue());
            assertEquals( "@Stop{LastTag=true}", lastTag.getStringCellValue());
            assertEquals( "\\@Map{}", afterLastTag.getStringCellValue());
            assertEquals( "\\@Map{}", laterRowCell.getStringCellValue());
        }
    }

    @Test
    public void testLastTagStateDoesNotAffectNextSheet() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( new RecordingParser( "@Map"));
            sheetParser.addTagParser( new RecordingParser( "@Stop"));

            Sheet firstSheet = workbook.createSheet( "first");
            firstSheet.createRow( 0).createCell( 0).setCellValue( "@Stop{LastTag=true}");
            Cell firstSheetLaterCell = firstSheet.createRow( 1).createCell( 0);
            firstSheetLaterCell.setCellValue( "\\@Map{}");
            sheetParser.parseSheet( firstSheet, null);
            assertEquals( "\\@Map{}", firstSheetLaterCell.getStringCellValue());

            Sheet secondSheet = workbook.createSheet( "second");
            Cell secondSheetCell = secondSheet.createRow( 1).createCell( 0);
            secondSheetCell.setCellValue( "\\@Map{}");
            sheetParser.parseSheet( secondSheet, null);
            assertEquals( "@Map{}", secondSheetCell.getStringCellValue());
        }
    }

    @Test
    public void testCustomParserCanOverrideEscapedCellHandling() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Cell cell = sheet.createRow( 0).createCell( 0);
            cell.setCellValue( "~INLINE{}");
            InlineParser parser = new InlineParser();
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( parser);

            sheetParser.parseSheet( sheet, null);

            assertEquals( "INLINE{}", cell.getStringCellValue());
            assertEquals( 1, parser.getUnescapeCount());
        }
    }

    @Test(expected = ParseException.class)
    public void testUnescapeExceptionIsPropagated() throws Exception {
        try ( Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            sheet.createRow( 0).createCell( 0).setCellValue( "\\@Map{}");
            SheetParser sheetParser = new SheetParser();
            sheetParser.addTagParser( new ThrowingEscapeParser());

            sheetParser.parseSheet( sheet, null);
        }
    }

    private static String repeatBackslashes( int count) {
        StringBuilder result = new StringBuilder();
        for ( int i = 0; i < count; i++) {
            result.append( TagParser.TAG_ESCAPE);
        }
        return result.toString();
    }

    private static class RecordingParser extends TagParser<Object> {

        private final String replacement;

        private int parseCount;

        private RecordingParser( String tag) {
            this( tag, null);
        }

        private RecordingParser( String tag, String replacement) {
            super( tag);
            this.replacement = replacement;
        }

        @Override
        public Object parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            parseCount++;
            if ( replacement != null) {
                tagCell.setCellValue( replacement);
            }
            return null;
        }

        private int getParseCount() {
            return parseCount;
        }
    }

    private static class InlineParser extends TagParser<Object> {

        private int unescapeCount;

        private InlineParser() {
            super( "INLINE");
        }

        @Override
        public boolean isEscaped( Sheet sheet, Cell tagCell) {
            return tagCell != null && tagCell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                    && tagCell.getStringCellValue().startsWith( "~");
        }

        @Override
        public void unescape( Sheet sheet, Cell tagCell) {
            if ( isEscaped( sheet, tagCell)) {
                unescapeCount++;
                tagCell.setCellValue( tagCell.getStringCellValue().substring( 1));
            }
        }

        @Override
        public Object parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            return null;
        }

        private int getUnescapeCount() {
            return unescapeCount;
        }
    }

    private static class ThrowingEscapeParser extends TagParser<Object> {

        private ThrowingEscapeParser() {
            super( "@Map");
        }

        @Override
        public boolean isEscaped( Sheet sheet, Cell tagCell) throws ParseException {
            if ( tagCell != null && tagCell.getCellType() == CellType.STRING) {
                throw new ParseException( tagCell);
            }
            return false;
        }

        @Override
        public Object parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            return null;
        }
    }
}
