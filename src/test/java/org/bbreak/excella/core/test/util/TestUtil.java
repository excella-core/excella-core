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

package org.bbreak.excella.core.test.util;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * テストユーティリティクラス
 *
 * @since 1.0
 */
public final class TestUtil {

    /** ログ */
    private static Log log = LogFactory.getLog( TestUtil.class);

    public static void checkCell( Cell expected, Cell actual) throws CheckException {

        List<CheckMessage> errors = new ArrayList<CheckMessage>();

        // ----------------------
        // セル単位のチェック
        // ----------------------

        if ( expected == null && actual == null) {
            return;
        }

        if ( expected == null) {
            errors.add( new CheckMessage( "セル(" + actual.getRowIndex() + "," + actual.getColumnIndex() + ")", null, actual.toString()));
            throw new CheckException( errors);
        }
        if ( actual == null) {
            errors.add( new CheckMessage( "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")", expected.toString(), null));
            throw new CheckException( errors);
        }

        // 型
        if ( expected.getCellType() != actual.getCellType()) {
            errors.add( new CheckMessage( "型[" + "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + "]", String.valueOf( expected.getCellType()), String.valueOf( actual
                .getCellType())));
            throw new CheckException( errors);
        }

        try {
            checkCellStyle( expected.getRow().getSheet().getWorkbook(), expected.getCellStyle(), actual.getRow().getSheet().getWorkbook(), actual.getCellStyle());
        } catch ( CheckException e) {
            CheckMessage checkMessage = e.getCheckMessages().iterator().next();
            checkMessage.setMessage( "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + checkMessage.getMessage());
            errors.add( checkMessage);
            throw new CheckException( errors);
        }

        // 値
        log.error( "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")");
        if ( !getCellValue( expected).equals( getCellValue( actual))) {
            log.error( getCellValue( expected) + " / " + getCellValue( actual));
            errors.add( new CheckMessage( "値[" + "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + "]", String.valueOf( getCellValue( actual)), String
                .valueOf( getCellValue( expected))));
            throw new CheckException( errors);
        }

    }

    private static Object getCellValue( Cell cell) {
        String value = null;

        if ( cell != null) {
            switch ( cell.getCellType()) {
                case BLANK:
                    value = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    value = String.valueOf( cell.getBooleanCellValue());
                    break;
                case ERROR:
                    value = String.valueOf( cell.getErrorCellValue());
                    break;
                case NUMERIC:
                    value = String.valueOf( cell.getNumericCellValue());
                    break;
                case STRING:
                    value = cell.getStringCellValue();
                    break;
                case FORMULA:
                    value = cell.getCellFormula();
                default:
                    value = "";
            }
        }
        return value;
    }

    private static void checkCellStyle( Workbook expectedWorkbook, CellStyle expected, Workbook actualWorkbook, CellStyle actual) throws CheckException {

        List<CheckMessage> errors = new ArrayList<CheckMessage>();

        if ( expected == null && actual == null) {
            return;
        }

        if ( expected == null) {
            errors.add( new CheckMessage( "スタイル", null, actual.toString()));
            throw new CheckException( errors);
        }

        if ( actual == null) {
            errors.add( new CheckMessage( "スタイル", expected.toString(), null));
            throw new CheckException( errors);
        }

        String eCellStyleString = getCellStyleString(expectedWorkbook, expected);
        String aCellStyleString = getCellStyleString(actualWorkbook, actual);

        if ( !eCellStyleString.equals( aCellStyleString)) {
            errors.add( new CheckMessage( "スタイル", eCellStyleString, aCellStyleString));
            throw new CheckException( errors);

        }
    }

    private static String getCellStyleString(Workbook workbook, CellStyle cellStyle) {
        StringBuffer sb = new StringBuffer();
        if ( cellStyle != null) {
            sb.append("Font=").append(getFontString(workbook, cellStyle)).append(",");
            sb.append( "DataFormat=").append( cellStyle.getDataFormat()).append( ",");
            sb.append( "DataFormatString=").append( cellStyle.getDataFormatString()).append( ",");
            sb.append( "Hidden=").append( cellStyle.getHidden()).append( ",");
            sb.append( "Locked=").append( cellStyle.getLocked()).append( ",");
            sb.append( "Alignment=").append( cellStyle.getAlignment()).append( ",");
            sb.append( "WrapText=").append( cellStyle.getWrapText()).append( ",");
            sb.append( "VerticalAlignment=").append( cellStyle.getVerticalAlignment()).append( ",");
            sb.append( "Rotation=").append( cellStyle.getRotation()).append( ",");
            sb.append( "Indention=").append( cellStyle.getIndention()).append( ",");
            sb.append( "BorderLeft=").append( cellStyle.getBorderLeft()).append( ",");
            sb.append( "BorderRight=").append( cellStyle.getBorderRight()).append( ",");
            sb.append( "BorderTop=").append( cellStyle.getBorderTop()).append( ",");
            sb.append( "BorderBottom=").append( cellStyle.getBorderBottom()).append( ",");

            sb.append( "LeftBorderColor=").append( getColorString(workbook, cellStyle, ColorPosition.LEFT_BORDER)).append( ",");
            sb.append("RightBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.RIGHT_BORDER))
                    .append(",");
            sb.append("TopBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.TOP_BORDER))
                    .append(",");
            sb.append("BottomBorderColor=").append(getColorString(workbook, cellStyle, ColorPosition.BOTTOM_BORDER))
                    .append(",");

            sb.append( "FillPattern=").append( cellStyle.getFillPattern()).append( ",");
            sb.append("FillForegroundColor=").append(getColorString(workbook, cellStyle, ColorPosition.FILL_FOREGROUND))
                    .append(",");
            sb.append("FillBackgroundColor=")
                    .append(getColorString(workbook, cellStyle, ColorPosition.FILL_BACKGROUND));
        }
        return sb.toString();
    }

    private enum ColorPosition {
        LEFT_BORDER(CellStyle::getLeftBorderColor, XSSFCellStyle::getLeftBorderXSSFColor),
        RIGHT_BORDER(CellStyle::getRightBorderColor, XSSFCellStyle::getRightBorderXSSFColor),
        TOP_BORDER(CellStyle::getTopBorderColor, XSSFCellStyle::getTopBorderXSSFColor),
        BOTTOM_BORDER(CellStyle::getBottomBorderColor, XSSFCellStyle::getBottomBorderXSSFColor),
        FILL_FOREGROUND(CellStyle::getFillForegroundColor, XSSFCellStyle::getFillForegroundXSSFColor),
        FILL_BACKGROUND(CellStyle::getFillBackgroundColor, XSSFCellStyle::getFillBackgroundXSSFColor);

        private final Function<CellStyle, Short> colorIndexAccessor;
        private final Function<XSSFCellStyle, XSSFColor> xssfColorAccessor;

        ColorPosition(Function<CellStyle, Short> colorIndexAccessor,
                Function<XSSFCellStyle, XSSFColor> xssfColorAccessor) {
            this.colorIndexAccessor = colorIndexAccessor;
            this.xssfColorAccessor = xssfColorAccessor;
        }

        private short getColorIndex(CellStyle cellStyle) {
            return colorIndexAccessor.apply(cellStyle);
        }

        private XSSFColor getXSSFColor(CellStyle cellStyle) {
            if (cellStyle instanceof XSSFCellStyle) {
                return xssfColorAccessor.apply((XSSFCellStyle) cellStyle);
            }
            throw new IllegalArgumentException("cellStyle is not instanceof XSSFCellStyle: " + cellStyle.getClass());
        }
    }

    private static String getFontString(Workbook workbook, CellStyle cellStyle) {
        if (cellStyle instanceof HSSFCellStyle) {
            HSSFFont font = ((HSSFCellStyle) cellStyle).getFont(workbook);
            return getHSSFFontString((HSSFWorkbook) workbook, font);
        } else if (cellStyle instanceof XSSFCellStyle) {
            XSSFFont font = ((XSSFCellStyle) cellStyle).getFont();
            return font.getCTFont().toString();
        }
        return "";
    }

    private static String getHSSFFontString( HSSFWorkbook workbook, HSSFFont font) {
        StringBuffer sb = new StringBuffer();
        sb.append( "[FONT]");
        sb.append( "fontheight=").append( Integer.toHexString( font.getFontHeight())).append( ",");
        sb.append( "italic=").append( font.getItalic()).append( ",");
        sb.append( "strikout=").append( font.getStrikeout()).append( ",");
        sb.append( "colorpalette=").append( getHSSFColorString( ( HSSFWorkbook) workbook, font.getColor())).append( ",");
        sb.append( "boldweight=").append( font.getBold()).append( ",");
        sb.append( "supersubscript=").append( Integer.toHexString( font.getTypeOffset())).append( ",");
        sb.append( "underline=").append( Integer.toHexString( font.getUnderline())).append( ",");
        sb.append( "charset=").append( Integer.toHexString( font.getCharSet())).append( ",");
        sb.append( "fontname=").append( font.getFontName());
        sb.append( "[/FONT]");
        return sb.toString();
    }

    private static String getColorString(Workbook workbook, CellStyle cellStyle, ColorPosition position) {
        if (cellStyle instanceof HSSFCellStyle) {
            return getHSSFColorString((HSSFWorkbook) workbook, position.getColorIndex(cellStyle));
        } else {
            return getXSSFColorString(position.getXSSFColor(cellStyle));
        }
    }

    private static String getHSSFColorString( HSSFWorkbook workbook, short index) {
        HSSFPalette palette = workbook.getCustomPalette();
        if ( palette.getColor( index) != null) {
            HSSFColor color = palette.getColor( index);
            return color.getHexString();
        } else {
            return "";
        }
    }

    private static String getXSSFColorString( XSSFColor color) {
        StringBuffer sb = new StringBuffer( "[");
        if ( color != null) {
            sb.append( "Indexed=").append( color.getIndexed()).append( ",");
            sb.append( "Rgb=");
            if ( color.getRGB() != null) {
                for ( byte b : color.getRGB()) {
                    sb.append( String.format( "%02x", b).toUpperCase());
                }
            }
            sb.append( ",");
            sb.append( "Tint=").append( color.getTint()).append( ",");
            sb.append( "Theme=").append( color.getTheme()).append( ",");
            sb.append( "Auto=").append( color.isAuto());
        }
        return sb.append( "]").toString();
    }
}
