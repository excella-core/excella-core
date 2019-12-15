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
        if ( expected.getCellTypeEnum() != actual.getCellTypeEnum()) {
            errors.add( new CheckMessage( "型[" + "セル(" + expected.getRowIndex() + "," + expected.getColumnIndex() + ")" + "]", String.valueOf( expected.getCellTypeEnum()), String.valueOf( actual
                .getCellTypeEnum())));
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
            switch ( cell.getCellTypeEnum()) {
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

        String eCellStyleString = null;
        String aCellStyleString = null;

        if ( expected instanceof HSSFCellStyle) {
            HSSFCellStyle expectedStyle = ( HSSFCellStyle) expected;
            HSSFCellStyle actualStyle = ( HSSFCellStyle) actual;
            eCellStyleString = getCellStyleString( expectedWorkbook, expectedStyle);
            aCellStyleString = getCellStyleString( actualWorkbook, actualStyle);
        } else if ( expected instanceof XSSFCellStyle) {
            XSSFCellStyle expectedStyle = ( XSSFCellStyle) expected;
            XSSFCellStyle actualStyle = ( XSSFCellStyle) actual;
            eCellStyleString = getCellStyleString( expectedStyle);
            aCellStyleString = getCellStyleString( actualStyle);
        }

        if ( !eCellStyleString.equals( aCellStyleString)) {
            errors.add( new CheckMessage( "スタイル", eCellStyleString, aCellStyleString));
            throw new CheckException( errors);

        }
    }

    private static String getCellStyleString( Workbook workbook, HSSFCellStyle cellStyle) {
        StringBuffer sb = new StringBuffer();
        if ( cellStyle != null) {
            HSSFFont font = cellStyle.getFont( workbook);
            // sb.append("FontIndex=").append( cellStyle.getFontIndex()).append( ",");
            sb.append( "Font=").append( getHSSFFontString( ( HSSFWorkbook) workbook, font)).append( ",");

            sb.append( "DataFormat=").append( cellStyle.getDataFormat()).append( ",");
            sb.append( "DataFormatString=").append( cellStyle.getDataFormatString()).append( ",");
            sb.append( "Hidden=").append( cellStyle.getHidden()).append( ",");
            sb.append( "Locked=").append( cellStyle.getLocked()).append( ",");
            sb.append( "Alignment=").append( cellStyle.getAlignmentEnum()).append( ",");
            sb.append( "WrapText=").append( cellStyle.getWrapText()).append( ",");
            sb.append( "VerticalAlignment=").append( cellStyle.getVerticalAlignmentEnum()).append( ",");
            sb.append( "Rotation=").append( cellStyle.getRotation()).append( ",");
            sb.append( "Indention=").append( cellStyle.getIndention()).append( ",");
            sb.append( "BorderLeft=").append( cellStyle.getBorderLeftEnum()).append( ",");
            sb.append( "BorderRight=").append( cellStyle.getBorderRightEnum()).append( ",");
            sb.append( "BorderTop=").append( cellStyle.getBorderTopEnum()).append( ",");
            sb.append( "BorderBottom=").append( cellStyle.getBorderBottomEnum()).append( ",");

            sb.append( "LeftBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getLeftBorderColor())).append( ",");
            sb.append( "RightBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getRightBorderColor())).append( ",");
            sb.append( "TopBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getTopBorderColor())).append( ",");
            sb.append( "BottomBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getBottomBorderColor())).append( ",");

            sb.append( "FillPattern=").append( cellStyle.getFillPatternEnum()).append( ",");
            sb.append( "FillForegroundColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getFillForegroundColor())).append( ",");
            sb.append( "FillBackgroundColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getFillBackgroundColor()));
        }
        return sb.toString();
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

    private static String getHSSFColorString( HSSFWorkbook workbook, short index) {
        HSSFPalette palette = workbook.getCustomPalette();
        if ( palette.getColor( index) != null) {
            HSSFColor color = palette.getColor( index);
            return color.getHexString();
        } else {
            return "";
        }
    }

    private static String getCellStyleString( XSSFCellStyle cellStyle) {
        StringBuffer sb = new StringBuffer();
        if ( cellStyle != null) {
            XSSFFont font = cellStyle.getFont();
            sb.append( "Font=").append( font.getCTFont()).append( ",");
            sb.append( "DataFormat=").append( cellStyle.getDataFormat()).append( ",");
            sb.append( "DataFormatString=").append( cellStyle.getDataFormatString()).append( ",");
            sb.append( "Hidden=").append( cellStyle.getHidden()).append( ",");
            sb.append( "Locked=").append( cellStyle.getLocked()).append( ",");
            sb.append( "Alignment=").append( cellStyle.getAlignmentEnum()).append( ",");
            sb.append( "WrapText=").append( cellStyle.getWrapText()).append( ",");
            sb.append( "VerticalAlignment=").append( cellStyle.getVerticalAlignmentEnum()).append( ",");
            sb.append( "Rotation=").append( cellStyle.getRotation()).append( ",");
            sb.append( "Indention=").append( cellStyle.getIndention()).append( ",");
            sb.append( "BorderLeft=").append( cellStyle.getBorderLeftEnum()).append( ",");
            sb.append( "BorderRight=").append( cellStyle.getBorderRightEnum()).append( ",");
            sb.append( "BorderTop=").append( cellStyle.getBorderTopEnum()).append( ",");
            sb.append( "BorderBottom=").append( cellStyle.getBorderBottomEnum()).append( ",");

            sb.append( "LeftBorderColor=").append( getXSSFColorString( cellStyle.getLeftBorderXSSFColor())).append( ",");
            sb.append( "RightBorderColor=").append( getXSSFColorString( cellStyle.getRightBorderXSSFColor())).append( ",");
            sb.append( "TopBorderColor=").append( getXSSFColorString( cellStyle.getTopBorderXSSFColor())).append( ",");
            sb.append( "BottomBorderColor=").append( getXSSFColorString( cellStyle.getBottomBorderXSSFColor())).append( ",");

            sb.append( "FillPattern=").append( cellStyle.getFillPatternEnum()).append( ",");
            sb.append( "FillForegroundColor=").append( getXSSFColorString( cellStyle.getFillForegroundXSSFColor())).append( ",");
            sb.append( "FillBackgroundColor=").append( getXSSFColorString( cellStyle.getFillBackgroundXSSFColor()));
        }
        return sb.toString();
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
