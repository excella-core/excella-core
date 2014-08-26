/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: TestUtil.java 98 2009-06-12 08:05:23Z tomo-shibata $
 * $Revision: 98 $
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
                case Cell.CELL_TYPE_BLANK:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    value = String.valueOf( cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    value = String.valueOf( cell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    value = String.valueOf( cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA:
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
            sb.append( "Alignment=").append( cellStyle.getAlignment()).append( ",");
            sb.append( "WrapText=").append( cellStyle.getWrapText()).append( ",");
            sb.append( "VerticalAlignment=").append( cellStyle.getVerticalAlignment()).append( ",");
            sb.append( "Rotation=").append( cellStyle.getRotation()).append( ",");
            sb.append( "Indention=").append( cellStyle.getIndention()).append( ",");
            sb.append( "BorderLeft=").append( cellStyle.getBorderLeft()).append( ",");
            sb.append( "BorderRight=").append( cellStyle.getBorderRight()).append( ",");
            sb.append( "BorderTop=").append( cellStyle.getBorderTop()).append( ",");
            sb.append( "BorderBottom=").append( cellStyle.getBorderBottom()).append( ",");

            sb.append( "LeftBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getLeftBorderColor())).append( ",");
            sb.append( "RightBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getRightBorderColor())).append( ",");
            sb.append( "TopBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getTopBorderColor())).append( ",");
            sb.append( "BottomBorderColor=").append( getHSSFColorString( ( HSSFWorkbook) workbook, cellStyle.getBottomBorderColor())).append( ",");

            sb.append( "FillPattern=").append( cellStyle.getFillPattern()).append( ",");
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
        sb.append( "boldweight=").append( Integer.toHexString( font.getBoldweight())).append( ",");
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
            sb.append( "Alignment=").append( cellStyle.getAlignment()).append( ",");
            sb.append( "WrapText=").append( cellStyle.getWrapText()).append( ",");
            sb.append( "VerticalAlignment=").append( cellStyle.getVerticalAlignment()).append( ",");
            sb.append( "Rotation=").append( cellStyle.getRotation()).append( ",");
            sb.append( "Indention=").append( cellStyle.getIndention()).append( ",");
            sb.append( "BorderLeft=").append( cellStyle.getBorderLeft()).append( ",");
            sb.append( "BorderRight=").append( cellStyle.getBorderRight()).append( ",");
            sb.append( "BorderTop=").append( cellStyle.getBorderTop()).append( ",");
            sb.append( "BorderBottom=").append( cellStyle.getBorderBottom()).append( ",");

            sb.append( "LeftBorderColor=").append( getXSSFColorString( cellStyle.getLeftBorderXSSFColor())).append( ",");
            sb.append( "RightBorderColor=").append( getXSSFColorString( cellStyle.getRightBorderXSSFColor())).append( ",");
            sb.append( "TopBorderColor=").append( getXSSFColorString( cellStyle.getTopBorderXSSFColor())).append( ",");
            sb.append( "BottomBorderColor=").append( getXSSFColorString( cellStyle.getBottomBorderXSSFColor())).append( ",");

            sb.append( "FillPattern=").append( cellStyle.getFillPattern()).append( ",");
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
            if ( color.getRgb() != null) {
                for ( byte b : color.getRgb()) {
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
