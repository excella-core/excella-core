/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: DebugErrorHandler.java 142 2009-11-13 05:53:41Z akira-yokoi $
 * $Revision: 142 $
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
package org.bbreak.excella.core.handler;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.StringUtil;

/**
 * デバッグ用エラーハンドリングクラス ParseExceptionが発生した際にエラーとなったセルをマークして実行ディレクトリ配下にエラーファイルを出力する。
 * 
 * @since 1.0
 */
public class DebugErrorHandler implements ParseErrorHandler {
    /** コメントのカラムサイズ */
    private static final short ERROR_COMENT_COL_SIZE = 7;

    /** コメントの行サイズ */
    private static final short ERROR_COMENT_ROW_SIZE = 2;

    /** ログ */
    private static Log log = LogFactory.getLog( DebugErrorHandler.class);

    /** エラーファイルの書き込み先 */
    private String errorFilePath = null;

    public void notifyException( Workbook workbook, Sheet sheet, ParseException exception) {
        // エラーセルのマーク
        if ( exception != null) {
            markupErrorCell( workbook, exception);
        }

        // エラーブックの書き込み
        writeErrorBook( workbook);
    }

    /**
     * エラーセルをマーキングする
     * 
     * @param workbook ワークブック
     * @param errorCell
     * @param exception
     */
    protected void markupErrorCell( Workbook workbook, ParseException exception) {
        Cell errorCell = exception.getCell();
        if ( errorCell == null) {
            return;
        }
        // エラーセルの前景色を赤にする
        workbook.setActiveSheet( workbook.getSheetIndex( errorCell.getSheet()));
        errorCell.setAsActiveCell();

        if ( workbook instanceof XSSFWorkbook) {
            XSSFWorkbook xssfWorkbook = ( XSSFWorkbook) workbook;

            CellStyle errorCellStyle = xssfWorkbook.createCellStyle();
            errorCellStyle.setFillForegroundColor( HSSFColor.ROSE.index);
            errorCellStyle.setFillPattern( XSSFCellStyle.SOLID_FOREGROUND);
            errorCell.setCellStyle( errorCellStyle);

            // TODO:コメントをつけたいけど、うまくいかない。。。
            // XSSFComment xssfComment = ((XSSFSheet)sheet).createComment();
            // xssfComment.setRow( errorCell.getRowIndex());
            // xssfComment.setColumn( (short)errorCell.getColumnIndex());
            // XSSFRichTextString string = new XSSFRichTextString( ex.getMessage());
            // xssfComment.setString( ex.getMessage());
        } else {
            HSSFWorkbook hssfWorkbook = ( HSSFWorkbook) workbook;
            int sheetNum = hssfWorkbook.getNumberOfSheets();
            for ( int cnt = 0; cnt < sheetNum; cnt++) {
                hssfWorkbook.getSheetAt( cnt).setSelected( false);
            }

            // エラーセルに背景色を設定
            CellStyle errorCellStyle = hssfWorkbook.createCellStyle();
            errorCellStyle.setFillForegroundColor( HSSFColor.ROSE.index);
            errorCellStyle.setFillPattern( HSSFCellStyle.SOLID_FOREGROUND);
            errorCell.setCellStyle( errorCellStyle);

            // エラーセルにコメントを追加
            short commentColFrom = ( short) (errorCell.getColumnIndex() + 1);
            short commentColTo = ( short) (errorCell.getColumnIndex() + ERROR_COMENT_COL_SIZE);
            int commentRowFrom = errorCell.getRowIndex();
            int commentRowTo = errorCell.getRowIndex() + ERROR_COMENT_ROW_SIZE;

            HSSFSheet hssfSheet = ( HSSFSheet) errorCell.getSheet();
            HSSFPatriarch patr = hssfSheet.createDrawingPatriarch();
            hssfSheet.setSelected( true);
            HSSFComment comment = patr.createComment( new HSSFClientAnchor( 0, 0, 0, 0, commentColFrom, commentRowFrom, commentColTo, commentRowTo));
            comment.setVisible( true);
            comment.setString( new HSSFRichTextString( createCommentMessage( exception)));
            errorCell.setCellComment( comment);
        }
    }

    /**
     * コメントに出力するメッセージの生成
     * 
     * @param exception 発生した例外
     * @return コメントに出力するメッセージ
     */
    protected String createCommentMessage( ParseException exception) {
        StringBuilder commentMessageBuf = new StringBuilder();
        
        if( exception.getMessage() != null){
            commentMessageBuf.append( exception.getMessage());            
        }
        if ( exception.getCause() != null) {
            commentMessageBuf.append( "\n");
            commentMessageBuf.append( StringUtil.getPrintStackTrace( exception.getCause()));
        }
        return commentMessageBuf.toString();
    }

    /**
     * エラーブックのファイルをカレントディレクトリに書き込む
     * 
     * @param workbook
     */
    protected void writeErrorBook( Workbook workbook) {
        String resultFileName = errorFilePath;
        if ( resultFileName == null) {
            resultFileName = "./" + System.currentTimeMillis();
            if ( workbook instanceof XSSFWorkbook) {
                resultFileName += BookController.XSSF_SUFFIX;
            } else {
                resultFileName += BookController.HSSF_SUFFIX;
            }
        }

        try {
            PoiUtil.writeBook( workbook, resultFileName);
            System.out.println( "\n" + resultFileName + "にエラーファイルを生成しました");
        } catch ( Exception e) {
            log.warn( "エラーファイルの生成に失敗しました", e);
        }
    }

    /**
     * エラーファイルの書き込み先パスの取得
     * 
     * @return エラーファイルの書き込み先パス
     */
    public String getErrorFilePath() {
        return errorFilePath;
    }

    /**
     * エラーファイルの書き込み先パスの設定
     * 
     * @param errorFilePath エラーファイルの書き込み先パス
     */
    public void setErrorFilePath( String errorFilePath) {
        this.errorFilePath = errorFilePath;
    }
}
