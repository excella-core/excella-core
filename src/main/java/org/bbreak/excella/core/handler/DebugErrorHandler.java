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

package org.bbreak.excella.core.handler;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
            errorCellStyle.setFillForegroundColor( HSSFColorPredefined.ROSE.getIndex());
            errorCellStyle.setFillPattern( FillPatternType.SOLID_FOREGROUND);
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
            errorCellStyle.setFillForegroundColor( HSSFColorPredefined.ROSE.getIndex());
            errorCellStyle.setFillPattern( FillPatternType.SOLID_FOREGROUND);
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
