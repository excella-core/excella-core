/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ParseException.java 126 2009-07-02 01:22:11Z akira-yokoi $
 * $Revision: 126 $
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
package org.bbreak.excella.core.exception;

import org.apache.poi.ss.usermodel.Cell;

/**
 * パース例外
 * 
 * @since 1.0
 */
@SuppressWarnings("serial")
public class ParseException extends Exception {
	
    /** エラーとなったセル */
    private Cell cell;

    /**
     * コンストラクタ
     * 
     * @param cell エラーセル
     */
    public ParseException( Cell cell) {
        this.cell = cell;
    }

    /**
     * コンストラクタ
     * 
     * @param cell エラーセル
     * @param message メッセージ
     */
    public ParseException( Cell cell, String message) {
        super( message);
        this.cell = cell;
    }

    /**
     * コンストラクタ
     * 
     * @param cell エラーセル
     * @param throwable スロー可能オブジェクト
     */
    public ParseException( Cell cell, Throwable throwable) {
        super( throwable);
        this.cell = cell;
    }

    /**
     * コンストラクタ
     * 
     * @param cell エラーセル
     * @param message メッセージ
     * @param throwable スロー可能オブジェクト
     */
    public ParseException( Cell cell, String message, Throwable throwable) {
        super( message, throwable);
        this.cell = cell;
    }

    /**
     * コンストラクタ
     * 
     * @param message メッセージ
     */
    public ParseException( String message) {
        super( message);
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell( Cell cell) {
        this.cell = cell;
    }

    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        if ( cell != null) {
            builder.append( "Sheet[" + cell.getSheet().getSheetName() + "] ");
            builder.append( "Cell[" + cell.getRowIndex() + "," + cell.getColumnIndex() + "] ");
        }
        if ( getMessage() != null) {
            builder.append( "Message[" + getMessage() + "] ");
        }
        if ( getCause() != null) {
            builder.append( "Cause[" + getCause() + "] ");
        }
        return builder.toString();
    }
}
