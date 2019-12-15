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
