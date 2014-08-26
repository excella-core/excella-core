/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: CheckMessage.java 87 2009-05-22 08:27:33Z yuta-takahashi $
 * $Revision: 87 $
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

/**
 * チェックメッセージクラス
 *
 * @since 1.0
 */
public class CheckMessage {

    /** メッセージ */
    private String message;

    /** 期待値 */
    private String expected;

    /** 実測値 */
    private String actual;

    /**
     * コンストラクタ
     */
    public CheckMessage() {
    }

    /**
     * コンストラクタ
     * 
     * @param message メッセージ
     * @param expected 期待値
     * @param actual 実測値
     */
    public CheckMessage( String message, String expected, String actual) {
        this.message = message;
        this.expected = expected;
        this.actual = actual;
    }

    /**
     * @return the message
     */
    public String getMessage() {
        return message;
    }

    /**
     * @param message the message to set
     */
    public void setMessage( String message) {
        this.message = message;
    }

    /**
     * @return the expected
     */
    public String getExpected() {
        return expected;
    }

    /**
     * @param expected the expected to set
     */
    public void setExpected( String expected) {
        this.expected = expected;
    }

    /**
     * @return the actual
     */
    public String getActual() {
        return actual;
    }

    /**
     * @param actual the actual to set
     */
    public void setActual( String actual) {
        this.actual = actual;
    }

}
