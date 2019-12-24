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
