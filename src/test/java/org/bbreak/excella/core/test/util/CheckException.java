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
import java.util.Collection;
import java.util.List;

/**
 * チェック例外クラス
 * 
 * @since 1.0
 */
@SuppressWarnings( "serial")
public class CheckException extends Exception {

    /** メッセージリスト */
    private List<CheckMessage> checkMessages = new ArrayList<CheckMessage>();

    /**
     * コンストラクタ
     * @param checkMessages メッセージリスト
     */
    public CheckException( List<CheckMessage> checkMessages) {
        this.checkMessages.addAll( checkMessages);
    }

    public boolean add( CheckMessage checkMessage) {
        return checkMessages.add( checkMessage);
    }

    public boolean addAll( Collection<? extends CheckMessage> collection) {
        return checkMessages.addAll( collection);
    }

    public void clear() {
        checkMessages.clear();
    }

    public List<CheckMessage> getCheckMessages() {
        return checkMessages;
    }

    public String getCheckMessagesToString() {
        StringBuffer buffer = new StringBuffer();
        if ( !checkMessages.isEmpty()) {
            buffer.append( "相違がありました。\n");
        }
        for ( CheckMessage checkMessage : checkMessages) {
            buffer.append( "[" + checkMessage.getMessage()).append( "]\n");
            buffer.append( "期待値：").append( checkMessage.getExpected()).append( "\n");
            buffer.append( "実測値：").append( checkMessage.getActual()).append( "\n");
        }
        return buffer.toString();
    }

    @Override
    public String getMessage() {
        return getCheckMessagesToString() + System.lineSeparator()+ super.getMessage();
    }
}
