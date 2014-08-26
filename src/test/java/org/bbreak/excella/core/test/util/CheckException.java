/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: CheckException.java 88 2009-05-22 08:41:57Z yuta-takahashi $
 * $Revision: 88 $
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
}
