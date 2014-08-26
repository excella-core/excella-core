/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: StringUtil.java 161 2014-08-11 10:09:05Z kamisono_bb $
 * $Revision: 161 $
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
package org.bbreak.excella.core.util;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;

/**
 * 文字列操作ユーティリティクラス
 *
 * @since 1.4
 */
public class StringUtil {
    /**
     * ThrowableのStackTraceから文字列を生成する
     * 
     * @param ex 対象の例外
     * @return StatckTraceの文字列表現
     */
    public static String getPrintStackTrace(Throwable ex) {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            PrintStream out = new PrintStream(baos);
            ex.printStackTrace(out);
            return baos.toString();
        } finally {
            try {
                baos.close();
            } catch (IOException ioEx) {
                ioEx.printStackTrace();
            }
        }
    }

    /**
     * 文字列が空かどうかを返す
     * @param str 対象文字列
     * @return 空の場合Trueを返す
     */
    public static boolean isEmpty( String str) {
        return str == null || str.isEmpty();
    }

}
