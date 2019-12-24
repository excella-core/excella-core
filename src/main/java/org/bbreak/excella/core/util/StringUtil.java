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
