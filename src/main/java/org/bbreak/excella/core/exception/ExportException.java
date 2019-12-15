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

/**
 * 出力処理例外
 * 
 * @since 1.0
 */
@SuppressWarnings( "serial")
public class ExportException extends Exception {

    /**
     * コンストラクタ
     * 
     * @param throwable スロー可能オブジェクト
     */
    public ExportException( Throwable throwable) {
        super( throwable);
    }
}
