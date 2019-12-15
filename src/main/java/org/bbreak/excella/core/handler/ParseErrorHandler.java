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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;

/**
 * パース例外処理用インターフェイス
 * 
 * @since 1.0
 */
public interface ParseErrorHandler extends ErrorHandler<ParseException> {

    /**
     * 例外の通知
     * 
     * @param workbook 処理中のワークブック
     * @param sheet 処理中のシート
     * @param exception 発生した例外
     */
    void notifyException( Workbook workbook, Sheet sheet, ParseException exception);
}
