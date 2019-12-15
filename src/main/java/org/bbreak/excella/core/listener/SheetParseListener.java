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

package org.bbreak.excella.core.listener;

import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ParseException;

/**
 * シート解析時のイベント通知リスナ
 * @since 1.0
 */
public interface SheetParseListener {

    /**
     * シート解析前に呼び出されるメソッド
     * 
     * @param sheet 対象シート
     * @param sheetParser 対象パーサ
     */
    void preParse( Sheet sheet, SheetParser sheetParser) throws ParseException;

    /**
     * シート解析後に呼び出されるメソッド
     * 
     * @param sheet 対象シート
     * @param sheetParser 対象パーサ
     * @param sheetData 解析結果
     */
    void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException;
}
