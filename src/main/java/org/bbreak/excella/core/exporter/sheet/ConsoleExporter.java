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

package org.bbreak.excella.core.exporter.sheet;

import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.exception.ExportException;

/**
 * 解析結果のデータをコンソール(標準出力)に出力するクラス
 * 
 * @since 1.0
 */
public class ConsoleExporter implements SheetExporter {
    
    /**
     * 初期化処理
     */
    public void setup() {
    }

    /**
     * 処理実行
     */
    public void export( Sheet sheet, SheetData sheetdata) throws ExportException {
        System.out.println( sheetdata.toString());
    }

    /**
     * 終了処理
     */
    public void tearDown() {
    }
}
