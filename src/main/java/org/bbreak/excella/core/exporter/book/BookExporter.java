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

package org.bbreak.excella.core.exporter.book;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;

/**
 * ブック解析結果の出力処理用インターフェイス
 * 
 * 下記の順でメソッドが呼び出されます。
 * 
 * 1.setup
 * 2.export
 * 3.tearDown
 * 
 * @since 1.0
 */
public interface BookExporter {
	/**
     * 初期化処理
     */
    void setup();

    /**
     * 出力処理の実行
     * 
     * @param book 処理後のブック
     * @param bookdata 処理結果のデータ
     */
    void export( Workbook book, BookData bookdata) throws ExportException;

    /**
     * 終了処理
     */
    void tearDown();
}
