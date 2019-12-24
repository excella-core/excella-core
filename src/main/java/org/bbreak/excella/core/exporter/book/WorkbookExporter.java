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

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.util.PoiUtil;

/**
 * 解析結果(ブック)を出力するクラス
 * 
 * 指定されたパス(filePath)にブックを出力します。
 * 
 * @since 1.0
 */
public class WorkbookExporter implements BookExporter {
	
    /**
     * ロガー
     */
    private static Log log = LogFactory.getLog( WorkbookExporter.class);

    /**
     * 出力先ファイルパス
     */
    private String filePath = null;

    /**
     * 出力先ファイルパスの取得
     * 
     * @return 出力先ファイルパス
     */
    public String getFilePath() {
        return filePath;
    }

    /**
     * 出力先ファイルパスの設定
     * 
     * @param filePath 出力先ファイルパス
     */
    public void setFilePath( String filePath) {
        this.filePath = filePath;
    }

    /**
     * 初期化処理
     */
    public void setup() {
    }

    /**
     * 処理実行
     */
    public void export( Workbook book, BookData bookdata) throws ExportException {
        if ( log.isInfoEnabled()) {
            log.info( "処理結果を" + filePath + "に出力します");
        }
        try {
            PoiUtil.writeBook( book, filePath);
        } catch ( Exception e) {
            throw new ExportException( e);
        }
    }

    /**
     * 終了処理
     */
    public void tearDown() {
    }
}
