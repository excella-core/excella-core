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

package org.bbreak.excella.samples;

import java.net.URL;
import java.net.URLDecoder;
import java.util.List;

import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.exporter.book.ConsoleExporter;
import org.bbreak.excella.core.handler.DebugErrorHandler;
import org.bbreak.excella.core.tag.excel2java.ArraysParser;
import org.bbreak.excella.core.tag.excel2java.ListParser;
import org.bbreak.excella.core.tag.excel2java.MapParser;
import org.bbreak.excella.core.tag.excel2java.MapsParser;
import org.bbreak.excella.core.tag.excel2java.ObjectsParser;

/**
 * 勤務表取り込みのサンプルクラス
 * 
 * @since 1.0
 */
public class WorkFactReportReader {

    public static void main( String[] args) throws Exception {

        // クラスの場所から読み込むファイルのパスを取得
        String filename = "勤務表.xls";
        URL url = WorkFactReportReader.class.getResource( filename);
        String filepath = URLDecoder.decode( url.getFile(), "UTF-8");

        ///// パース処理 /////
        BookController controller = new BookController( filepath);

        // 標準パーサの追加
        controller.addTagParser( new ListParser( "@List"));
        controller.addTagParser( new MapParser( "@Map"));
        controller.addTagParser( new ArraysParser( "@Arrays"));
        controller.addTagParser( new ObjectsParser( "@Objects"));
        controller.addTagParser( new MapsParser( "@Maps"));

        // コンソール出力用エクスポータの設定
        controller.addBookExporter( new ConsoleExporter());

        // デバッグ用エラーハンドラの設定(エラー時はエラーファイルを作成)
        controller.setErrorHandler( new DebugErrorHandler());

        // パースの実行
        controller.parseBook();

        // パース結果の取得
        BookData bookData = controller.getBookData();

        // パース結果の操作
        List<String> sheetNames = bookData.getSheetNames();
        for ( String sheetName : sheetNames) {
            bookData.getSheetData( sheetName);
        }
    }
}
