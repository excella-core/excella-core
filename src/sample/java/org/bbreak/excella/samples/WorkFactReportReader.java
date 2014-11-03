/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: WorkFactReportReader.java 113 2009-06-24 07:58:15Z yuta-takahashi $
 * $Revision: 113 $
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
