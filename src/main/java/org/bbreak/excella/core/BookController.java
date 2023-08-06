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

package org.bbreak.excella.core;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.exporter.book.BookExporter;
import org.bbreak.excella.core.exporter.sheet.SheetExporter;
import org.bbreak.excella.core.handler.ParseErrorHandler;
import org.bbreak.excella.core.listener.PostSheetParseListener;
import org.bbreak.excella.core.listener.PreSheetParseListener;
import org.bbreak.excella.core.listener.SheetParseListener;
import org.bbreak.excella.core.tag.TagParser;

/**
 * ワークブックの解析を行うクラス
 * 
 * @since 1.0
 */
public class BookController {

    /** ログ */
    private static Log log = LogFactory.getLog( BookController.class);

    /** Excel2007のファイル末尾 */
    public static final String XSSF_SUFFIX = ".xlsx";

    /** Excel2003以前のファイル末尾 */
    public static final String HSSF_SUFFIX = ".xls";

    /** コメントのプレフィックス */
    public static final String COMMENT_PREFIX = "-";

    /** 処理対象のブック */
    private Workbook workbook;

    /** シート名の一覧 */
    private List<String> sheetNames = new ArrayList<String>();

    /** パーサーのリスト */
    private List<SheetTagParserInfo> tagParsers = new ArrayList<SheetTagParserInfo>();

    /** シート解析イベント処理リスナ(pre parse) */
    private List<SheetListenerInfo<PreSheetParseListener>> preParseListeners = new ArrayList<>();

    /** シート解析イベント処理リスナ(post parse) */
    private List<SheetListenerInfo<PostSheetParseListener>> postParseListeners = new ArrayList<>();

    /** ブック解析結果出力クラス */
    private List<BookExporter> bookExporters = new ArrayList<BookExporter>();

    /** シート解析結果出力クラス */
    private List<SheetExporterInfo> sheetExporters = new ArrayList<SheetExporterInfo>();

    /** エラーハンドリングクラス */
    private ParseErrorHandler errorHandler = null;

    /** ブックデータ */
    private BookData bookData = new BookData();

    /**
     * コンストラクタ<BR>
     * ファイルの形式を判定してWorkbookを生成する
     * 
     * @param filepath ファイルパス
     * @throws IOException ファイルの読み込みに失敗した場合
     */
    public BookController( String filepath) throws IOException {
        if ( log.isInfoEnabled()) {
            log.info( filepath + "の読み込みを開始します");
        }
        workbook = WorkbookFactory.create( new File( filepath));

        // シート名を解析
        int numOfSheets = workbook.getNumberOfSheets();
        for ( int sheetCnt = 0; sheetCnt < numOfSheets; sheetCnt++) {
            String sheetName = workbook.getSheetName( sheetCnt);
            sheetNames.add( sheetName);
        }
    }

    /**
     * コンストラクタ
     * 
     * @param workbook 処理対象のブック
     */
    public BookController( Workbook workbook) {
        this.workbook = workbook;
        // シート名を解析
        int numOfSheets = workbook.getNumberOfSheets();
        for ( int sheetCnt = 0; sheetCnt < numOfSheets; sheetCnt++) {
            String sheetName = workbook.getSheetName( sheetCnt);
            sheetNames.add( sheetName);
        }
    }

    /**
     * ブックに含まれる全シート(コメントシートを除く)の解析の実行
     * 
     * @throws ParseException パースに失敗した場合
     * @throws IOException エラーファイルの書き込みに失敗した場合
     */
    public void parseBook() throws ParseException, ExportException {
        parseBook( null);
    }

    /**
     * ブックに含まれる全シート(コメントシートを除く)の解析の実行
     * 
     * @param data BookControllerのparseBook(), parseSheet()メソッド、 SheetParserのparseSheetメソッドで引数を渡した場合に TagParserまで引き継がれる処理データ
     * @throws ParseException パースに失敗した場合
     * @throws ExportException 出力処理に失敗した場合
     */
    public void parseBook( Object data) throws ParseException, ExportException {
        bookData.clear();
        for ( String sheetName : sheetNames) {
            if ( sheetName.startsWith( COMMENT_PREFIX)) {
                continue;
            }
            // 解析の実行
            SheetData sheetData = parseSheet( sheetName, data);
            // 結果の追加
            bookData.putSheetData( sheetName, sheetData);
        }

        // 出力処理の実行
        for ( BookExporter exporter : bookExporters) {
            if ( exporter != null) {
                exporter.setup();
                try {
                    exporter.export( workbook, bookData);
                } finally {
                    exporter.tearDown();
                }
            }
        }
    }

    /**
     * 現時点での解析データの取得
     * 
     * @return 現時点での解析データ
     */
    public BookData getBookData() {
        return bookData;
    }

    /**
     * 現時点でのWorkbookの取得
     * 
     * @return 現時点でのWorkbook
     */
    public Workbook getBook() {
        return workbook;
    }

    /**
     * シートデータの解析
     * 
     * @param sheetName 解析対象のシート名
     * @return シートの解析結果
     * @throws ParseException パースに失敗した場合
     * @throws ExportException エクスポート処理エラー
     */
    public SheetData parseSheet( String sheetName) throws ParseException, ExportException {
        return parseSheet( sheetName, null);
    }

    /**
     * シートデータの解析
     * 
     * @param sheetName 解析対象のシート名
     * @param data BookControllerのparseBook(), parseSheet()メソッド、 SheetParserのparseSheetメソッドで引数を渡した場合に TagParserまで引き継がれる処理データ
     * @return シートの解析結果
     * @throws ParseException パース処理エラー
     * @throws ExportException エクスポート処理エラー
     */
    public SheetData parseSheet( String sheetName, Object data) throws ParseException, ExportException {
        Sheet sheet = workbook.getSheet( sheetName);

        SheetData sheetData = null;
        try {
            SheetParser sheetParser = new SheetParser();

            if ( log.isInfoEnabled()) {
                log.info( "シート[" + sheetName + "]の処理を開始します");
            }

            // タグパーサの追加
            for ( SheetTagParserInfo parserInfo : tagParsers) {
                String targetSheetName = parserInfo.getSheetName();
                if ( targetSheetName == null || targetSheetName.equals( sheetName)) {
                    sheetParser.addTagParser( parserInfo.getParser());
                }
            }

            // シート処理前イベントの通知
            for ( SheetListenerInfo<PreSheetParseListener> listenerInfo : preParseListeners) {
                String targetSheetName = listenerInfo.getSheetName();
                if ( targetSheetName == null || targetSheetName.equals( sheetName)) {
                    listenerInfo.getListener().preParse( sheet, sheetParser);
                }
            }

            // シート解析処理
            sheetData = sheetParser.parseSheet( sheet, data);

            // シート処理後イベントの通知
            for ( SheetListenerInfo<PostSheetParseListener> listnerInfo : postParseListeners) {
                String targetSheetName = listnerInfo.getSheetName();
                if ( targetSheetName == null || targetSheetName.equals( sheetName)) {
                    listnerInfo.getListener().postParse( sheet, sheetParser, sheetData);
                }
            }
        } catch ( ParseException ex) {
            if ( log.isWarnEnabled()) {
                log.warn( sheetName + "処理中に例外が発生しました", ex);
            }

            // エラーハンドラーへの通知
            if ( errorHandler != null) {
                errorHandler.notifyException( workbook, sheet, ex);
            }
            throw ex;
        }

        // 出力処理の実行
        for ( SheetExporterInfo exporterInfo : sheetExporters) {
            String targetSheetName = exporterInfo.getSheetName();
            if ( targetSheetName == null || targetSheetName.equals( sheetName)) {
                SheetExporter exporter = exporterInfo.getExporter();
                exporter.setup();
                try {
                    exporter.export( sheet, sheetData);
                } finally {
                    exporter.tearDown();
                }
            }
        }
        return sheetData;
    }

    /**
     * ブックに含まれるシート名の一覧取得(コメントシート含む)
     * 
     * @return シート名の一覧
     */
    public List<String> getSheetNames() {
        return sheetNames;
    }

    /**
     * タグパーサの追加
     * 
     * @param parser 追加するタグパーサ
     */
    public void addTagParser( TagParser<?> parser) {
        tagParsers.add( new SheetTagParserInfo( parser));
    }

    /**
     * 対象シート指定でのタグパーサの追加
     * 
     * @param sheetName 対象シート名
     * @param parser 追加するタグパーサ
     */
    public void addTagParser( String sheetName, TagParser<?> parser) {
        tagParsers.add( new SheetTagParserInfo( sheetName, parser));
    }

    /**
     * 指定タグのタグパーサ情報を削除する
     * 
     * @param tag タグ
     */
    public void removeTagParser( String tag) {
        List<SheetTagParserInfo> removeList = new ArrayList<SheetTagParserInfo>();
        for ( SheetTagParserInfo sheetTagParserInfo : tagParsers) {
            if ( sheetTagParserInfo.getParser().getTag().equals( tag)) {
                removeList.add( sheetTagParserInfo);
            }
        }
        tagParsers.removeAll( removeList);
    }

    /**
     * すべてのタグパーサを削除する
     */
    public void clearTagParsers() {
        tagParsers.clear();
    }

    /**
     * シート処理リスナの追加
     * 
     * @param listener 追加するリスナ
     */
    public void addSheetParseListener( SheetParseListener listener) {
        addPreSheetParseListener( listener);
        addPostSheetParseListener( listener);
    }

    /**
     * シート処理リスナの追加
     * 
     * @param sheetName 対象シート名
     * @param listener 追加するリスナ
     */
    public void addSheetParseListener( String sheetName, SheetParseListener listener) {
        addPreSheetParseListener( sheetName, listener);
        addPostSheetParseListener( sheetName, listener);
    }

    /**
     * 全てのシート処理リスナを削除する
     */
    public void clearSheetParseListeners() {
        clearPreSheetParseListeners();
        clearPostSheetParseListeners();
    }

    /**
     * シート処理前イベントリスナの追加
     * 
     * @param listener 追加するリスナ
     * @since 2.1
     */
    public void addPreSheetParseListener( PreSheetParseListener listener) {
        preParseListeners.add( new SheetListenerInfo<>( listener));
    }

    /**
     * シート処理前イベントリスナの追加
     * 
     * @param sheetName 対象シート名
     * @param listener 追加するリスナ
     * @since 2.1
     */
    public void addPreSheetParseListener( String sheetName, PreSheetParseListener listener) {
        preParseListeners.add( new SheetListenerInfo<>( sheetName, listener));
    }

    /**
     * 全てのシート処理前イベントリスナを削除する
     * 
     * @since 2.1
     */
    public void clearPreSheetParseListeners() {
        preParseListeners.clear();
    }

    /**
     * シート処理後イベントリスナの追加
     * 
     * @param listener 追加するリスナ
     * @since 2.1
     */
    public void addPostSheetParseListener( PostSheetParseListener listener) {
        postParseListeners.add( new SheetListenerInfo<>( listener));
    }

    /**
     * シート処理後イベントリスナの追加
     * 
     * @param sheetName 対象シート名
     * @param listener 追加するリスナ
     * @since 2.1
     */
    public void addPostSheetParseListener( String sheetName, PostSheetParseListener listener) {
        postParseListeners.add( new SheetListenerInfo<>( sheetName, listener));
    }

    /**
     * 全てのシート処理後イベントリスナを削除する
     * 
     * @since 2.1
     */
    public void clearPostSheetParseListeners() {
        postParseListeners.clear();
    }

    /**
     * 出力処理クラスの取得
     * 
     * @return 出力処理クラス
     */
    public List<BookExporter> getExporter() {
        return bookExporters;
    }

    /**
     * シート解析結果出力クラスの追加
     * 
     * @param exporter 追加する出力クラス
     */
    public void addSheetExporter( SheetExporter exporter) {
        sheetExporters.add( new SheetExporterInfo( exporter));
    }

    /**
     * シート解析結果出力クラスの追加
     * 
     * @param sheetName 対象シート名
     * @param exporter 追加する出力クラス
     */
    public void addSheetExporter( String sheetName, SheetExporter exporter) {
        sheetExporters.add( new SheetExporterInfo( sheetName, exporter));
    }

    /**
     * すべての解析結果出力クラスを削除する
     */
    public void clearSheetExporters() {
        sheetExporters.clear();
    }

    /**
     * エラーハンドラの取得
     * 
     * @return エラーハンドラ
     */
    public ParseErrorHandler getErrorHandler() {
        return errorHandler;
    }

    /**
     * エラーハンドラの設定
     * 
     * @param errorHandler エラーハンドラ
     */
    public void setErrorHandler( ParseErrorHandler errorHandler) {
        this.errorHandler = errorHandler;
    }

    /**
     * ブック出力処理クラスの追加
     * 
     * @param exporter ブック出力処理クラス
     */
    public void addBookExporter( BookExporter exporter) {
        bookExporters.add( exporter);
    }

    /**
     * 全てのブック出力処理クラスを削除する
     */
    public void clearBookExporters() {
        bookExporters.clear();
    }

    /**
     * シートごとのパーサ情報を保持するクラス
     * 
     * @since 1.0
     */
    private class SheetTagParserInfo {

        /** シート名 */
        private String sheetName = null;

        /** パーサ */
        private TagParser<?> parser = null;

        /**
         * コンストラクタ
         * 
         * @param sheetName シート名
         * @param parser パーサ
         */
        public SheetTagParserInfo( String sheetName, TagParser<?> parser) {
            this.sheetName = sheetName;
            this.parser = parser;
        }

        /**
         * コンストラクタ
         * 
         * @param parser パーサ
         */
        public SheetTagParserInfo( TagParser<?> parser) {
            this.parser = parser;
        }

        public String getSheetName() {
            return sheetName;
        }

        public TagParser<?> getParser() {
            return parser;
        }
    }

    /**
     * シートごとのリスナ情報を保持するクラス
     * 
     * @since 1.0
     */
    private class SheetListenerInfo<T> {

        /** シート名 */
        private String sheetName = null;

        /** リスナ */
        private T listener = null;

        /**
         * コンストラクタ
         * 
         * @param sheetName シート名
         * @param listener リスナ
         */
        public SheetListenerInfo( String sheetName, T listener) {
            this.sheetName = sheetName;
            this.listener = listener;
        }

        /**
         * コンストラクタ
         * 
         * @param listener リスナ
         */
        public SheetListenerInfo( T listener) {
            this.listener = listener;
        }

        public String getSheetName() {
            return sheetName;
        }

        public T getListener() {
            return listener;
        }
    }

    /**
     * シートごとのExporter情報を保持するクラス
     * 
     * @since 1.0
     */
    private class SheetExporterInfo {

        /** シート名 */
        private String sheetName = null;

        /** エクスポータ */
        private SheetExporter exporter = null;

        /**
         * コンストラクタ
         * 
         * @param sheetName シート名
         * @param exporter エクスポータ
         */
        public SheetExporterInfo( String sheetName, SheetExporter exporter) {
            this.sheetName = sheetName;
            this.exporter = exporter;
        }

        /**
         * コンストラクタ
         * 
         * @param exporter エクスポータ
         */
        public SheetExporterInfo( SheetExporter exporter) {
            this.exporter = exporter;
        }

        public String getSheetName() {
            return sheetName;
        }

        public SheetExporter getExporter() {
            return exporter;
        }
    }
}
