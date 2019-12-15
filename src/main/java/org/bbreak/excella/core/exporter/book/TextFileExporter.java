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

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;


import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.exception.ExportException;

/**
 * 解析結果(データ)をシート単位でファイルに分割して出力する処理クラス
 * 
 * ディレクトリ指定とベースファイルパス指定が可能
 * 両方を指定した場合は双方の設定でファイルが出力される。
 * 
 * ディレクトリ(exportDir)を指定した場合・・・指定ディレクトリに"シート名.txt"でファイルを出力
 * ディレクトリ(baseFilePath)を指定した場合・・・baseFilePath + "シート名.txt"でファイルを出力
 * 
 * @since 1.0
 *
 */
public class TextFileExporter implements BookExporter {
	
    /**
     * ロガー
     */
    private static Log log = LogFactory.getLog( TextFileExporter.class);

    /**
     * 出力先のディレクトリパス
     */
    private String directoryPath = null;

    /**
     * 出力先のベースとなるファイルパス
     */
    private String baseFilePath = null;

    /**
     * 初期化処理
     */
    public void setup() {
    }
		
	/**
	 * 出力処理の実行
	 * 
	 * @param book ワークブック
	 * @param bookdata ブックデータ
	 * @throws ExportException ファイルは存在するが、普通のファイルではなくディレクトリである場合、
     * ファイルは存在せず作成もできない場合、またはなんらかの理由で開くことができない場合
	 */
	public void export( Workbook book, BookData bookdata) throws ExportException {

        List<String> sheetNames = bookdata.getSheetNames();

        if ( directoryPath == null && baseFilePath == null) {
            if ( log.isErrorEnabled()) {
                log.error( "出力先を指定してください。");
            }

        } else {
            // シート毎にファイル書き出し
            for ( String sheetName : sheetNames) {
                SheetData sheetData = bookdata.getSheetData( sheetName);

                // directoryPathが指定されていた場合
                if ( directoryPath != null) {
                    File file = new File( directoryPath, sheetName + ".txt");
                    writeTextFile( file, sheetData.toString());

                }

                // baseFilePathが指定されていた場合
                if ( baseFilePath != null) {
                    int parentPathIndex = baseFilePath.lastIndexOf( File.separatorChar);
                    String parentPath = "";
                    String headFileName = "";

                    try {
                        parentPath = baseFilePath.substring( 0, parentPathIndex);
                        headFileName = baseFilePath.substring( parentPathIndex + 1);

                    } catch ( IndexOutOfBoundsException iobe) {
                        // parentPathIndexが-1の場合(baseFilePathにセパレート文字が無い場合)
                        throw new ExportException( iobe);
                    }

                    File file = new File( parentPath, headFileName + sheetName + ".txt");
                    writeTextFile( file, sheetData.toString());

                }
            }
        }
    }

    /**
     * 終了処理
     */
    public void tearDown() {
    }

    /**
     * 出力先のディレクトリパスを返します。
     * 
     * @return 出力先のディレクトリパス
     */
    public String getDirectoryPath() {
        return directoryPath;
    }

    /**
     * 出力先のディレクトリパスを設定します。
     * 
     * @param directoryPath 出力先のディレクトリパス
     */
    public void setDirectoryPath( String directoryPath) {
        this.directoryPath = directoryPath;
    }

    /**
     * 出力先のベースとなるファイルパスを返します。
     * 
     * @return 出力先のベースとなるファイルパス
     */
    public String getBaseFilePath() {
        return baseFilePath;
    }

    /**
     * 出力先のベースとなるファイルパスを設定します。
     * 
     * @param baseFilePath 出力先のベースとなるファイルパス
     */
    public void setBaseFilePath( String baseFilePath) {
        this.baseFilePath = baseFilePath;
    }

    /**
     * テキストファイル書き込みメソッド
     * 
     * @param file 書き込むファイルオブジェクト
     * @param fileData ファイルに書き込まれる内容
     * @throws ExportException ファイルは存在するが、普通のファイルではなくディレクトリである場合、 
     * ファイルは存在せず作成もできない場合、またはなんらかの理由で開くことができない場合
     */
    private void writeTextFile( File file, String fileData) throws ExportException {
        BufferedWriter writer = null;
        try {
            writer = new BufferedWriter( new FileWriter( file));
            writer.write( fileData);

        } catch ( IOException ioe) {
            if ( log.isErrorEnabled()) {
                log.error( "ファイルを書き込むことができません。");
            }
            throw new ExportException( ioe);

        } finally {
            try {
                if ( writer != null) {
                    writer.close();
                }
            } catch ( IOException ioe) {
                throw new ExportException( ioe);
            }
        }
    }
}
