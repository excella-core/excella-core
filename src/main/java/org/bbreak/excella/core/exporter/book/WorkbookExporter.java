/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: WorkbookExporter.java 62 2009-05-21 07:35:40Z akira-yokoi $
 * $Revision: 62 $
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
