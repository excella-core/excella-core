/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ConsoleExporter.java 62 2009-05-21 07:35:40Z akira-yokoi $
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

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.exception.ExportException;

/**
 * 解析結果のデータをコンソール(標準出力)に出力するクラス
 * 
 * @since 1.0
 */
public class ConsoleExporter implements BookExporter {

    /**
     * 初期化処理
     */
    public void setup() {
    }

    /**
     * 処理実行
     */
    public void export( Workbook book, BookData bookdata) throws ExportException {
        List<String> sheetNames = bookdata.getSheetNames();
        // シート単位で出力を行う
        for ( String sheetName : sheetNames) {
            SheetData sheetData = bookdata.getSheetData( sheetName);
            System.out.println( sheetData.toString());
        }
    }

    /**
     * 終了処理
     */
    public void tearDown() {
    }
}
