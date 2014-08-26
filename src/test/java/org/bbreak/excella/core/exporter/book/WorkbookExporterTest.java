/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: WorkbookExporterTest.java 2 2009-05-08 07:39:20Z yuta-takahashi $
 * $Revision: 2 $
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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.File;

import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.WorkbookTest;
import org.bbreak.excella.core.exception.ExportException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * WorkbookExporterテストクラス
 *
 * @since 1.0
 */
public class WorkbookExporterTest extends WorkbookTest {

    /**
     * 作業ディレクトリ
     */
    private File workDire;

    /**
     * 作業ディレクトリ作成結果
     */
    private boolean result;

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public WorkbookExporterTest( String version) {
        super( version);
    }

    @Before
    public void setUp() throws Exception {
        workDire = new File( "workDire");
        result = workDire.mkdir();
        if ( result) {
            System.out.println( "作業ディレクトリ作成 : " + workDire.getAbsolutePath());
        } else {
            System.out.println( "作業ディレクトリが作成できませんでした。 : " + workDire.getAbsolutePath());
            System.out.println( "テスト中断");
        }
    }

    @After
    public void tearDown() throws Exception {
        if ( result) {
            if ( workDire.delete()) {
                System.out.println( "作業ディレクトリ消去 : " + workDire.getAbsolutePath());
            } else {
                System.out.println( "作業ディレクトリが消去できませんでした。 : " + workDire.getAbsolutePath());
            }
        }
    }

    @Test
    public void testWorkbookExporter() throws Exception {

        BookData bookdata = null;
        String fileName = "WorkbookExporterTestFile";

        // No.1 正常ルート
        String filePath = workDire.getAbsolutePath() + File.separatorChar + fileName;
        WorkbookExporter exporter = new WorkbookExporter();
        exporter.setFilePath( filePath);
        exporter.setup();
        exporter.export( getWorkbook(), bookdata);
        exporter.tearDown();

        assertEquals( filePath, exporter.getFilePath());

        File file = new File( filePath);
        file.delete();

        // No.2 不正ルート(ファイルパス不正)
        try {
            String filePath2 = workDire.getAbsolutePath() + File.separatorChar + "dir" + File.separatorChar + fileName;
            WorkbookExporter exporter2 = new WorkbookExporter();
            exporter2.setFilePath( filePath2);
            exporter2.setup();
            exporter2.export( getWorkbook(), bookdata);
            exporter2.tearDown();
            fail();

        } catch ( ExportException ee) {
        }
    }

}
