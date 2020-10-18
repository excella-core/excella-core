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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.File;
import java.io.IOException;

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
            throw new IOException( "作業ディレクトリが作成できませんでした。 : " + workDire.getAbsolutePath());
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
            fail( "ExportException excepted, but no exception thrown.");

        } catch ( ExportException ee) {
        }
    }

}
