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

package org.bbreak.excella.core.exporter.sheet;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.exception.ExportException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * TextFileExporterテストクラス
 * 
 * @since 1.0
 */
public class TextFileExporterTest {

    /**
     * 作業ディレクトリ
     */
    private File workDire;

    /**
     * 作業ディレクトリ作成結果
     */
    private boolean result;

    /**
     * @throws java.lang.Exception
     */
    @Before
    public void setUp() throws Exception {
        workDire = new File( "TextFileExporterTestworkDire");
        result = workDire.mkdir();
        if ( result) {
            System.out.println( "作業ディレクトリ作成 : " + workDire.getAbsolutePath());
        } else {
            System.out.println( "作業ディレクトリが作成できませんでした。 : " + workDire.getAbsolutePath());
            System.out.println( "テスト中断");
        }
    }

    /**
     * @throws java.lang.Exception
     */
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

    /**
     * @throws ExportException
     */
    @Test
    public final void testTextFileExporter() throws Exception {
        if ( result) {
            Sheet sheet = null;
            String sheetName = "TextFileExporterTest";
            String headFileName = "test";

            // sheetdata作成
            SheetData sheetdata = new SheetData( sheetName);
            String tagName = "testTag";
            List<Object> result = new ArrayList<Object>();
            result.add( "要素１");
            result.add( "要素２");
            result.add( "要素３");
            sheetdata.put( tagName, result);

            // No.1 正常ルート(directoryPath設定)
            File testFile = null;
            BufferedReader reader = null;
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setDirectoryPath( workDire.getPath());
                exporter.setup();
                exporter.export( sheet, sheetdata);
                exporter.tearDown();

                testFile = new File( workDire.getPath(), sheetName + ".txt");
                reader = new BufferedReader( new FileReader( testFile));

                assertEquals( "=====================" + sheetName + "=====================", reader.readLine());
                assertEquals( "\t" + tagName + "\t" + sheetdata.get( tagName), reader.readLine());
                assertNull( reader.readLine());

            } finally {
                reader.close();
                testFile.delete();
            }

            // No.2 正常ルート(baseFilePath設定)
            File testFile2 = null;
            BufferedReader reader2 = null;
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setBaseFilePath( workDire.getPath() + File.separatorChar + headFileName);
                exporter.setup();
                exporter.export( sheet, sheetdata);
                exporter.tearDown();

                testFile2 = new File( workDire.getPath(), headFileName + sheetName + ".txt");
                reader2 = new BufferedReader( new FileReader( testFile2));

                assertEquals( "=====================" + sheetName + "=====================", reader2.readLine());
                assertEquals( "\t" + tagName + "\t" + sheetdata.get( tagName), reader2.readLine());
                assertNull( reader2.readLine());

            } finally {
                reader2.close();
                testFile2.delete();
            }

            // No.3 正常ルート(directoryPath、baseFilePath両設定)
            File testFile3Dire = null;
            File testFile3Base = null;
            BufferedReader reader3Dire = null;
            BufferedReader reader3Base = null;
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setDirectoryPath( workDire.getPath());
                exporter.setBaseFilePath( workDire.getPath() + File.separatorChar + headFileName);
                exporter.setup();
                exporter.export( sheet, sheetdata);
                exporter.tearDown();

                testFile3Dire = new File( workDire.getPath(), sheetName + ".txt");
                testFile3Base = new File( workDire.getPath(), headFileName + sheetName + ".txt");

                reader3Dire = new BufferedReader( new FileReader( testFile3Dire));
                reader3Base = new BufferedReader( new FileReader( testFile3Base));

                assertEquals( "=====================" + sheetName + "=====================", reader3Dire.readLine());
                assertEquals( "\t" + tagName + "\t" + sheetdata.get( tagName), reader3Dire.readLine());
                assertNull( reader3Dire.readLine());

                assertEquals( "=====================" + sheetName + "=====================", reader3Base.readLine());
                assertEquals( "\t" + tagName + "\t" + sheetdata.get( tagName), reader3Base.readLine());
                assertNull( reader3Base.readLine());

                assertEquals( workDire.getPath(), exporter.getDirectoryPath());
                assertEquals( workDire.getPath() + File.separatorChar + headFileName, exporter.getBaseFilePath());

            } finally {
                reader3Dire.close();
                reader3Base.close();

                testFile3Dire.delete();
                testFile3Base.delete();
            }

            // No.4 不正ルート(directoryPath、baseFilePath両設定無し)
            TextFileExporter exporter4 = new TextFileExporter();
            exporter4.setup();
            exporter4.export( sheet, sheetdata);
            exporter4.tearDown();

            // No.5 不正ルート(baseFilePathディレクトリのみを設定)
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setBaseFilePath( workDire.getPath());
                exporter.setup();
                exporter.export( sheet, sheetdata);
                exporter.tearDown();
                fail();

            } catch ( ExportException ee) {
                System.out.println( ee);
            }

            // No.6 不正ルート(directoryPath設定存在しないパス)
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setDirectoryPath( workDire.getPath() + File.separatorChar + "dir");
                exporter.setup();
                exporter.export( sheet, sheetdata);
                exporter.tearDown();
                fail();

            } catch ( ExportException ee) {
                System.out.println( ee);
            }

        } else {
            fail();
        }
    }

}
