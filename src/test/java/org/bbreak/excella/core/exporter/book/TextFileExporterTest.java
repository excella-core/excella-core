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
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
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

    @Before
    public void setUp() throws Exception {
        workDire = new File("TextFileExporterTestworkDire");
        result = workDire.mkdir();
        if (result) {
            System.out.println("作業ディレクトリ作成 : " + workDire.getAbsolutePath());
        } else {
            System.out.println("作業ディレクトリが作成できませんでした。 : " + workDire.getAbsolutePath());
            System.out.println("テスト中断");
        }
    }

    @After
    public void tearDown() throws Exception {
        if (result) {
            if (workDire.delete()) {
                System.out.println("作業ディレクトリ消去 : " + workDire.getAbsolutePath());
            } else {
                System.out.println("作業ディレクトリが消去できませんでした。 : " + workDire.getAbsolutePath());
            }
        }
    }

    @Test
    public void testTextFileExporter() throws Exception {
        if (result) {
            Workbook book = null;
            String headFileName = "test";
            
            // sheetdata1作成
            String sheetName1 = "sheetName1";
            SheetData sheetdata1 = new SheetData(sheetName1);
            String sheet1Tag = "sheet1Tag";
            List<Object> result1 = new ArrayList<Object>();
            result1.add( "要素１");
            result1.add( "要素２");
            result1.add( "要素３");
            sheetdata1.put( sheet1Tag, result1);
            
            // sheetdata2作成
            String sheetName2 = "sheetName2";
            SheetData sheetdata2 = new SheetData( sheetName2);
            String sheet2Tag = "sheet2Tag";
            List<Object> result2 = new ArrayList<Object>();
            result2.add( "要素４");
            result2.add( "要素５");
            result2.add( "要素６");
            sheetdata2.put( sheet2Tag, result2);

            // bookdata作成
            BookData bookdata = new BookData();
            bookdata.putSheetData( sheetName1, sheetdata1);
            bookdata.putSheetData( sheetName2, sheetdata2);
            
            
            //No.1 正常ルート(directoryPath設定)
            File test1File1 = null;
            BufferedReader test1reader1 = null;
            File test1File2 = null;
            BufferedReader test1reader2 = null;
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setDirectoryPath( workDire.getPath());
                exporter.setup();
                exporter.export( book, bookdata);
                exporter.tearDown();

                //sheetdata1
                test1File1 = new File( workDire.getPath(), sheetName1 + ".txt");
                test1reader1 = new BufferedReader( new FileReader( test1File1));

                assertEquals( "=====================" + sheetName1 + "=====================", test1reader1.readLine());
                assertEquals( "\t" + sheet1Tag + "\t" + sheetdata1.get( sheet1Tag), test1reader1.readLine());
                assertNull( test1reader1.readLine());

                //sheetdata2
                test1File2 = new File( workDire.getPath(), sheetName2 + ".txt");
                test1reader2 = new BufferedReader( new FileReader( test1File2));

                assertEquals( "=====================" + sheetName2 + "=====================", test1reader2.readLine());
                assertEquals( "\t" + sheet2Tag + "\t" + sheetdata2.get( sheet2Tag), test1reader2.readLine());
                assertNull( test1reader2.readLine());

            } finally {
                test1reader1.close();
                test1File1.delete();
                test1reader2.close();
                test1File2.delete();
            }

            //No.2 正常ルート(baseFilePath設定)
            File test2File1 = null;
            BufferedReader test2reader1 = null;
            File test2File2 = null;
            BufferedReader test2reader2 = null;
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setBaseFilePath( workDire.getPath() + File.separatorChar + headFileName);
                exporter.setup();
                exporter.export( book, bookdata);
                exporter.tearDown();

                //sheetdata1
                test2File1 = new File( workDire.getPath(), headFileName + sheetName1 + ".txt");
                test2reader1 = new BufferedReader( new FileReader( test2File1));

                assertEquals( "=====================" + sheetName1 + "=====================", test2reader1.readLine());
                assertEquals( "\t" + sheet1Tag + "\t" + sheetdata1.get( sheet1Tag), test2reader1.readLine());
                assertNull( test2reader1.readLine());

                //sheetdata2
                test2File2 = new File( workDire.getPath(), headFileName + sheetName2 + ".txt");
                test2reader2 = new BufferedReader( new FileReader( test2File2));

                assertEquals( "=====================" + sheetName2 + "=====================", test2reader2.readLine());
                assertEquals( "\t" + sheet2Tag + "\t" + sheetdata2.get( sheet2Tag), test2reader2.readLine());
                assertNull( test2reader2.readLine());

            } finally {
                test2reader1.close();
                test2File1.delete();
                test2reader2.close();
                test2File2.delete();
            }

            //No.3 正常ルート(directoryPath、baseFilePath両設定)
            File test3File1dir = null;
            BufferedReader test3reader1dir = null;
            File test3File2dir = null;
            BufferedReader test3reader2dir = null;
            File test3File1base = null;
            BufferedReader test3reader1base = null;
            File test3File2base = null;
            BufferedReader test3reader2base = null;
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setDirectoryPath( workDire.getPath());
                exporter.setBaseFilePath( workDire.getPath() + File.separatorChar + headFileName);
                exporter.setup();
                exporter.export( book, bookdata);
                exporter.tearDown();

                //sheetdata1
                test3File1dir = new File( workDire.getPath(), sheetName1 + ".txt");
                test3reader1dir = new BufferedReader( new FileReader( test3File1dir));

                assertEquals( "=====================" + sheetName1 + "=====================", test3reader1dir.readLine());
                assertEquals( "\t" + sheet1Tag + "\t" + sheetdata1.get( sheet1Tag), test3reader1dir.readLine());
                assertNull( test3reader1dir.readLine());

                test3File1base = new File( workDire.getPath(), headFileName + sheetName1 + ".txt");
                test3reader1base = new BufferedReader( new FileReader( test3File1base));

                assertEquals( "=====================" + sheetName1 + "=====================", test3reader1base.readLine());
                assertEquals( "\t" + sheet1Tag + "\t" + sheetdata1.get( sheet1Tag), test3reader1base.readLine());
                assertNull( test3reader1base.readLine());

                //sheetdata2
                test3File2dir = new File( workDire.getPath(), sheetName2 + ".txt");
                test3reader2dir = new BufferedReader( new FileReader( test3File2dir));

                assertEquals( "=====================" + sheetName2 + "=====================", test3reader2dir.readLine());
                assertEquals( "\t" + sheet2Tag + "\t" + sheetdata2.get( sheet2Tag), test3reader2dir.readLine());
                assertNull( test3reader2dir.readLine());

                test3File2base = new File( workDire.getPath(), headFileName + sheetName2 + ".txt");
                test3reader2base = new BufferedReader( new FileReader( test3File2base));

                assertEquals( "=====================" + sheetName2 + "=====================", test3reader2base.readLine());
                assertEquals( "\t" + sheet2Tag + "\t" + sheetdata2.get( sheet2Tag), test3reader2base.readLine());
                assertNull( test3reader2base.readLine());

                assertEquals( workDire.getPath(), exporter.getDirectoryPath());
                assertEquals( workDire.getPath() + File.separatorChar + headFileName, exporter.getBaseFilePath());

            } finally {
                test3reader1dir.close();
                test3File1dir.delete();
                test3reader1base.close();
                test3File1base.delete();

                test3reader2dir.close();
                test3File2dir.delete();
                test3reader2base.close();
                test3File2base.delete();
            }

            //No.4 不正ルート(directoryPath、baseFilePath両設定無し)
            TextFileExporter exporter4 = new TextFileExporter();
            exporter4.setup();
            exporter4.export( book, bookdata);
            exporter4.tearDown();

            //No.5 不正ルート(baseFilePathディレクトリのみを設定)
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setBaseFilePath( workDire.getPath());
                exporter.setup();
                exporter.export( book, bookdata);
                exporter.tearDown();
                fail();

            } catch ( ExportException ee) {
                System.out.println( ee);
            }

            //No.6 不正ルート(directoryPath設定存在しないパス)
            try {
                TextFileExporter exporter = new TextFileExporter();
                exporter.setDirectoryPath( workDire.getPath() + File.separatorChar + "dir");
                exporter.setup();
                exporter.export( book, bookdata);
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
