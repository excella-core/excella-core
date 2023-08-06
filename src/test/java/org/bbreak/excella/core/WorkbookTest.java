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
import java.net.URL;
import java.net.URLDecoder;
import java.util.Arrays;
import java.util.Collection;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

/**
 * Workbookテストクラス
 * 
 * @since 1.0
 */
@RunWith( Parameterized.class)
public class WorkbookTest {

    /** Excelファイルのバージョン */
    private String version = null;

    /** ファイルパス */
    private String filepath = null;

    /**
     * コンストラクタ
     * 
     * @param version Excelファイルのバージョン
     */
    public WorkbookTest( String version) {
        this.version = version;
    }

    @Parameters
    public static Collection<?> parameters() {
        return Arrays.asList( new Object[][] {{"2003"}, {"2007"}});
    }

    protected Workbook getWorkbook() {

        Workbook workbook = null;

        String filename = this.getClass().getSimpleName();

        if ( version.equals( "2007")) {
            filename = filename + ".xlsx";
        } else if ( version.equals( "2003")) {
            filename = filename + ".xls";
        }

        URL url = this.getClass().getResource( filename);
        try {
            filepath = URLDecoder.decode( url.getFile(), "UTF-8");
            workbook = WorkbookFactory.create( new File( url.getFile()));
        } catch ( IOException | EncryptedDocumentException e) {
            Assert.fail( e.toString());
        }
        return workbook;
    }

    /**
     * Excelファイルのパスを返却
     * 
     * @return Excelファイルのパス
     */
    public String getFilepath() {
        return filepath;
    }
}
