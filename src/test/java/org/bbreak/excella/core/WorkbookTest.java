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

import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.net.URLDecoder;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Workbookテストクラス
 * 
 * @since 1.0
 */
public abstract class WorkbookTest {
    
    /** ファイルパス */
    private String filepath = null;

    protected static final String VERSIONS = "2003, 2007";

    protected Workbook getWorkbook( String version) throws IOException {

        Workbook workbook = null;

        String filename = this.getClass().getSimpleName();

        if ( version.equals( "2007")) {
            filename = filename + ".xlsx";
        } else if ( version.equals( "2003")) {
            filename = filename + ".xls";
        }

        URL url = this.getClass().getResource( filename);
        filepath = URLDecoder.decode( url.getFile(), "UTF-8");

        if ( filepath.endsWith( ".xlsx")) {
            workbook = new XSSFWorkbook( filepath);
        } else if ( filepath.endsWith( ".xls")) {
            FileInputStream stream = null;
            stream = new FileInputStream( filepath);
            POIFSFileSystem fs = null;
            fs = new POIFSFileSystem( stream);
            workbook = new HSSFWorkbook( fs);
            stream.close();
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
