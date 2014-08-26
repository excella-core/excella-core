/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: WorkbookTest.java 2 2009-05-08 07:39:20Z yuta-takahashi $
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
package org.bbreak.excella.core;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Arrays;
import java.util.Collection;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

/**
 * Workbookテストクラス
 * 
 * @since 1.0
 */
@RunWith(Parameterized.class)
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

            if ( filepath.endsWith( ".xlsx")) {
                try {
                    workbook = new XSSFWorkbook( filepath);
                } catch ( IOException e) {
                    Assert.fail();
                }
            } else if ( filepath.endsWith( ".xls")) {
                FileInputStream stream = null;
                try {
                    stream = new FileInputStream( filepath);
                } catch ( FileNotFoundException e) {
                    Assert.fail();
                }
                POIFSFileSystem fs = null;
                try {
                    fs = new POIFSFileSystem( stream);
                } catch ( IOException e) {
                    Assert.fail();
                }
                try {
                    workbook = new HSSFWorkbook( fs);
                } catch ( IOException e) {
                    Assert.fail();
                }
                try {
                    stream.close();
                } catch ( IOException e) {
                    Assert.fail();
                }
            }
        } catch ( UnsupportedEncodingException e) {
            Assert.fail();
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
