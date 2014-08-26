/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: BookDataTest.java 2 2009-05-08 07:39:20Z yuta-takahashi $
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

import java.util.Collection;

import org.junit.Assert;
import org.junit.Test;

/**
 * BookDataテストクラス
 * 
 * @since 1.0
 */
public class BookDataTest {
    
    @Test
    public void testBookData() {
        
        BookData bookData = new BookData();
        SheetData sheetData = new SheetData("Sheet1");
        bookData.putSheetData( "Sheet1", sheetData);
        
        // ===============================================
        // containsSheet(String sheetName)
        // ===============================================
        //存在するシート名
        Assert.assertTrue( bookData.containsSheet( "Sheet1"));
        //存在しないシート名
        Assert.assertFalse( bookData.containsSheet( "hogehoge"));

        
        // ===============================================
        // getSheetDatas()
        // ===============================================
        Collection<SheetData> collection = bookData.getSheetDatas();
        Assert.assertEquals( sheetData, collection.iterator().next());
        
        
        // ===============================================
        // putSheetData(String sheetName, SheetData sheetData)
        // ===============================================
        //存在するSheetData
        Assert.assertSame( sheetData, bookData.putSheetData( "Sheet1", sheetData));
        //存在しないSheetData
        Assert.assertNotSame( sheetData, bookData.putSheetData( "Sheet4", sheetData));

    }
    
}
