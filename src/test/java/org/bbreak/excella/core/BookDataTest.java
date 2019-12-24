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
