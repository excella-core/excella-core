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

import java.util.List;
import java.util.Set;

import org.junit.Assert;
import org.junit.Test;

/**
 * SheetDataテストクラス
 * 
 * @since 1.0
 */
public class SheetDataTest {
    
    @Test
    public void testSheetData() {
        
        SheetData sheetData = new SheetData( "Sheet1");
        
        
        // ===============================================
        // setSheetName(String sheetName)
        // ===============================================
        sheetData.setSheetName( "hogehoge");
        
        
        // ===============================================
        // getSheetName()
        // ===============================================
        //存在するシート名
        Assert.assertEquals( "hogehoge", sheetData.getSheetName());
        
        
        // ===============================================
        // put(String tagName, Object result)
        // ===============================================
        sheetData.put( "tagName1", "testValue");
        
        // ===============================================
        // getTagNames()
        // ===============================================
        Set<String> tagNames = sheetData.getTagNames();
        Assert.assertEquals( "tagName1", tagNames.iterator().next());
        
        // ===============================================
        // containsTag( String tagName)
        // ===============================================
        Assert.assertEquals( Boolean.TRUE, sheetData.containsTag( "tagName1"));
        
        // ===============================================
        // getKeyList( String tagName)
        // ===============================================
        sheetData.put( "tagName2", "testValue2");
        List<String> keyList = sheetData.getKeyList();
        Assert.assertEquals( "tagName1", keyList.get( 0));
        Assert.assertEquals( "tagName2", keyList.get( 1));
        Assert.assertEquals( 2, keyList.size());
        
        // ===============================================
        // put(String tagName, Object result)
        // ===============================================
        // キーが同一ものを追加
        sheetData.put( "tagName1", "testValue1");

        // ===============================================
        // get(String tagName)
        // ===============================================
        String value = (String) sheetData.get( "tagName1");
        Assert.assertEquals( "testValue1", value);

        // ===============================================
        // getList(String tagName)
        // ===============================================
        List<Object> values = sheetData.getList("tagName1");
        Assert.assertEquals( "testValue", values.get( 0));
        Assert.assertEquals( "testValue1", values.get( 1));
        
        // ===============================================
        // remove( String key)
        // ===============================================
        sheetData.remove( "tagName1");
        tagNames = sheetData.getTagNames();
        Assert.assertEquals( Boolean.FALSE, sheetData.containsTag( "tagName1"));

        // ===============================================
        // toString()
        // ===============================================
        String string = "=====================hogehoge=====================\n\ttagName2\ttestValue2";
        Assert.assertEquals( string, sheetData.toString());
        
    }

}
