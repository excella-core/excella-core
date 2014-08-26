/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: SheetDataTest.java 155 2010-09-09 04:15:12Z akira-yokoi $
 * $Revision: 155 $
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
