/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: TagUtilTest.java 33 2009-05-14 09:51:09Z yuta-takahashi $
 * $Revision: 33 $
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
package org.bbreak.excella.core.util;

import static org.junit.Assert.assertEquals;
import java.util.Map;

import org.bbreak.excella.core.tag.TagParser;
import org.junit.Test;

/**
 * TagUtilテストクラス
 * 
 * @since 1.0
 */
public class TagUtilTest {

    @Test
    public void testTagUtil() {

        String tagDef = "@Map{DataRowFrom=2,DataRowTo=3,KeyColumn=0,ValueColumn=2}";

        // ===============================================
        // getParams( String tagDef)
        // ===============================================
        Map<String, String> resultMap = TagUtil.getParams( "");
        resultMap = TagUtil.getParams( tagDef);
        assertEquals( "2", resultMap.get( "DataRowFrom"));
        assertEquals( "3", resultMap.get( "DataRowTo"));
        assertEquals( "0", resultMap.get( "KeyColumn"));
        assertEquals( "2", resultMap.get( "ValueColumn"));

        // ===============================================
        // getParam( String tagDef)
        // ===============================================
        String param = TagUtil.getParam( "");
        param = TagUtil.getParam( tagDef);
        assertEquals( "DataRowFrom=2,DataRowTo=3,KeyColumn=0,ValueColumn=2", param);

        // ===============================================
        // getTag( String tagDef)
        // ===============================================
        String tag = TagUtil.getTag( "");
        tag = TagUtil.getTag( tagDef);
        assertEquals( "@Map", tag);

        // ===============================================
        // getTag( String tagDef, String tagParamPrefix)
        // ===============================================
        tag = TagUtil.getTag( tagDef, "hogehoge");
        tag = TagUtil.getTag( tagDef, TagParser.TAG_PARAM_PREFIX);
        assertEquals( "@Map", tag);

        // ===============================================
        // adjustValue( int baseValue, Map<String, String> paramMap, String paramKey, int defaultAdjust)
        // ===============================================
        int adjustValue = TagUtil.adjustValue( 0, resultMap, "hogehoge", 0);
        assertEquals( 0, adjustValue);
        adjustValue = TagUtil.adjustValue( 0, resultMap, "DataRowFrom", 0);
        assertEquals( 2, adjustValue);

        // ===============================================
        // getParams( String tagDef)
        // ===============================================
        // イコールのないタグ定義
        tagDef = "@Map{DataRowFrom}";
        resultMap = TagUtil.getParams( tagDef);
        assertEquals( "DataRowFrom", resultMap.get( ""));

    }
}
