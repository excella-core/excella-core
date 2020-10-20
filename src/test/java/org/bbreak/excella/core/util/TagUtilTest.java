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

package org.bbreak.excella.core.util;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.Map;

import org.bbreak.excella.core.tag.TagParser;
import org.junit.jupiter.api.Test;

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
