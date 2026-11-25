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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
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

    @Test
    public void testEscapeTagOperations() {
        String tag = "@Map";
        String suffix = "{Key=key,Value=value}";

        for ( int backslashCount = 0; backslashCount <= 4; backslashCount++) {
            String value = repeatBackslashes( backslashCount) + tag + suffix;
            assertEquals( backslashCount > 0, TagUtil.isEscapedTag( value, tag));
            assertEquals( repeatBackslashes( backslashCount + 1) + tag + suffix,
                    TagUtil.escapeTag( value, tag));
            assertEquals( repeatBackslashes( Math.max( 0, backslashCount - 1)) + tag + suffix,
                    TagUtil.unescapeTag( value, tag));
            assertEquals( value, TagUtil.unescapeTag( TagUtil.escapeTag( value, tag), tag));
        }

        assertTrue( TagUtil.isEscapedTag( "\\@Map", tag));
        assertEquals( "@Map", TagUtil.unescapeTag( "\\@Map", tag));
        assertEquals( "@Map{}", TagUtil.unescapeTag( "\\@Map{}", tag));
    }

    @Test
    public void testEscapeTagOperationsWithSpecialTagsAndNewlines() {
        String[] tags = { "$R[]", "$BR[]" };
        for ( String tag : tags) {
            String value = "\\" + tag + "{*}";
            assertTrue( TagUtil.isEscapedTag( value, tag));
            assertEquals( "\\" + value, TagUtil.escapeTag( value, tag));
            assertEquals( tag + "{*}", TagUtil.unescapeTag( value, tag));
        }

        String newlineValue = "\\@Map{first line\nsecond line}";
        assertTrue( TagUtil.isEscapedTag( newlineValue, "@Map"));
        assertEquals( "\\\\@Map{first line\nsecond line}",
                TagUtil.escapeTag( newlineValue, "@Map"));
        assertEquals( newlineValue, TagUtil.unescapeTag(
                TagUtil.escapeTag( newlineValue, "@Map"), "@Map"));
    }

    @Test
    public void testEscapeTagOperationsWithCustomParameterPrefix() {
        String value = "\\@Map[Key=Value]";

        assertTrue( TagUtil.isEscapedTag( value, "@Map", "["));
        assertEquals( "\\\\@Map[Key=Value]", TagUtil.escapeTag( value, "@Map", "["));
        assertEquals( value, TagUtil.unescapeTag(
                TagUtil.escapeTag( value, "@Map", "["), "@Map", "["));
        assertFalse( TagUtil.isEscapedTag( value, "@Map"));
    }

    @Test
    public void testEscapeTagOperationsLeaveUnrelatedValuesUnchanged() {
        String[] values = { "\\1,000", "C:\\temp\\a.txt", "\\", "\\\\" };

        for ( String value : values) {
            assertFalse( TagUtil.isEscapedTag( value, "@Map"));
            assertEquals( value, TagUtil.escapeTag( value, "@Map"));
            assertEquals( value, TagUtil.unescapeTag( value, "@Map"));
        }
    }

    @Test
    public void testEscapeTagOperationsWithNullOrEmptyValuesAndTags() {
        assertFalse( TagUtil.isEscapedTag( null, "@Map"));
        assertFalse( TagUtil.isEscapedTag( "\\@Map", null));
        assertFalse( TagUtil.isEscapedTag( "\\@Map", ""));
        assertEquals( null, TagUtil.escapeTag( null, "@Map"));
        assertEquals( "\\@Map", TagUtil.escapeTag( "\\@Map", null));
        assertEquals( "\\@Map", TagUtil.escapeTag( "\\@Map", ""));
        assertEquals( null, TagUtil.unescapeTag( null, "@Map"));
    }

    private String repeatBackslashes( int count) {
        StringBuilder result = new StringBuilder();
        for ( int i = 0; i < count; i++) {
            result.append( TagParser.TAG_ESCAPE);
        }
        return result.toString();
    }
}
