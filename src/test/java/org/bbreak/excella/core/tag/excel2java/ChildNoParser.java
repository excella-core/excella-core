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

package org.bbreak.excella.core.tag.excel2java;

import java.util.Map;

import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.excel2java.entity.ChildEntity;
import org.bbreak.excella.core.tag.excel2java.entity.TargetEntity;

/**
 * テスト用カスタムパーサ
 * 
 * @since 1.0
 */
public class ChildNoParser extends ObjectsPropertyParser {

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ChildNoParser( String tag) {
        super( tag);
    }

    @Override
    public void parse( Object object, Object cellValue, String tag, Map<String, String> params) throws ParseException {

        if ( cellValue == null) {
            return;
        }

        if ( object instanceof TargetEntity) {

            TargetEntity targetEntity = ( TargetEntity) object;
            ChildEntity childEntity = targetEntity.getChildEntity();
            if ( childEntity == null) {
                childEntity = new ChildEntity();
            }

            childEntity.setChildNo( Integer.valueOf( cellValue.toString()));
            targetEntity.setChildEntity( childEntity);
        }
    }
}
