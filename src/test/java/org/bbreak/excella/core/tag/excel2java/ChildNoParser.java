/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ChildNoParser.java 128 2009-07-02 06:32:17Z yuta-takahashi $
 * $Revision: 128 $
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
