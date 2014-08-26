/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: TagUtil.java 2 2009-05-08 07:39:20Z yuta-takahashi $
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
package org.bbreak.excella.core.util;

import java.util.HashMap;
import java.util.Map;
import java.util.StringTokenizer;

import org.bbreak.excella.core.tag.TagParser;

/**
 * タグ定義操作ユーティリティ
 * 
 * @since 1.0
 */
public class TagUtil {

    /**
     * コンストラクタ
     */
    private TagUtil() {
    }

    /**
     * タグ定義からパラメータ部分("," "="で分割したマップ)を取得する
     * 
     * @param tagDef タグ定義
     * @return パラメータのマップ
     */
    public static Map<String, String> getParams( String tagDef) {

        Map<String, String> params = new HashMap<String, String>();

        String paramDef = getParam( tagDef);

        if ( paramDef == null) {
            return params;
        }

        // ,で分割
        StringTokenizer commaTokenizer = new StringTokenizer( paramDef, TagParser.PARAM_DELIM);
        while ( commaTokenizer.hasMoreTokens()) {
            String keyValue = ( String) commaTokenizer.nextToken();
            if ( keyValue.indexOf( TagParser.VALUE_DELIM) == -1) {
                params.put( "", keyValue);
            } else {
                // =で分割
                StringTokenizer eqTokenizer = new StringTokenizer( keyValue, TagParser.VALUE_DELIM);
                while ( eqTokenizer.hasMoreTokens()) {
                    params.put( eqTokenizer.nextToken(), eqTokenizer.nextToken());
                }
            }
        }
        return params;
    }

    /**
     * タグ定義からパラーメータ部分の文字列を取得する。<BR>
     * タグ定義の開始文字が存在しない場合はnullを返す。
     * 
     * @param tagDef タグ定義
     * @return パラメータ文字列
     */
    public static String getParam( String tagDef) {

        return getParam( tagDef, TagParser.TAG_PARAM_PREFIX, TagParser.TAG_PARAM_SUFFIX);
    }

    /**
     * タグ定義からパラーメータ部分の文字列を取得する。<BR>
     * タグ定義の開始文字が存在しない場合はnullを返す。
     * 
     * @param tagDef タグ定義
     * @param tagParamPrefix タグの開始文字
     * @param tagParamSuffix タグの終了文字
     * @return パラメータ文字列
     */
    public static String getParam( String tagDef, String tagParamPrefix, String tagParamSuffix) {

        String param = null;

        int paramStartIdx = tagDef.indexOf( tagParamPrefix) + 1;
        int paramEndIdx = tagDef.lastIndexOf( tagParamSuffix);

        if ( paramStartIdx == 0) {
            return param;
        }

        // パラメータ部分を取得
        param = tagDef.substring( paramStartIdx, paramEndIdx);

        return param;
    }

    /**
     * タグ定義からパラメータ部分を除いた文字列を取得する
     * 
     * @param tagDef タグ定義
     * @return パラメータ部分を除いた文字列
     */
    public static String getTag( String tagDef) {

        return getTag( tagDef, TagParser.TAG_PARAM_PREFIX);
    }

    /**
     * タグ定義からパラメータ部分を除いた文字列を取得する
     * 
     * @param tagDef タグ定義
     * @param tagParamPrefix タグの開始文字
     * @return パラメータ部分を除いた文字列
     */
    public static String getTag( String tagDef, String tagParamPrefix) {

        int paramStartIdx = tagDef.indexOf( tagParamPrefix);
        if ( paramStartIdx == -1) {
            return tagDef;
        }
        return tagDef.substring( 0, paramStartIdx);
    }

    /**
     * ベースとなる値とパラメータから調整後の値を取得する
     * 
     * @param baseValue ベースとなる値
     * @param paramMap パラメータのマップ
     * @param paramKey パラメータのキー
     * @param defaultAdjust パラメータにキーが存在しない場合の調整値
     * @return 調整後の値
     */
    public static int adjustValue( int baseValue, Map<String, String> paramMap, String paramKey, int defaultAdjust) {
        int adjustValue = baseValue;

        if ( paramMap.containsKey( paramKey)) {
            String value = paramMap.get( paramKey);
            adjustValue += Integer.valueOf( value);
        } else {
            adjustValue += defaultAdjust;
        }
        return adjustValue;
    }
}
