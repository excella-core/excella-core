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

import java.util.HashMap;
import java.util.Map;
import java.util.StringTokenizer;

import org.bbreak.excella.core.tag.TagParser;

/**
 * タグ定義操作ユーティリティ
 *
 * タグ名の先頭にあるバックスラッシュはエスケープ文字として予約されており、タグ名には使用できません。
 * tagParamPrefixを受け取るメソッドではnullを指定できません。空文字の場合はタグに一致しません。
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
     * タグ定義がエスケープされているかどうかを判定する
     *
     * @param tagDef タグ定義
     * @param tag 対象タグ（先頭のバックスラッシュはエスケープ文字として予約）
     * @return タグ定義がエスケープされている場合はTrue
     */
    public static boolean isEscapedTag( String tagDef, String tag) {
        return isEscapedTag( tagDef, tag, TagParser.TAG_PARAM_PREFIX);
    }

    /**
     * タグ定義がエスケープされているかどうかを判定する
     *
     * @param tagDef タグ定義
     * @param tag 対象タグ（先頭のバックスラッシュはエスケープ文字として予約）
     * @param tagParamPrefix タグのパラメータ開始文字（null不可。空文字の場合は一致しない）
     * @return タグ定義がエスケープされている場合はTrue
     */
    public static boolean isEscapedTag( String tagDef, String tag, String tagParamPrefix) {
        if ( tagDef == null || tag == null || tag.isEmpty() || !tagDef.startsWith( TagParser.TAG_ESCAPE)) {
            return false;
        }

        String tagName = getTag( tagDef, tagParamPrefix);
        int escapeLength = getLeadingEscapeLength( tagName);
        return escapeLength > 0 && tagName.substring( escapeLength).equals( tag);
    }

    /**
     * タグ定義をエスケープする
     *
     * @param value タグ定義
     * @param tag 対象タグ（先頭のバックスラッシュはエスケープ文字として予約）
     * @return エスケープ後のタグ定義
     */
    public static String escapeTag( String value, String tag) {
        return escapeTag( value, tag, TagParser.TAG_PARAM_PREFIX);
    }

    /**
     * タグ定義をエスケープする
     *
     * @param value タグ定義
     * @param tag 対象タグ（先頭のバックスラッシュはエスケープ文字として予約）
     * @param tagParamPrefix タグのパラメータ開始文字（null不可。空文字の場合は一致しない）
     * @return エスケープ後のタグ定義
     */
    public static String escapeTag( String value, String tag, String tagParamPrefix) {
        if ( value == null || tag == null || tag.isEmpty()) {
            return value;
        }

        String tagName = getTag( value, tagParamPrefix);
        int escapeLength = getLeadingEscapeLength( tagName);
        if ( tagName.substring( escapeLength).equals( tag)) {
            return TagParser.TAG_ESCAPE + value;
        }
        return value;
    }

    /**
     * タグ定義のエスケープを解除する
     *
     * @param value タグ定義
     * @param tag 対象タグ（先頭のバックスラッシュはエスケープ文字として予約）
     * @return エスケープ解除後のタグ定義
     */
    public static String unescapeTag( String value, String tag) {
        return unescapeTag( value, tag, TagParser.TAG_PARAM_PREFIX);
    }

    /**
     * タグ定義のエスケープを解除する
     *
     * @param value タグ定義
     * @param tag 対象タグ（先頭のバックスラッシュはエスケープ文字として予約）
     * @param tagParamPrefix タグのパラメータ開始文字（null不可。空文字の場合は一致しない）
     * @return エスケープ解除後のタグ定義
     */
    public static String unescapeTag( String value, String tag, String tagParamPrefix) {
        if ( isEscapedTag( value, tag, tagParamPrefix)) {
            return value.substring( TagParser.TAG_ESCAPE.length());
        }
        return value;
    }

    private static int getLeadingEscapeLength( String value) {
        int escapeLength = 0;
        while ( value.startsWith( TagParser.TAG_ESCAPE, escapeLength)) {
            escapeLength += TagParser.TAG_ESCAPE.length();
        }
        return escapeLength;
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
