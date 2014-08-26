/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: SheetData.java 156 2010-09-09 04:16:50Z akira-yokoi $
 * $Revision: 156 $
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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * タグ単位でシートの解析結果を保持するクラス
 * 
 * @since 1.0
 */
public class SheetData {

    /**
     * シート名
     */
    private String sheetName = null;

    /**
     * タグ毎にデータを保持するマップ
     */
    private Map<String, List<Object>> tagValueMap = new HashMap<String, List<Object>>();

    /**
     * 追加された順序を保持したキーの一覧
     */
    private List<String> keyList = new ArrayList<String>();

    /**
     * コンストラクタ
     * 
     * @param sheetName シート名
     */
    public SheetData( String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * シート名の取得
     * 
     * @return シート名
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * シート名の設定
     * 
     * @param sheetName シート名
     */
    public void setSheetName( String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * タグに対応する解析結果の追加
     * 
     * @param tagName タグ名
     * @param result タグに対応する解析結果
     */
    public void put( String tagName, Object result) {
        // 既にリストに存在すれば削除して一旦追加
        if ( keyList.contains( tagName)) {
            keyList.remove( tagName);
        }
        keyList.add( tagName);

        List<Object> values = tagValueMap.get( tagName);
        if( values == null){
        	values = new ArrayList<Object>();
        	tagValueMap.put( tagName, values);
        }
        values.add( result);
    }

    /**
     * タグに対応する解析結果の取得
     * 
     * @param tagName タグ名
     * @return タグに対応する解析結果
     */
    public Object get( String tagName) {
    	List<Object> results = tagValueMap.get( tagName);
    	if( results != null && ! results.isEmpty()){
    		// 最終要素を返す(下位互換)
    		return results.get( results.size() - 1);
    	}
        return null;
    }
    
    /**
     * タグに対応する解析結果の取得
     * 
     * @param tagName タグ名
     * @return タグに対応する解析結果
     */
    public List<Object> getList( String tagName) {
        return tagValueMap.get( tagName);
    }

    /**
     * 保持するタグ名の一覧取得
     * 
     * @return 保持するタグ名の一覧
     */
    public Set<String> getTagNames() {
        return tagValueMap.keySet();
    }

    /**
     * タグが存在するかどうかのチェック
     * 
     * @param tagName 対象のタグ名
     * @return 存在する場合はtrue,存在しない場合はfalse
     */
    public boolean containsTag( String tagName) {
        return tagValueMap.containsKey( tagName);
    }

    /**
     * putされた順のキーの一覧取得
     * 
     * @return キーの一覧取得
     */
    public List<String> getKeyList() {
        return keyList;
    }

    /**
     * データの削除
     * 
     * @param key 削除対象のキー
     * @return 削除対象のデータ
     */
    public Object remove( Object key) {
        keyList.remove( key);
        return tagValueMap.remove( key);
    }

    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append( "=====================" + sheetName + "=====================");
        List<String> keys = getKeyList();
        for ( String key : keys) {
            builder.append( "\n");
            builder.append( "\t");
            builder.append( key);
            builder.append( "\t");
            builder.append( get( key));
        }
        return builder.toString();
    }
}
