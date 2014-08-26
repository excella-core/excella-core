/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: BookData.java 40 2009-05-20 08:08:45Z yuta-takahashi $
 * $Revision: 40 $
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
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * ワークブックの解析結果を保持するクラス
 * 
 * @since 1.0
 */
public class BookData {
    /**
     * 解析結果(シート名、シートデータ)を保持するマップ
     */
    private Map<String, SheetData> sheetDataMap = new LinkedHashMap<String, SheetData>();

    /**
     * 追加された順序を保持したシート名の一覧
     */
    private List<String> sheetNameList = new ArrayList<String>();

    /**
     * シートデータが存在するかどうかのチェック
     * 
     * @param sheetName シート名
     * @return 存在する場合はtrue,存在しない場合はfalse;
     */
    public boolean containsSheet( String sheetName) {
        return sheetDataMap.containsKey( sheetName);
    }

    /**
     * シートデータの取得
     * 
     * @param sheetName シート名
     * @return シートデータ
     */
    public SheetData getSheetData( String sheetName) {
        return sheetDataMap.get( sheetName);
    }

    /**
     * 含まれるシート名の一覧取得
     * 
     * @return 含まれるシート名の一覧
     */
    public List<String> getSheetNames() {
        return sheetNameList;
    }

    /**
     * シートデータの設定
     * 
     * @param sheetName シート名
     * @param sheetData シートデータ
     * @return シートデータ
     */
    public SheetData putSheetData( String sheetName, SheetData sheetData) {
        // 既にリストに存在すれば削除して一旦追加
        if ( sheetNameList.contains( sheetName)) {
            sheetNameList.remove( sheetName);
        }
        sheetNameList.add( sheetName);
        return sheetDataMap.put( sheetName, sheetData);
    }

    /**
     * 含まれるシートデータの一覧取得
     * 
     * @return 含まれるシートデータの一覧
     */
    public Collection<SheetData> getSheetDatas() {
        return sheetDataMap.values();
    }

    /**
     * 全シートデータのクリア
     */
    public void clear() {
        sheetDataMap.clear();
    }
}
