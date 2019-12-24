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
