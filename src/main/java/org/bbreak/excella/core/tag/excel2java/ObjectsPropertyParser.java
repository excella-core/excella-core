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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;

/**
 * Objectsパーサ用独自プロパティ解析クラス
 * 
 * @since 1.0
 */
public abstract class ObjectsPropertyParser {

    /**
     * このプロセッサで処理するタグ
     */
    private String tag;

    /**
     * コンストラクタ
     * @param tag タグ
     */
    public ObjectsPropertyParser( String tag) {
        this.tag = tag;
    }

    /**
     * タグを取得する
     * 
     * @return tag タグ
     */
    public String getTag() {
        return tag;
    }

    /**
     * タグを設定する
     * 
     * @param tag タグ
     */
    public void setTag( String tag) {
        this.tag = tag;
    }

    /**
     * パース処理を行うか否かの判定
     * 
     * @param sheet 対象シート
     * @param tagCell 対象セル
     * @return 処理対象の場合はTrue、処理対象外の場合はFalse
     * @throws ParseException
     */
    public boolean isParse( Sheet sheet, Cell tagCell) throws ParseException {
        if ( tagCell == null) {
            return false;
        }

        // 文字列かつ、タグを含むセルの場合は処理対象
        if ( tagCell.getCellTypeEnum() == CellType.STRING) {
            if ( tagCell.getStringCellValue().contains( tag)) {
                return true;
            }
        }
        return false;
    }

    /**
     * パース処理を実行する
     * 
     * @param object 対象オブジェクト
     * @param cellValue セルの値
     * @param tag タグ
     * @param params パラメータのマップ
     * @throws ParseException
     */
    public abstract void parse( Object object, Object cellValue, String tag, Map<String, String> params) throws ParseException;
}
