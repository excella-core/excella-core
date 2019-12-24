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

package org.bbreak.excella.core.tag;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.TagUtil;

/**
 * タグ処理のスーパークラス
 * 
 * @since 1.0
 *
 * @param <RESULT> 処理結果の型
 */
public abstract class TagParser<RESULT> {
	
	/**
     * パラメータ定義の開始文字
     */
    public static final String TAG_PARAM_PREFIX = "{";

    /**
     * パラメータ定義の終了文字
     */
    public static final String TAG_PARAM_SUFFIX = "}";

    /**
     * パラメータ区切り文字
     */
    public static final String PARAM_DELIM = ",";

    /**
     * キー、値の区切り文字
     */
    public static final String VALUE_DELIM = "=";

    /**
     * このパーサで処理するタグ
     */
    private String tag;

    /**
     * コンストラクタ
     * 
     * @param tag 対象タグ
     */
    public TagParser( String tag) {
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
        // 文字列かつ、タグを含むセルの場合は処理対象
        if ( tagCell.getCellType() == CellType.STRING) {
            String cellTag = TagUtil.getTag( tagCell.getStringCellValue());
            if ( tag.equals( cellTag)) {
                return true;
            }
        }
        return false;
    }

    /**
     * パース処理
     * 
     * @param sheet 対象シート
     * @param tagCell タグが定義されたセル
     * @param data BookControllerのparseBook(), parseSheet()メソッド、 
     *              SheetParserのparseSheetメソッドで引数を渡した場合に
     *              TagParserまで引き継がれる処理データ
     * @return パース結果
     * @throws ParseException パース例外
     */
    public abstract RESULT parse( Sheet sheet, Cell tagCell, Object data) throws ParseException;

    /**
     * 対象タグの取得
     * 
     * @return 対象タグ
     */
    public String getTag() {
        return tag;
    }

    /**
     * 対象タグの設定
     * 
     * @param tag 対象タグ
     */
    public void setTag( String tag) {
        this.tag = tag;
    }
}
