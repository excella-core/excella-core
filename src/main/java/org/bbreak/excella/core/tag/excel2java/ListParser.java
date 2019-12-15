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

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;

/**
 * 処理結果をListで返却するパーサ
 * 
 * @since 1.0
 */
public class ListParser extends TagParser<List<?>> {

    /**
     * データ行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_FROM = "DataRowFrom";

    /**
     * データ行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_TO = "DataRowTo";

    /**
     * 値列の調整パラメータ
     */
    protected static final String PARAM_VALUE_COLUMN = "ValueColumn";

    /**
     * デフォルトデータ開始行調整値
     */
    protected static final int DEFAULT_VALUE_ROW_FROM_ADJUST = 1;

    /**
     * デフォルトValue列調整値
     */
    protected static final int DEFAULT_VALUE_COLUMN_ADJUST = 0;

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ListParser( String tag) {
        super( tag);
    }

    /**
     * パース処理
     * 
     * @param sheet 対象シート
     * @param tagCell タグが定義されたセル
     * @param data BookControllerのparseBook(), parseSheet()メソッド、<BR>
     *              SheetParserのparseSheetメソッドで引数を渡した場合に<BR>
     *              TagParserまで引き継がれる処理データ<BR>
     * @return パース結果
     * @throws ParseException パース例外
     */
    @Override
    public List<?> parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        // パース結果格納リスト
        List<Object> results = new ArrayList<Object>();

        try {
            // データ範囲
            int tagRowIdx = tagCell.getRowIndex();
            int valueRowFromIdx;
            int valueRowToIdx = sheet.getLastRowNum();
            int valueColIdx = tagCell.getColumnIndex();

            Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

            // データ開始行の調整
            valueRowFromIdx = TagUtil.adjustValue( tagRowIdx, paramDef, PARAM_DATA_ROW_FROM, DEFAULT_VALUE_ROW_FROM_ADJUST);
            if ( valueRowFromIdx < 0 || valueRowFromIdx > sheet.getLastRowNum()) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_DATA_ROW_FROM);
            }

            // データ終了行の調整
            valueRowToIdx = TagUtil.adjustValue( tagRowIdx, paramDef, PARAM_DATA_ROW_TO, valueRowToIdx - tagRowIdx);
            if ( valueRowToIdx > sheet.getLastRowNum() || valueRowToIdx < 0) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_DATA_ROW_TO);
            }

            // データ開始行と終了行の関係チェック
            if ( valueRowFromIdx > valueRowToIdx) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_DATA_ROW_FROM + "," + PARAM_DATA_ROW_TO);
            }

            // Value列番号の調整
            valueColIdx = TagUtil.adjustValue( valueColIdx, paramDef, PARAM_VALUE_COLUMN, DEFAULT_VALUE_COLUMN_ADJUST);
            if ( valueColIdx > PoiUtil.getLastColNum( sheet) || valueColIdx < 0) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_VALUE_COLUMN);
            }

            // パース処理
            for ( int rowCnt = valueRowFromIdx; rowCnt <= valueRowToIdx; rowCnt++) {
                Row row = sheet.getRow( rowCnt);
                if ( row != null) {
                    Object value = PoiUtil.getCellValue( row.getCell( valueColIdx));
                    results.add( value);
                }
            }

        } catch ( Exception e) {
            if ( e instanceof ParseException) {
                throw ( ParseException) e;
            } else {
                throw new ParseException( tagCell, e);
            }
        }

        return results;
    }
}
