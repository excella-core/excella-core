/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: MapParser.java 128 2009-07-02 06:32:17Z yuta-takahashi $
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

import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;

/**
 * パース結果をマップで返却するパーサ
 * 
 * @since 1.0
 */
public class MapParser extends TagParser<Map<?, ?>> {

    /**
     * セル位置定義区切り文字
     */
    protected static final String PARAM_CELL_DELIM = ":";

    /**
     * 区切り文字前半用インデックス
     */
    protected static final int SPLIT_FIRST_INDEX = 0;

    /**
     * 区切り文字後半用インデックス
     */
    protected static final int SPLIT_LAST_INDEX = 1;

    /**
     * データ行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_FROM = "DataRowFrom";

    /**
     * データ行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_TO = "DataRowTo";

    /**
     * キー列の調整パラメータ
     */
    protected static final String PARAM_KEY_COLUMN = "KeyColumn";

    /**
     * 値列の調整パラメータ
     */
    protected static final String PARAM_VALUE_COLUMN = "ValueColumn";

    /**
     * キーセルの調整パラメータ
     */
    protected static final String PARAM_KEY_CELL = "KeyCell";

    /**
     * 固定キーの値パラメータ
     */
    protected static final String PARAM_KEY = "Key";

    /**
     * 固定値の値パラメータ
     */
    protected static final String PARAM_VALUE = "Value";

    /**
     * 値セルの調整パラメータ
     */
    protected static final String PARAM_VALUE_CELL = "ValueCell";

    /**
     * デフォルトデータ開始行調整値
     */
    protected static final int DEFAULT_VALUE_ROW_FROM_ADJUST = 1;

    /**
     * デフォルトキー列調整値
     */
    protected static final int DEFAULT_KEY_COLUMN_ADJUST = 0;

    /**
     * デフォルト値列調整値
     */
    protected static final int DEFAULT_VALUE_COLUMN_ADJUST = 1;

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public MapParser( String tag) {
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
    public Map<?, ?> parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        // タグ行
        int tagRowIdx = tagCell.getRowIndex();
        // タグ列
        int tagColIdx = tagCell.getColumnIndex();
        // データ開始行
        int valueRowFromIdx;
        // データ終了行
        int valueRowToIdx = sheet.getLastRowNum();

        // キー列
        int keyColIdx = 0;
        // 値列
        int valueColIdx = 0;

        // キータグ有無フラグ
        boolean keyTagFlag;
        // 値タグ有無フラグ
        boolean valueTagFlag;

        // キー
        String defKey = null;
        // 値
        String defValue = null;

        // キーセルタグ有無フラグ
        boolean keyCellTagFlag;
        // 値セルタグ有無フラグ
        boolean valueCellTagFlag;

        // キー行
        int keyRowIdx = 0;
        // 値行
        int valueRowIdx = 0;

        try {
            Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

            // パラメータチェック
            checkParam( paramDef, tagCell);

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

            keyTagFlag = paramDef.containsKey( PARAM_KEY);
            keyCellTagFlag = paramDef.containsKey( PARAM_KEY_CELL);

            if ( keyTagFlag) {
                // キータグがある場合
                defKey = paramDef.get( PARAM_KEY);

            } else if ( keyCellTagFlag) {
                // キーセルタグがある場合
                String value = paramDef.get( PARAM_KEY_CELL);

                // キー行番号の調整
                keyRowIdx = tagRowIdx + Integer.valueOf( value.split( PARAM_CELL_DELIM)[SPLIT_FIRST_INDEX]);
                if ( keyRowIdx < 0 || keyRowIdx > sheet.getLastRowNum()) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_KEY_CELL);
                }
                // キー列番号の調整
                keyColIdx = tagColIdx + Integer.valueOf( value.split( PARAM_CELL_DELIM)[SPLIT_LAST_INDEX]);
                if ( keyColIdx > PoiUtil.getLastColNum( sheet) || keyColIdx < 0) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_KEY_CELL);
                }

            } else {
                // それ以外の場合
                // キー列番号の調整
                keyColIdx = TagUtil.adjustValue( tagColIdx, paramDef, PARAM_KEY_COLUMN, DEFAULT_KEY_COLUMN_ADJUST);
                if ( keyColIdx > PoiUtil.getLastColNum( sheet) || keyColIdx < 0) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_KEY_COLUMN);
                }
            }

            valueTagFlag = paramDef.containsKey( PARAM_VALUE);
            valueCellTagFlag = paramDef.containsKey( PARAM_VALUE_CELL);

            if ( valueTagFlag) {
                // 値タグがある場合
                defValue = paramDef.get( PARAM_VALUE);

            } else if ( valueCellTagFlag) {
                // 値セルタグがある場合
                String value = paramDef.get( PARAM_VALUE_CELL);

                // 値行番号の調整
                valueRowIdx = tagRowIdx + Integer.valueOf( value.split( PARAM_CELL_DELIM)[SPLIT_FIRST_INDEX]);
                if ( valueRowIdx < 0 || valueRowIdx > sheet.getLastRowNum()) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_VALUE_CELL);
                }

                // 値列番号の調整
                valueColIdx = tagColIdx + Integer.valueOf( value.split( PARAM_CELL_DELIM)[SPLIT_LAST_INDEX]);
                if ( valueColIdx > PoiUtil.getLastColNum( sheet) || valueColIdx < 0) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_VALUE_CELL);
                }

            } else {
                // それ以外の場合
                // 値列番号の調整
                valueColIdx = TagUtil.adjustValue( tagColIdx, paramDef, PARAM_VALUE_COLUMN, DEFAULT_VALUE_COLUMN_ADJUST);
                if ( valueColIdx > PoiUtil.getLastColNum( sheet) || valueColIdx < 0) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_VALUE_COLUMN);
                }
            }

        } catch ( Exception e) {
            if ( e instanceof ParseException) {
                throw ( ParseException) e;
            } else {
                throw new ParseException( tagCell, e);
            }
        }

        // パース処理
        Map<Object, Object> results = new LinkedHashMap<Object, Object>();
        Row keyRow = null;
        Row valueRow = null;
        for ( int rowCnt = valueRowFromIdx; rowCnt <= valueRowToIdx; rowCnt++) {

            Object key;
            Object value;

            if ( keyTagFlag) {
                // キータグがある場合
                key = defKey;

            } else {
                // キータグがない場合

                // キー行の取得
                if ( keyCellTagFlag) {
                    // キーセルタグがある場合
                    keyRow = sheet.getRow( keyRowIdx);
                } else {
                    // キーセルタグがない場合
                    keyRow = sheet.getRow( rowCnt);
                }

                if ( keyRow == null) {
                    continue;
                }
                key = PoiUtil.getCellValue( keyRow.getCell( keyColIdx));
            }

            if ( valueTagFlag) {
                // 値タグがある場合
                value = defValue;

            } else {
                // 値タグがない場合

                // 値行の取得
                if ( valueCellTagFlag) {
                    // 値セルタグがある場合
                    valueRow = sheet.getRow( valueRowIdx);
                } else {
                    // 値セルタグがない場合
                    valueRow = sheet.getRow( rowCnt);
                }
                if ( valueRow != null) {
                    value = PoiUtil.getCellValue( valueRow.getCell( valueColIdx));
                } else {
                    value = null;
                }
            }
            results.put( key, value);
        }
        return results;
    }

    /**
     * 不正なパラメータがある場合、ParseExceptionをthrowする。
     * 
     * @param paramDef パラメータマップ
     * @param tagCell タグのあるセル
     * @throws ParseException パース例外
     */
    private void checkParam( Map<String, String> paramDef, Cell tagCell) throws ParseException {
        // キー
        if ( paramDef.containsKey( PARAM_KEY) && paramDef.containsKey( PARAM_KEY_COLUMN)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_KEY + "," + PARAM_KEY_COLUMN);
        } else if ( paramDef.containsKey( PARAM_KEY) && paramDef.containsKey( PARAM_KEY_CELL)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_KEY + "," + PARAM_KEY_CELL);
        } else if ( paramDef.containsKey( PARAM_KEY_COLUMN) && paramDef.containsKey( PARAM_KEY_CELL)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_KEY_COLUMN + "," + PARAM_KEY_CELL);
        }

        // 値
        if ( paramDef.containsKey( PARAM_VALUE) && paramDef.containsKey( PARAM_VALUE_COLUMN)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_VALUE + "," + PARAM_VALUE_COLUMN);
        } else if ( paramDef.containsKey( PARAM_VALUE) && paramDef.containsKey( PARAM_VALUE_CELL)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_VALUE + "," + PARAM_VALUE_CELL);
        } else if ( paramDef.containsKey( PARAM_VALUE_COLUMN) && paramDef.containsKey( PARAM_VALUE_CELL)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_VALUE_COLUMN + "," + PARAM_VALUE_CELL);
        }
    }
}
