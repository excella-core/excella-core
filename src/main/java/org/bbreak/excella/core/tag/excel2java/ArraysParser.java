/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ArraysParser.java 128 2009-07-02 06:32:17Z yuta-takahashi $
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
 * パース結果をList&lt;Object[]&gt;で返却するパーサ
 * 
 * @since 1.0
 */
public class ArraysParser extends TagParser<List<Object[]>> {

    /**
     * データ開始行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_FROM = "DataRowFrom";

    /**
     * データ終了行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_TO = "DataRowTo";

    /**
     * データ開始列の調整パラメータ
     */
    protected static final String PARAM_DATA_CLOMUN_FROM = "DataColumnFrom";

    /**
     * データ開始列の調整パラメータ
     */
    protected static final String PARAM_DATA_CLOMUN_TO = "DataColumnTo";

    /**
     * デフォルトデータ開始行調整値
     */
    protected static final int DEFAULT_VALUE_ROW_FROM_ADJUST = 1;

    /**
     * デフォルトデータ終了行調整値
     */
    protected static final int DEFAULT_VALUE_COLUMN_FROM_ADJUST = 0;

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ArraysParser( String tag) {
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
    public List<Object[]> parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        List<Object[]> resultList = new ArrayList<Object[]>();

        // タグ行
        int tagRowIdx = tagCell.getRowIndex();
        // タグ列
        int tagColIdx = tagCell.getColumnIndex();
        // データ開始行
        int valueRowFromIdx;
        // データ終了行
        int valueRowToIdx = sheet.getLastRowNum();
        // データ開始列
        int valueColumnFromIdx;
        // データ終了列
        int valueColumnToIdx = 0;

        // データ終了列指定フラグ
        boolean valueColumnToFlag = false;

        try {

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

            // データ開始列の調整
            valueColumnFromIdx = TagUtil.adjustValue( tagColIdx, paramDef, PARAM_DATA_CLOMUN_FROM, DEFAULT_VALUE_COLUMN_FROM_ADJUST);
            if ( valueColumnFromIdx < 0 || valueColumnFromIdx > PoiUtil.getLastColNum( sheet)) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_DATA_CLOMUN_FROM);
            }

            // データ終了列指定フラグ
            valueColumnToFlag = paramDef.containsKey( PARAM_DATA_CLOMUN_TO);
            if ( valueColumnToFlag) {
                // データ終了列の調整
                valueColumnToIdx = tagColIdx + Integer.valueOf( paramDef.get( PARAM_DATA_CLOMUN_TO));
                if ( valueColumnToIdx < 0 || valueColumnToIdx > PoiUtil.getLastColNum( sheet)) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_DATA_CLOMUN_TO);
                }

                // データ開始行と終了行の関係チェック
                if ( valueColumnFromIdx > valueColumnToIdx) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_DATA_CLOMUN_FROM + "," + PARAM_DATA_CLOMUN_TO);
                }
            }
        } catch ( Exception e) {
            if ( e instanceof ParseException) {
                throw ( ParseException) e;
            } else {
                throw new ParseException( tagCell, e);
            }
        }

        // データ行の解析
        for ( int rowCnt = valueRowFromIdx; rowCnt <= valueRowToIdx; rowCnt++) {

            List<Object> objList = new ArrayList<Object>();

            Row row = sheet.getRow( rowCnt);
            if ( row == null) {
                // 行がnullの場合
                continue;
            }
            if ( !valueColumnToFlag) {
                // データ終了列の指定がない場合
                // 各行の終了列をデータ終了列とする
                valueColumnToIdx = row.getLastCellNum() - 1;
            }
            for ( int cellCnt = valueColumnFromIdx; cellCnt <= valueColumnToIdx; cellCnt++) {

                Cell cell = row.getCell( cellCnt);
                Object cellValue = PoiUtil.getCellValue( cell);
                objList.add( cellValue);
            }

            resultList.add( objList.toArray());
        }
        return resultList;
    }
}
