/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: MapsParser.java 128 2009-07-02 06:32:17Z yuta-takahashi $
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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;

/**
 * パース結果をList&lt;Map&gt;で返却するパーサ
 * 
 * @since 1.0
 */
public class MapsParser extends TagParser<List<Map<?, ?>>> {

    /**
     * キー行の調整パラメータ
     */
    protected static final String PARAM_KEY_ROW = "KeyRow";

    /**
     * データ開始行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_FROM = "DataRowFrom";

    /**
     * データ終了行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_TO = "DataRowTo";

    /**
     * デフォルトキー行調整値
     */
    protected static final int DEFAULT_KEY_ROW_ADJUST = 1;

    /**
     * デフォルトデータ開始行調整値
     */
    protected static final int DEFAULT_VALUE_ROW_FROM_ADJUST = 2;

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public MapsParser( String tag) {
        super( tag);
    }

    /**
     * パース処理
     * 
     * @param sheet 対象シート
     * @param tagCell タグが定義されたセル
     * @param data BookControllerのparseBook(), parseSheet()メソッド、<BR>
     * SheetParserのparseSheetメソッドで引数を渡した場合に<BR>
     * TagParserまで引き継がれる処理データ<BR>
     * @return パース結果
     * @throws ParseException パース例外
     */
    @Override
    public List<Map<?, ?>> parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        List<Map<?, ?>> resultList = new ArrayList<Map<?, ?>>();

        // タグ行
        int tagRowIdx = tagCell.getRowIndex();
        // キー行
        int keyRowIdx;
        // データ開始行
        int valueRowFromIdx;
        // データ終了行
        int valueRowToIdx = sheet.getLastRowNum();

        try {
            Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

            // キー行の調整
            keyRowIdx = TagUtil.adjustValue( tagRowIdx, paramDef, PARAM_KEY_ROW, DEFAULT_KEY_ROW_ADJUST);
            if ( keyRowIdx < 0 || keyRowIdx > sheet.getLastRowNum()) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_KEY_ROW);
            }

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

        } catch ( Exception e) {
            if ( e instanceof ParseException) {
                throw ( ParseException) e;
            } else {
                throw new ParseException( tagCell, e);
            }
        }

        // キー行の解析
        List<Integer> targetColNums = new ArrayList<Integer>();
        Row keyRow = sheet.getRow( keyRowIdx);
        if ( keyRow == null) {
            // キー行がnullの場合
            return resultList;
        }
        int firstCellNum = keyRow.getFirstCellNum();
        int lastCellNum = keyRow.getLastCellNum();
        for ( int cellCnt = firstCellNum; cellCnt < lastCellNum; cellCnt++) {
            Cell cell = keyRow.getCell( cellCnt);
            Object cellValue = PoiUtil.getCellValue( cell);
            if ( cellValue instanceof String) {
                String keyName = ( String) cellValue;
                if ( keyName.startsWith( BookController.COMMENT_PREFIX)) {
                    continue;
                }
            }
            if ( cellValue != null) {
                targetColNums.add( cellCnt);
            }
        }

        if ( targetColNums.size() > 0) {
            // 処理対象列がある場合

            // データの解析
            for ( int rowCnt = valueRowFromIdx; rowCnt <= valueRowToIdx; rowCnt++) {
                Row dataRow = sheet.getRow( rowCnt);
                if ( dataRow == null) {
                    continue;
                }
                Map<Object, Object> map = new LinkedHashMap<Object, Object>();
                for ( Integer colCnt : targetColNums) {
                    Cell keyCell = keyRow.getCell( colCnt);
                    Cell valueCell = dataRow.getCell( colCnt);

                    Object key = PoiUtil.getCellValue( keyCell);
                    Object value = PoiUtil.getCellValue( valueCell);

                    map.put( key, value);
                }
                resultList.add( map);
            }
        }

        return resultList;
    }
}
