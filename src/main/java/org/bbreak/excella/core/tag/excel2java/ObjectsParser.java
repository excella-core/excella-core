/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ObjectsParser.java 128 2009-07-02 06:32:17Z yuta-takahashi $
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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.BookController;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;

/**
 * パース結果をList&lt;Object&gt;で返却するパーサ
 * 
 * @since 1.0
 */
public class ObjectsParser extends TagParser<List<Object>> {

    /**
     * クラス定義パラメータ
     */
    protected static final String PARAM_CLASS = "Class";

    /**
     * プロパティ行の調整パラメータ
     */
    protected static final String PARAM_PROPERTY_ROW = "PropertyRow";

    /**
     * データ開始行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_FROM = "DataRowFrom";

    /**
     * データ終了行の調整パラメータ
     */
    protected static final String PARAM_DATA_ROW_TO = "DataRowTo";

    /**
     * デフォルトプロパティ行調整値
     */
    protected static final int DEFAULT_PROPERTY_ROW_ADJUST = 1;

    /**
     * デフォルトデータ開始行調整値
     */
    protected static final int DEFAULT_VALUE_ROW_FROM_ADJUST = 2;

    /**
     * カスタムプロパティのリスト 
     */
    private List<ObjectsPropertyParser> customPropertyParsers = new ArrayList<ObjectsPropertyParser>();

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ObjectsParser( String tag) {
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
    public List<Object> parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        List<Object> resultList = new ArrayList<Object>();
        Class<?> clazz = null;

        // タグ行
        int tagRowIdx = tagCell.getRowIndex();
        // プロパティ行
        int propertyRowIdx;
        // データ開始行
        int valueRowFromIdx;
        // データ終了行
        int valueRowToIdx = sheet.getLastRowNum();

        try {
            Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

            clazz = Class.forName( paramDef.get( PARAM_CLASS));

            // プロパティ行の調整
            propertyRowIdx = TagUtil.adjustValue( tagRowIdx, paramDef, PARAM_PROPERTY_ROW, DEFAULT_PROPERTY_ROW_ADJUST);
            if ( propertyRowIdx < 0 || propertyRowIdx > sheet.getLastRowNum()) {
                throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_PROPERTY_ROW);
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

        // プロパティと対応するデータ型のマップ
        Map<Integer, Class<?>> propertyClassMap = new HashMap<Integer, Class<?>>();
        // カラムインデックスとプロパティ名のマップ
        Map<Integer, String> propertyNameMap = new HashMap<Integer, String>();
        // プロパティとカスタムパーサのマップ
        Map<String, List<ObjectsPropertyParser>> customPropertyParserMap = new HashMap<String, List<ObjectsPropertyParser>>();

        // プロパティ行の解析
        List<Integer> targetColNums = new ArrayList<Integer>();
        Row propertyRow = sheet.getRow( propertyRowIdx);
        if ( propertyRow == null) {
            // プロパティ行がnullの場合
            return resultList;
        }
        int firstCellNum = propertyRow.getFirstCellNum();
        int lastCellNum = propertyRow.getLastCellNum();
        for ( int cellCnt = firstCellNum; cellCnt < lastCellNum; cellCnt++) {
            Cell cell = propertyRow.getCell( cellCnt);
            if ( cell == null) {
                continue;
            }
            try {
                String propertyName = cell.getStringCellValue();
                if ( propertyName.startsWith( BookController.COMMENT_PREFIX)) {
                    continue;
                }

                Object obj = clazz.newInstance();
                Class<?> propertyClass = PropertyUtils.getPropertyType( obj, propertyName);
                if ( propertyClass != null) {
                    propertyClassMap.put( cellCnt, propertyClass);
                    propertyNameMap.put( cellCnt, propertyName);
                    targetColNums.add( cellCnt);
                } else {
                    // プロパティから処理対象のパーサを取得
                    for ( ObjectsPropertyParser parser : customPropertyParsers) {
                        if ( parser.isParse( sheet, cell)) {
                            List<ObjectsPropertyParser> propertyParsers = customPropertyParserMap.get( propertyName);
                            if ( propertyParsers == null) {
                                propertyParsers = new ArrayList<ObjectsPropertyParser>();
                            }
                            // 複数回同じパーサが呼ばれる不具合の修正
                            if ( !propertyParsers.contains( parser)) {
                                propertyParsers.add( parser);                                
                            }
                            customPropertyParserMap.put( propertyName, propertyParsers);

                            if ( !targetColNums.contains( cellCnt)) {
                                propertyNameMap.put( cellCnt, propertyName);
                                targetColNums.add( cellCnt);
                            }
                        }
                    }
                }

            } catch ( Exception e) {
                throw new ParseException( cell, e);
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
                Object obj;
                try {
                    obj = clazz.newInstance();
                    for ( Integer colCnt : targetColNums) {
                        Cell cell = dataRow.getCell( colCnt);

                        try {
                            Class<?> propertyClass = propertyClassMap.get( colCnt);
                            String propertyName = propertyNameMap.get( colCnt);
                            // カスタムプロパティの判定
                            if ( customPropertyParserMap.containsKey( propertyName)) {
                                List<ObjectsPropertyParser> propertyParsers = customPropertyParserMap.get( propertyName);
                                Map<String, String> params = TagUtil.getParams( propertyName);
                                Object cellValue = PoiUtil.getCellValue( cell);

                                // カスタムパーサの処理実行
                                for ( ObjectsPropertyParser propertyParser : propertyParsers) {
                                    propertyParser.parse( obj, cellValue, TagUtil.getTag( propertyName), params);
                                }
                            } else {
                                Object value = null;
                                if ( cell != null) {
                                    value = PoiUtil.getCellValue( cell, propertyClass);
                                }
                                PropertyUtils.setProperty( obj, propertyName, value);
                            }
                        } catch ( Exception e) {
                            throw new ParseException( cell, e);
                        }
                    }
                } catch ( Exception e) {
                    if ( e instanceof ParseException) {
                        throw ( ParseException) e;
                    } else {
                        throw new ParseException( tagCell, e);
                    }
                }
                resultList.add( obj);
            }
        }
        return resultList;
    }

    /**
     * カスタムプロパティ解析クラスの追加
     * 
     * @param parser 追加するカスタムプロパティ解析クラス
     */
    public void addPropertyParser( ObjectsPropertyParser parser) {
        customPropertyParsers.add( parser);
    }

    /**
     * カスタムプロパティ解析クラスの削除
     * 
     * @param parser 削除するカスタムプロパティ解析クラス
     */
    public void removePropertyParser( ObjectsPropertyParser parser) {
        customPropertyParsers.remove( parser);
    }

    /**
     * カスタムプロパティ解析クラスを全削除する
     */
    public void clearPropertyParsers() {
        customPropertyParsers.clear();
    }
}
