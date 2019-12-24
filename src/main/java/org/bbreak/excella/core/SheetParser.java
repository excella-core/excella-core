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
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;

/**
 * シートの解析を行うクラス
 * 
 * タグを検査して、一致するタグが存在した場合は対応するタグパーサを呼び出し、
 * 結果をSheetDataに設定する
 * 
 * タグの検査は行列方向([1,A] [1,B]・・・、[2,A] [2,B]・・・)で行う
 * 
 * タグにLastTag=Trueのパラメータが設定されていた場合はそのタグを処理して終了する。
 * 
 * @since 1.0
 */
public class SheetParser {
	
	/** 結果キーパラメータ */
    protected static final String PARAM_RESULT_KEY = "ResultKey";

    /** 最終タグパラメータ */
    protected static final String PARAM_LAST_TAG = "LastTag";

    /** ログ */
    private static Log log = LogFactory.getLog( SheetParser.class);

    /** タグパーサのリスト */
    private List<TagParser<?>> tagParsers = new ArrayList<TagParser<?>>();

    /**
     * シートの解析
     * 
     * @param sheet 解析対象シート
     * @param data BookControllerのparseBook(), parseSheet()メソッド、
     *              SheetParserのparseSheetメソッドで引数を渡した場合に
     *              TagParserまで引き継がれる処理データ
     * @return 解析結果
     * @throws ParseException 解析に失敗した場合にThrowされる
     */
    public SheetData parseSheet( Sheet sheet, Object data) throws ParseException {
        // 解析結果
        String sheetName = PoiUtil.getSheetName( sheet);
        SheetData sheetData = new SheetData( sheetName);

        int firstRowNum = sheet.getFirstRowNum();

        // タグセルの検査
        for ( int rowCnt = firstRowNum; rowCnt <= sheet.getLastRowNum(); rowCnt++) {
            // 行の取得
            Row row = sheet.getRow( rowCnt);
            if ( row == null) {
                continue;
            }
            if ( row != null) {
                for ( int columnIdx = 0; columnIdx < row.getLastCellNum(); columnIdx++) {
                    // セルの取得
                    Cell cell = row.getCell( columnIdx);
                    if ( cell == null) {
                        continue;
                    }
                    if( parseCell(sheet, data, sheetData, cell, row, columnIdx)) {
                    	// 終了タグによる終了
                    	return sheetData;
                    }
                }
            }
        }

        return sheetData;
    }

    /**
     * セルの解析
     * 
     * @param sheet 解析対象シート
     * @param data 処理データ
     * @param sheetData 返却データ
     * @param cell 解析対象セル
     * @param row 解析対象行
     * @param columnIdx 解析対象列番号
     * @return 終了タグによる終了かどうかを返す
     * @throws ParseException 解析に失敗した場合にThrowされる
     */
    private boolean parseCell( Sheet sheet, Object data, SheetData sheetData, Cell cell, Row row, int columnIdx) throws ParseException {
        for ( TagParser<?> parser : tagParsers) {
            // 処理判定チェック
            if ( parser.isParse( sheet, cell)) {
                String strCellValue = cell.getStringCellValue();
                Map<String, String> paramDef = TagUtil.getParams( strCellValue);
                // 処理実行
                Object result = parser.parse( sheet, cell, data);

                // 結果の追加
                if ( result != null) {
                    // パラメータにResultKeyが指定されていればResultKeyを使用する。
                    // 未指定の場合はTag名を利用する。
                    String resultKey = parser.getTag();
                    if ( paramDef.containsKey( PARAM_RESULT_KEY)) {
                        resultKey = paramDef.get( PARAM_RESULT_KEY);
                    }
                    sheetData.put( resultKey, result);
                }
                // パース結果のロギング
                if ( log.isInfoEnabled()) {
                    StringBuilder resultBuf = new StringBuilder( strCellValue + "の処理結果:");
                    if ( result instanceof Map) {
                        Map<?, ?> mapResult = ( Map<?, ?>) result;
                        Set<?> keyset = mapResult.keySet();
                        for ( Object key : keyset) {
                            resultBuf.append( "[" + key + ":" + mapResult.get( key) + "]");
                        }
                    } else if ( result instanceof Collection) {
                        Collection<?> listResult = ( Collection<?>) result;
                        resultBuf.append( listResult.getClass() + " Size=" + listResult.size());
                    } else {
                        resultBuf.append( result);
                    }
                    log.info( resultBuf.toString());
                }

                // 大量データを考慮して最終タグ判定。最終タグの場合はそのタグでシートの処理は抜ける

                if ( paramDef.containsKey( PARAM_LAST_TAG)) {
                    String strLastTag = paramDef.get( PARAM_LAST_TAG);
                    try {
                        if ( Boolean.parseBoolean( strLastTag)) {
                        	return true;
                        }
                    } catch ( Exception e) {
                        throw new ParseException( cell);
                    }
                }
                
                cell = row.getCell(columnIdx);
                if ( cell != null && cell.getCellType() == CellType.STRING && !strCellValue.equals( cell.getStringCellValue())) {
                    // 解析後、値が変わっていれば、再帰的に解析する
                    if ( parseCell( sheet, data, sheetData, cell, row, columnIdx)) {
                        // 終了タグによる終了
                        return true;
                    }
                }
                break;
            }
        }
        return false;
    }
    
    /**
     * パーサの追加
     * 
     * @param tagParser 対象のTagParser
     */
    public void addTagParser( TagParser<?> tagParser) {
        tagParsers.add( tagParser);
    }

    /**
     * パーサの一覧取得
     * 
     * @return 現在設定されているパーサのリスト
     */
    public List<TagParser<?>> getTagParsers() {
        return tagParsers;
    }
}
