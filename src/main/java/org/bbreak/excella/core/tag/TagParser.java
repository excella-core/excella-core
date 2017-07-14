/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: TagParser.java 56 2009-05-21 04:13:39Z yuta-takahashi $
 * $Revision: 56 $
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
        if ( tagCell.getCellTypeEnum() == CellType.STRING) {
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
