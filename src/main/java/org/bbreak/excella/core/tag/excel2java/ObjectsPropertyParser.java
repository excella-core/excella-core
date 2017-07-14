/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ObjectsPropertyParser.java 2 2009-05-08 07:39:20Z yuta-takahashi $
 * $Revision: 2 $
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
