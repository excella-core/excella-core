/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: SheetParseListener.java 2 2009-05-08 07:39:20Z yuta-takahashi $
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
package org.bbreak.excella.core.listener;

import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ParseException;

/**
 * シート解析時のイベント通知リスナ
 * @since 1.0
 */
public interface SheetParseListener {

    /**
     * シート解析前に呼び出されるメソッド
     * 
     * @param sheet 対象シート
     * @param sheetParser 対象パーサ
     */
    void preParse( Sheet sheet, SheetParser sheetParser) throws ParseException;

    /**
     * シート解析後に呼び出されるメソッド
     * 
     * @param sheet 対象シート
     * @param sheetParser 対象パーサ
     * @param sheetData 解析結果
     */
    void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException;
}
