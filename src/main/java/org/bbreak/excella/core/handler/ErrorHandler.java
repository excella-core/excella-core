/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: ErrorHandler.java 119 2009-06-24 08:49:08Z yuta-takahashi $
 * $Revision: 119 $
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
package org.bbreak.excella.core.handler;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * エラーハンドリング用インターフェイス
 * @since 1.0
 *
 * @param <EX> このエラーハンドリングクラスで扱う例外の型
 */
public interface ErrorHandler<EX> {

    /**
     * 例外の通知
     * 
     * @param workbook 処理中のワークブック
     * @param sheet 処理中のシート
     * @param throwable 発生した例外
     */
    void notifyException( Workbook workbook, Sheet sheet, EX throwable);
}
