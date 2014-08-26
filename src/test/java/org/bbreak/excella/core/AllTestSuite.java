/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Core - ExcelファイルをJavaから利用するための共通基盤
 *
 * $Id: AllTestSuite.java 29 2009-05-14 09:23:50Z yuta-takahashi $
 * $Revision: 29 $
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
package org.bbreak.excella.core;

import org.bbreak.excella.core.exception.ExceptionTestSuite;
import org.bbreak.excella.core.exporter.book.ExporterBookTestSuite;
import org.bbreak.excella.core.exporter.sheet.ExporterSheetTestSuite;
import org.bbreak.excella.core.handler.HandlerTestSuite;
import org.bbreak.excella.core.listener.ListenerTestSuite;
import org.bbreak.excella.core.tag.TagTestSuite;
import org.bbreak.excella.core.tag.excel2java.TagExcel2JavaTestSuite;
import org.bbreak.excella.core.util.UtilTestSuite;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;

/**
 * 全テスト実行用テストスイート
 * 
 * @since 1.0
 */
@RunWith(Suite.class)
@SuiteClasses({
	TestSuite.class,
	ExceptionTestSuite.class,
	ExporterBookTestSuite.class,
	ExporterSheetTestSuite.class,
	HandlerTestSuite.class,
	ListenerTestSuite.class,
	TagTestSuite.class,
    TagExcel2JavaTestSuite.class,
	UtilTestSuite.class
})
public class AllTestSuite {
    
}
