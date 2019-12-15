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
