/**
 * <pre>
 * Copyright © 2016 Artem Smirnov
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * </pre>
 */
package com.github.svrtm.xlreport;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * @author Artem.Smirnov
 */
public final class Excel_2007XL {
    private static final int ROW_ACCESS_WINDOW_SIZE = 10000;

    final static public Body instanceBody(final Header header) {
        header.init(new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE),
                SpreadsheetVersion.EXCEL2007.getLastRowIndex(), Row07XL.class);
        header.prepareHeader();
        return new Body(header);
    }

    final static public Body instanceBody() {
        return new Body(new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE),
                SpreadsheetVersion.EXCEL2007.getLastRowIndex());
    }

    public static abstract class Header
            extends AHeader<Header, Row07XL<Header>> {

        @Override
        public Header withAutoSizeColumn(final int column) {
            ((SXSSFSheet) sheet).trackColumnForAutoSizing(column);
            return super.withAutoSizeColumn(column);
        }
    }

    final static public class Body extends ABody<Body, Row07XL<Body>> {

        private Body(final Workbook wb, final int maxrow) {
            init(wb, maxrow, Row07XL.class);
        }

        private Body(final Header header) {
            init(header);
        }

        @Override
        public Body withAutoSizeColumn(final int column) {
            ((SXSSFSheet) sheet).trackColumnForAutoSizing(column);
            return super.withAutoSizeColumn(column);
        }
    }
}
