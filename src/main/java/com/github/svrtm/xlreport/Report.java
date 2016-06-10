/**
 * <pre>
 * Copyright © 2012,2016 Artem Smirnov
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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Artem.Smirnov
 */
public enum Report {

    EXCEL97 {
        @Override
        final public Body instanceBody(final Header header) {
            header.init(new HSSFWorkbook(),
                    SpreadsheetVersion.EXCEL97.getLastRowIndex());
            header.prepareHeader();
            return new Body(header);
        }

        @Override
        final public Body instanceBody() {
            return new Body(new HSSFWorkbook(),
                    SpreadsheetVersion.EXCEL97.getLastRowIndex());
        }
    },

    EXCEL2007 {
        /**
         * Specifies how many rows can be accessed at most via getRow().
         * When a new node is created via createRow() and the total number
         * of unflushed records would exceed the specified value, then the
         * row with the lowest index value is flushed and cannot be accessed
         * via getRow() anymore.
         */
        private static final int ROW_ACCESS_WINDOW_SIZE = 10000;

        @Override
        final public Body instanceBody(final Header header) {
            header.init(new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE),
                    SpreadsheetVersion.EXCEL2007.getLastRowIndex());
            header.prepareHeader();
            return new Body(header);
        }

        @Override
        final public Body instanceBody() {
            return new Body(new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE),
                    SpreadsheetVersion.EXCEL2007.getLastRowIndex());
        }
    },

    EXCEL2007_NEW {
        @Override
        final public Body instanceBody(final Header header) {
            header.init(new XSSFWorkbook(), -1);
            header.prepareHeader();
            return new Body(header);
        }

        @Override
        final public Body instanceBody() {
            return new Body(new XSSFWorkbook(), -1);
        }
    };

    public abstract Body instanceBody(final Header header);

    public abstract Body instanceBody();

    //
    //
    final public class Final {
        private final Workbook wb;

        private Final(final Body body) {
            wb = body.wb;
        }

        public Workbook getWorkbook() {
            return wb;
        }
    }

    //
    //
    private interface IHeader {
        void prepareHeader();
    }

    public static abstract class Header extends ABuilder<Header>
            implements IHeader {}

    //
    //
    final public class Body extends ABuilder<Body> {

        private Body(final Workbook wb, final int maxrow) {
            init(wb, maxrow);
        }

        private Body(final Header header) {
            init(header);
        }

        private void init(final Header header) {
            this.header = header;
            maxrow = header.maxrow;

            wb = header.wb;
            sheet = wb.getSheetAt(wb.getActiveSheetIndex());

            cacheFont = header.cacheFont;
            cacheCellStyle = header.cacheCellStyle;
            cacheDataFormat = header.cacheDataFormat;
        }

        public Final instanceFinal() {
            return new Final(this);
        }
    }

}
