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
        final public Header createHeader() {
            return new Header(new HSSFWorkbook(),
                    SpreadsheetVersion.EXCEL97.getLastRowIndex());
        }

        @Override
        final public Body createBody() {
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
        final public Header createHeader() {
            return new Header(new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE),
                    SpreadsheetVersion.EXCEL2007.getLastRowIndex());
        }

        @Override
        final public Body createBody() {
            return new Body(new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE),
                    SpreadsheetVersion.EXCEL2007.getLastRowIndex());
        }
    },

    EXCEL2007_NEW {
        @Override
        final public Header createHeader() {
            return new Header(new XSSFWorkbook(), -1);
        }

        @Override
        final public Body createBody() {
            return new Body(new XSSFWorkbook(), -1);
        }
    };

    public abstract Header createHeader();

    public abstract Body createBody();

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
        void builder();
    }

    public static abstract class AHeader extends ABuilder<AHeader>
            implements IHeader {}

    final public class Header {
        private final Workbook wb;
        private final int maxrow;

        private Header(final Workbook wb, final int maxrow) {
            this.wb = wb;
            this.maxrow = maxrow;
        }

        public Body build(final AHeader header) {
            header.init(wb, maxrow);

            header.builder();
            return new Body(header);
        }
    }

    //
    //
    final public class Body extends ABuilder<Body> {

        public Body(final Workbook wb, final int maxrow) {
            init(wb, maxrow);
        }

        private Body(final AHeader header) {
            init(header);
        }

        private void init(final AHeader header) {
            this.header = header;
            maxrow = header.maxrow;

            wb = header.wb;
            sheet = wb.getSheetAt(wb.getActiveSheetIndex());

            cacheFont = header.cacheFont;
            cacheCellStyle = header.cacheCellStyle;
            cacheDataFormat = header.cacheDataFormat;
        }

        public Final build() {
            return new Final(this);
        }
    }

}
