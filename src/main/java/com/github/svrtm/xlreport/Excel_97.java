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

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Artem.Smirnov
 */
public final class Excel_97 {

    final static public Body instanceBody(final Header header) {
        header.init(new HSSFWorkbook(),
                SpreadsheetVersion.EXCEL97.getLastRowIndex(), Row97.class);
        header.prepareHeader();
        return new Body(header);
    }

    final static public Body instanceBody() {
        return new Body(new HSSFWorkbook(),
                SpreadsheetVersion.EXCEL97.getLastRowIndex());
    }

    public static abstract class Header extends AHeader<Header, Row97<Header>> {

        /**
         * Only for {@link HSSFWorkbook}
         *
         * @param i
         *            - the palette index, between 0x8 to 0x40 inclusive
         * @param red
         *            - the RGB red component, between 0 and 255 inclusive
         * @param green
         *            - the RGB green component, between 0 and 255 inclusive
         * @param blue
         *            - the RGB blue component, between 0 and 255 inclusive
         * @return this
         */
        public Header replaceColor(final IndexedColors color, final byte red,
                                   final byte green, final byte blue) {
            final HSSFWorkbook hssfwb = (HSSFWorkbook) wb;
            final HSSFPalette palette = hssfwb.getCustomPalette();
            palette.setColorAtIndex(color.index, red, green, blue);

            return this;
        }
    }

    final static public class Body extends ABody<Body, Row97<Body>> {

        private Body(final Workbook wb, final int maxrow) {
            init(wb, maxrow, Row97.class);
        }

        private Body(final Header header) {
            init(header);
        }

        /**
         * Only for {@link HSSFWorkbook}
         *
         * @param i
         *            - the palette index, between 0x8 to 0x40 inclusive
         * @param red
         *            - the RGB red component, between 0 and 255 inclusive
         * @param green
         *            - the RGB green component, between 0 and 255 inclusive
         * @param blue
         *            - the RGB blue component, between 0 and 255 inclusive
         * @return this
         */
        public Body replaceColor(final IndexedColors color, final byte red,
                                 final byte green, final byte blue) {
            final HSSFWorkbook hssfwb = (HSSFWorkbook) wb;
            final HSSFPalette palette = hssfwb.getCustomPalette();
            palette.setColorAtIndex(color.index, red, green, blue);

            return this;
        }
    }
}
