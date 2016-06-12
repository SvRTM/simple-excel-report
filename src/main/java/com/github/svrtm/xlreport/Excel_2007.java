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

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Artem.Smirnov
 */
public final class Excel_2007 {

    final static public Body instanceBody(final Header header) {
        header.init(new XSSFWorkbook(), -1, Row07.class);
        header.prepareHeader();
        return new Body(header);
    }

    final static public Body instanceBody() {
        return new Body(new XSSFWorkbook(), -1);
    }

    public static abstract class Header
            extends AHeader<Header, Row07<Header>> {}

    final static public class Body extends ABody<Body, Row07<Body>> {

        private Body(final Workbook wb, final int maxrow) {
            init(wb, maxrow, Row07.class);
        }

        private Body(final Header header) {
            init(header);
        }
    }
}
