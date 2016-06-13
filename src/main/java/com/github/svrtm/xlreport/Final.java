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

/**
 * @author Artem.Smirnov
 */
final public class Final {
    private final Workbook wb;

    Final(final ABody<?, ?> body) {
        wb = body.wb;
    }

    public Workbook getWorkbook() {
        return wb;
    }
}

//
//
interface IHeader {
    void prepareHeader();
}

abstract class AHeader<H, TR extends Row<H, TR, ?, ?>> extends ABuilder<H, TR>
        implements IHeader {}

abstract class ABody<HB, TR extends Row<HB, TR, ?, ?>>
        extends ABuilder<HB, TR> {

    void init(final AHeader<?, ?> header) {
        this.header = header;
        maxrow = header.maxrow;

        wb = header.wb;
        sheet = wb.getSheetAt(wb.getActiveSheetIndex());

        cacheFont = header.cacheFont;
        cacheCellStyle = header.cacheCellStyle;
        cacheDataFormat = header.cacheDataFormat;

        rowClass = header.rowClass;
    }

    public Final instanceFinal() {
        return new Final(this);
    }
}
