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

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author Artem.Smirnov
 */
public class Font<TCS extends ACellStyle<?, ?, ?>> {
    final private TCS cellStyle;
    final private ABuilder<?, ?> builder;

    final Font_p font_p;

    public enum Boldweight {
        /** */
        NORMAL(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_NORMAL),
        /** */
        BOLD(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD);

        short idx;

        private Boldweight(final short idx) {
            this.idx = idx;
        }
    }

    Font(final TCS cellStyle) {
        this.cellStyle = cellStyle;
        this.builder = cellStyle.builder;
        font_p = new Font_p();
    }

    public TCS configureFont() {
        getFont();
        cellStyle.cellStyle_p.font_p = font_p;

        return cellStyle;
    }

    public Font<TCS> heightInPoints(final short height) {
        font_p.setFontHeightInPoints(height);
        return this;
    }

    public Font<TCS> color(final IndexedColors color) {
        font_p.setColor(color.index);
        return this;
    }

    public Font<TCS> italic() {
        font_p.setItalic(true);
        return this;
    }

    public Font<TCS> boldweight(final Boldweight boldweight) {
        font_p.setBoldweight(boldweight.idx);
        return this;
    }

    public Font<TCS> name(final String name) {
        font_p.setName(name);
        return this;
    }

    private org.apache.poi.ss.usermodel.Font getFont() {
        org.apache.poi.ss.usermodel.Font poiFont;
        poiFont = builder.cacheFont.get(font_p);
        if (poiFont == null) {
            poiFont = cellStyle.wb.createFont();
            font_p.copyTo(poiFont);
            builder.cacheFont.put(font_p, poiFont);
        }

        return poiFont;
    }
}
