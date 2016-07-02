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
        /** Normal boldness (not bold) */
        NORMAL(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_NORMAL),
        /** Bold boldness (bold) */
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

    /**
     * Creates a new prepared font
     *
     * @return implementation of <code>CellStyle</code>
     */

    /**
     * Finalization of the implementation <code>Font</code>.
     *
     * @return an instance of the implementation <code>ACellStyle</code>
     * @see com.github.svrtm.xlreport.ACellStyle#addFont()
     */
    public TCS configureFont() {
        generatePoiFont();
        cellStyle.cellStyle_p.font_p = font_p;

        return cellStyle;
    }

    /**
     * Set the font height in points.
     *
     * @param height
     *            - height in the familiar unit of measure - points
     * @return this
     */
    public Font<TCS> heightInPoints(final short height) {
        font_p.setFontHeightInPoints(height);
        return this;
    }

    /**
     * Set the indexed color for the font
     *
     * @param color
     *            - indexed color to use
     * @return this
     * @see org.apache.poi.ss.usermodel.IndexedColors
     */
    public Font<TCS> color(final IndexedColors color) {
        font_p.setColor(color.index);
        return this;
    }

    /**
     * Set the italic to use
     *
     * @return this
     */
    public Font<TCS> italic() {
        font_p.setItalic(true);
        return this;
    }

    /**
     * Set the boldness to use
     *
     * @param boldweight
     * @return this
     */
    public Font<TCS> boldweight(final Boldweight boldweight) {
        font_p.setBoldweight(boldweight.idx);
        return this;
    }

    /**
     * Set the name for the font (i.e. Arial).
     *
     * @param name
     *            String representing the name of the font to use
     * @return
     */
    public Font<TCS> name(final String name) {
        font_p.setName(name);
        return this;
    }

    /**
     * Creates a new font if the prepared font with the given attributes is
     * missing in the cache and saves in cache or returns the font from cache.
     *
     * @return Apache POI excel font
     */
    private org.apache.poi.ss.usermodel.Font generatePoiFont() {
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
