/**
 * <pre>
 * Copyright © 2012 Artem Smirnov
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

import net.sf.cglib.beans.BeanCopier;

/**
 * @author Artem.Smirnov
 */
final public class Font<T extends ACellStyle<?, ?>> {
    final private T cellStyle;
    final private ABuilder<?> builder;

    final private Font_p font_p;

    Font(final T cellStyle) {
        this.cellStyle = cellStyle;
        this.builder = cellStyle.builder;
        font_p = new Font_p();
    }

    public T buildFont() {
        getFont();
        cellStyle.cellStyle_p.font_p = font_p;

        return cellStyle;
    }

    org.apache.poi.ss.usermodel.Font getFont() {
        org.apache.poi.ss.usermodel.Font poiFont =
                                                 builder.cacheFont.get(font_p);
        if (poiFont == null) {
            poiFont = cellStyle.wb.createFont();
            final BeanCopier beanCopier = BeanCopier.create(Font_p.class,
                    builder.fontClass, false);
            beanCopier.copy(font_p, poiFont, null);
            builder.cacheFont.put(font_p, poiFont);
        }

        return poiFont;
    }

    public Font<T> heightInPoints(final short height) {
        font_p.setFontHeightInPoints(height);
        return this;
    }

    public Font<T> color(final short i) {
        font_p.setColor(i);
        return this;
    }

    public Font<T> italic() {
        font_p.setItalic(true);
        return this;
    }

    public Font<T> boldweight(final short boldweight) {
        font_p.setBoldweight(boldweight);
        return this;
    }
}
