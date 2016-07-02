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

/**
 * <code>Font07</code> used for <code>XSSF (Excel_2007)</code> and
 * <code>SXSSF (Excel_2007XL)</code>
 *
 * @author Artem.Smirnov
 */
public class Font07<TCS extends ACellStyle<?, ?, ?>> extends Font<TCS> {

    Font07(final TCS cellStyle) {
        super(cellStyle);
    }

    /**
     * Set the color for the font in Standard Alpha Red Green Blue (ARGB or RGB)
     * color value
     *
     * @param rgb
     * @return this
     */
    public Font<TCS> color(final byte[] rgb) {
        font_p.setColor(rgb);
        return this;
    }

    /**
     * Set the color for the font
     *
     * @param red
     * @param green
     * @param blue
     * @return this
     */
    public Font<TCS> color(final byte red, final byte green, final byte blue) {
        font_p.setColor(new byte[] { red, green, blue });
        return this;
    }
}
