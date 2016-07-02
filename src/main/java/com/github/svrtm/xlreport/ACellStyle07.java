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

/**
 * @author Artem.Smirnov
 */
abstract class ACellStyle07<TC extends ACell<?, ?>, TCS extends ACellStyle07<TC, TCS, TF>, TF extends Font<TCS>>
        extends ACellStyle<TC, TCS, TF> {

    ACellStyle07(final TC cell) {
        super(cell);
    }

    /**
     * Prepared style with fill foreground color for header
     * Set values of style:
     * center alignment of a horizontal and vertical
     * wrapped text
     * edging of a cell of black color
     * font size of 10pt, bold weight, black color
     *
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TCS headerWithForegroundColor(final byte red, final byte green,
                                         final byte blue) {
        fillForegroundColor(new byte[] { red, green, blue });
        headerWithForegroundColor();
        return (TCS) this;
    }

    /**
     * Specifying the fill background color for cell styles
     *
     * @param fgRgb
     *            Sets the Red Green Blue or Alpha Red Green Blue
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TCS fillForegroundColor(final byte[] fgRgb) {
        cellStyle_p.setFillForegroundColor(fgRgb);
        return (TCS) this;
    }

    /**
     * Specifying the fill background color for cell styles
     *
     * @param red
     * @param green
     * @param blue
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TCS fillForegroundColor(final byte red, final byte green,
                                   final byte blue) {
        cellStyle_p.setFillForegroundColor(new byte[] { red, green, blue });
        return (TCS) this;
    }

    /**
     * Prepared style with fill foreground color for 'Total' cell/s
     * Set values of style:
     * horizontal alignment
     * edging of a cell of black color
     * font size of 10pt, bold weight, black color, italic
     *
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TCS totalWithForegroundColor(final byte red, final byte green,
                                        final byte blue) {
        fillForegroundColor(new byte[] { red, green, blue });
        totalWithForegroundColor();
        return (TCS) this;
    }

}
