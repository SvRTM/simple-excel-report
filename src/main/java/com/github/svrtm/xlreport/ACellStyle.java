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

import static com.github.svrtm.xlreport.CellStyle.Alignment.CENTER;
import static com.github.svrtm.xlreport.CellStyle.Alignment.VERTICAL_CENTER;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK;

import java.lang.reflect.Constructor;
import java.lang.reflect.ParameterizedType;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.svrtm.xlreport.Font.Boldweight;

/**
 * @author Artem.Smirnov
 */
abstract class ACellStyle<TC extends ACell<?, ?>, TCS extends ACellStyle<TC, TCS, TF>, TF extends Font<TCS>>
        implements com.github.svrtm.xlreport.CellStyle {
    final TC cell;
    final ABuilder<?, ?> builder;

    final Workbook wb;

    final CellStyle_p cellStyle_p;

    ACellStyle(final TC cell) {
        this.cell = cell;
        this.builder = cell.builder;
        this.wb = cell.builder.sheet.getWorkbook();

        cellStyle_p = new CellStyle_p();
    }

    public abstract TC createStyle();

    CellStyle getStyle() {
        CellStyle poiStyle = builder.cacheCellStyle.get(cellStyle_p);
        if (poiStyle == null) {
            poiStyle = wb.createCellStyle();
            cellStyle_p.copyTo(poiStyle);
            final Font_p fontWrapper = cellStyle_p.font_p;
            if (fontWrapper != null) {
                final org.apache.poi.ss.usermodel.Font font;
                font = builder.cacheFont.get(fontWrapper);
                poiStyle.setFont(font);
            }
            builder.cacheCellStyle.put(cellStyle_p, poiStyle);
        }

        return poiStyle;
    }

    @SuppressWarnings("unchecked")
    public TCS alignment(final Alignment alignment) {
        if (Alignment.AlignmentType.V == alignment.getAlignmentType())
            cellStyle_p.setVerticalAlignment(alignment.idx);
        else
            cellStyle_p.setAlignment(alignment.idx);

        return (TCS) this;
    }

    /**
     * Edging of a cell of black color
     *
     * @return
     */
    @SuppressWarnings("unchecked")
    public TCS defaultEdging() {
        cellStyle_p.setBorderBottom(BORDER_THIN);
        cellStyle_p.setBottomBorderColor(BLACK.getIndex());
        cellStyle_p.setBorderLeft(BORDER_THIN);
        cellStyle_p.setLeftBorderColor(BLACK.getIndex());
        cellStyle_p.setBorderRight(BORDER_THIN);
        cellStyle_p.setRightBorderColor(BLACK.getIndex());
        cellStyle_p.setBorderTop(BORDER_THIN);
        cellStyle_p.setTopBorderColor(BLACK.getIndex());

        return (TCS) this;
    }

    @SuppressWarnings("unchecked")
    public TCS wrapText() {
        cellStyle_p.setWrapText(true);
        return (TCS) this;
    }

    @SuppressWarnings("unchecked")
    public TCS fillPattern(final FillPattern fp) {
        cellStyle_p.setFillPattern(fp.idx);
        return (TCS) this;
    }

    @SuppressWarnings("unchecked")
    public TCS fillForegroundColor(final IndexedColors color) {
        cellStyle_p.setFillForegroundColor(color.index);
        return (TCS) this;
    }

    @SuppressWarnings("unchecked")
    public TCS dataFormat(final String format) {
        Short fmt = builder.cacheDataFormat.get(format);
        if (fmt == null) {
            fmt = wb.createDataFormat().getFormat(format);
            builder.cacheDataFormat.put(format, fmt);
        }
        cellStyle_p.setDataFormat(fmt);

        return (TCS) this;
    }

    @SuppressWarnings("unchecked")
    public TF addFont() {
        final ParameterizedType pt = (ParameterizedType) getClass()
                .getGenericSuperclass();
        final ParameterizedType paramType = (ParameterizedType) pt
                .getActualTypeArguments()[2];
        final Class<TF> fontClass = (Class<TF>) paramType.getRawType();
        try {
            final Constructor<TF> c = fontClass
                    .getDeclaredConstructor(ACellStyle.class);
            return c.newInstance(this);
        }
        catch (final Exception e) {
            throw new ReportBuilderException(e);
        }
    }

    /**
     * Prepared style for title
     * Set values of style:
     * center alignment of a horizontal and vertical
     * edging of a cell of black color
     * font size of 11pt, bold weight
     *
     * @return
     */
    @SuppressWarnings("unchecked")
    public TCS title() {
        alignment(CENTER);
        alignment(VERTICAL_CENTER);
        wrapText();
        defaultEdging();
        addFont().heightInPoints((short) 11).boldweight(Boldweight.BOLD)
                .configureFont();

        return (TCS) this;
    }

    /**
     * Prepared style for header
     * Set values of style:
     * center alignment of a horizontal and vertical
     * wrapped text
     * edging of a cell of black color
     * font size of 10pt, bold weight, black color
     *
     * @return
     */
    @SuppressWarnings("unchecked")
    public TCS header() {
        alignment(CENTER);
        alignment(VERTICAL_CENTER);
        wrapText();
        defaultEdging();
        addFont().heightInPoints((short) 10).color(IndexedColors.BLACK)
                .boldweight(Boldweight.BOLD).configureFont();

        return (TCS) this;
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
    public TCS headerWithForegroundColor(final IndexedColors color) {
        fillForegroundColor(color);
        headerWithForegroundColor();
        return (TCS) this;
    }

    void headerWithForegroundColor() {
        header();
        fillPattern(FillPattern.SOLID_FOREGROUND);
    }

    /**
     * Prepared style for 'Total' cell/s
     * Set values of style:
     * horizontal alignment
     * edging of a cell of black color
     * font size of 10pt, bold weight, black color, italic
     *
     * @return
     */
    @SuppressWarnings("unchecked")
    public TCS total() {
        alignment(CENTER);
        defaultEdging();
        addFont().heightInPoints((short) 10).color(IndexedColors.BLACK).italic()
                .boldweight(Boldweight.BOLD).configureFont();

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
    public TCS totalWithForegroundColor(final IndexedColors fg) {
        fillForegroundColor(fg);
        totalWithForegroundColor();
        return (TCS) this;
    }

    void totalWithForegroundColor() {
        total();
        fillPattern(FillPattern.SOLID_FOREGROUND);
    }

}
