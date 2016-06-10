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

import static com.github.svrtm.xlreport.ACellStyle.Alignment.CENTER;
import static com.github.svrtm.xlreport.ACellStyle.Alignment.VERTICAL_CENTER;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.svrtm.xlreport.Font.Boldweight;

/**
 * @author Artem.Smirnov
 */
public abstract class ACellStyle<T extends ACellStyle<T, ? extends ACell<HB, ?>>, HB extends ACell<HB, ?>> {
    final HB cell;
    final ABuilder<?> builder;

    final Workbook wb;

    final CellStyle_p cellStyle_p;

    public enum Alignment {
        /** */
        CENTER(CellStyle.ALIGN_CENTER, AlignmentType.H),
        /** */
        RIGHT(CellStyle.ALIGN_RIGHT, AlignmentType.H),
        /** */
        LEFT(ALIGN_LEFT, AlignmentType.H),
        /** */
        VERTICAL_CENTER(CellStyle.VERTICAL_CENTER, AlignmentType.V);

        private enum AlignmentType {
            H, V
        };

        short idx;
        private AlignmentType alignmentType;

        private Alignment(final short idx, final AlignmentType alignmentType) {
            this.idx = idx;
            this.alignmentType = alignmentType;
        }

        public AlignmentType getAlignmentType() {
            return alignmentType;
        }
    }

    public enum FillPattern {
        /** */
        NO_FILL(CellStyle.NO_FILL),
        /** */
        SOLID_FOREGROUND(CellStyle.SOLID_FOREGROUND),
        /** */
        FINE_DOTS(CellStyle.FINE_DOTS),
        /** */
        ALT_BARS(CellStyle.ALT_BARS),
        /** */
        SPARSE_DOTS(CellStyle.SPARSE_DOTS),
        /** */
        THICK_HORZ_BANDS(CellStyle.THICK_HORZ_BANDS),
        /** */
        THICK_VERT_BANDS(CellStyle.THICK_VERT_BANDS),
        /** */
        THICK_BACKWARD_DIAG(CellStyle.THICK_BACKWARD_DIAG),
        /** */
        THICK_FORWARD_DIAG(CellStyle.THICK_FORWARD_DIAG),
        /** */
        BIG_SPOTS(CellStyle.BIG_SPOTS),
        /** */
        BRICKS(CellStyle.BRICKS),
        /** */
        THIN_HORZ_BANDS(CellStyle.THIN_HORZ_BANDS),
        /** */
        THIN_VERT_BANDS(CellStyle.THIN_VERT_BANDS),
        /** */
        THIN_BACKWARD_DIAG(CellStyle.THIN_BACKWARD_DIAG),
        /** */
        THIN_FORWARD_DIAG(CellStyle.THIN_FORWARD_DIAG),
        /** */
        SQUARES(CellStyle.SQUARES),
        /** */
        DIAMONDS(CellStyle.DIAMONDS);

        short idx;

        private FillPattern(final short idx) {
            this.idx = idx;
        }
    }

    ACellStyle(final HB cell) {
        this.cell = cell;
        this.builder = cell.builder;
        this.wb = cell.builder.sheet.getWorkbook();

        cellStyle_p = new CellStyle_p();
    }

    public abstract HB createStyle();

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
    public T alignment(final Alignment alignment) {
        if (Alignment.AlignmentType.V == alignment.getAlignmentType())
            cellStyle_p.setVerticalAlignment(alignment.idx);
        else
            cellStyle_p.setAlignment(alignment.idx);

        return (T) this;
    }

    /**
     * Edging of a cell of black color
     *
     * @return
     */
    @SuppressWarnings("unchecked")
    public T defaultEdging() {
        cellStyle_p.setBorderBottom(BORDER_THIN);
        cellStyle_p.setBottomBorderColor(BLACK.getIndex());
        cellStyle_p.setBorderLeft(BORDER_THIN);
        cellStyle_p.setLeftBorderColor(BLACK.getIndex());
        cellStyle_p.setBorderRight(BORDER_THIN);
        cellStyle_p.setRightBorderColor(BLACK.getIndex());
        cellStyle_p.setBorderTop(BORDER_THIN);
        cellStyle_p.setTopBorderColor(BLACK.getIndex());

        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T wrapText() {
        cellStyle_p.setWrapText(true);
        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T fillPattern(final FillPattern fp) {
        cellStyle_p.setFillPattern(fp.idx);
        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T fillForegroundColor(final IndexedColors color) {
        cellStyle_p.setFillForegroundColor(color.index);
        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T dataFormat(final String format) {
        Short fmt = builder.cacheDataFormat.get(format);
        if (fmt == null) {
            fmt = wb.createDataFormat().getFormat(format);
            builder.cacheDataFormat.put(format, fmt);
        }
        cellStyle_p.setDataFormat(fmt);

        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public Font<T> addFont() {
        return new Font<T>((T) this);
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
    public T title() {
        alignment(CENTER);
        alignment(VERTICAL_CENTER);
        wrapText();
        defaultEdging();
        addFont().heightInPoints((short) 11).boldweight(Boldweight.BOLD)
                .configureFont();

        return (T) this;
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
    public T header() {
        alignment(CENTER);
        alignment(VERTICAL_CENTER);
        wrapText();
        defaultEdging();
        addFont().heightInPoints((short) 10).color(IndexedColors.BLACK)
                .boldweight(Boldweight.BOLD).configureFont();

        return (T) this;
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
    public T headerWithForegroundColor(final IndexedColors color) {
        header();
        fillForegroundColor(color).fillPattern(FillPattern.SOLID_FOREGROUND);
        return (T) this;
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
    public T total() {
        alignment(CENTER);
        defaultEdging();
        addFont().heightInPoints((short) 10).color(IndexedColors.BLACK).italic()
                .boldweight(Boldweight.BOLD).configureFont();

        return (T) this;
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
    public T totalWithForegroundColor(final IndexedColors fg) {
        total();
        fillForegroundColor(fg).fillPattern(FillPattern.SOLID_FOREGROUND);
        return (T) this;
    }

}
