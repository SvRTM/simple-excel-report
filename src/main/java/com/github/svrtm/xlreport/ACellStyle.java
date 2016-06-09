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

import static com.github.svrtm.xlreport.ACellStyle.Alignment.CENTER;
import static com.github.svrtm.xlreport.ACellStyle.Alignment.VERTICAL_CENTER;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN;
import static org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND;
import static org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK;

import org.apache.poi.hssf.util.HSSFColor.AQUA;
import org.apache.poi.hssf.util.HSSFColor.BLUE;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.cglib.beans.BeanCopier;

/**
 * @author Artem.Smirnov
 */
public abstract class ACellStyle<T extends ACellStyle<T, ? extends ACell<?, ?>>, HB extends ACell<HB, ?>> {
    final protected HB cell;
    final protected ABuilder<?> builder;

    final protected Workbook wb;
    final protected Row poiRow;

    final CellStyle_p cellStyle_p;

    public enum Alignment {
        /**
         *
         */
        CENTER(CellStyle.ALIGN_CENTER, AlignmentType.G),
        /**
         *
         */
        RIGHT(CellStyle.ALIGN_RIGHT, AlignmentType.G),
        /**
         *
         */
        LEFT(ALIGN_LEFT, AlignmentType.G),
        /**
         *
         */
        VERTICAL_CENTER(CellStyle.VERTICAL_CENTER, AlignmentType.V);

        private enum AlignmentType {
            G, V
        };

        short code;
        AlignmentType alignmentType;

        private Alignment(final short code, final AlignmentType alignmentType) {
            this.code = code;
            this.alignmentType = alignmentType;
        }

        public short getCode() {
            return code;
        }

        public AlignmentType getAlignmentType() {
            return alignmentType;
        }
    }

    ACellStyle(final HB cell) {
        this.cell = cell;
        this.builder = cell.builder;
        this.wb = cell.builder.sheet.getWorkbook();

        this.poiRow = cell.row.poiRow;
        cellStyle_p = new CellStyle_p();
    }

    public abstract HB buildStyle();

    CellStyle getStyle() {
        CellStyle poiStyle = builder.cacheCellStyle.get(cellStyle_p);
        if (poiStyle == null) {
            poiStyle = wb.createCellStyle();
            final BeanCopier beanCopier = BeanCopier.create(CellStyle_p.class,
                    builder.cellStyleClass, false);
            beanCopier.copy(cellStyle_p, poiStyle, null);
            final Font_p fontWrapper = cellStyle_p.font_p;
            if (fontWrapper != null) {
                final org.apache.poi.ss.usermodel.Font font = builder.cacheFont
                        .get(fontWrapper);
                poiStyle.setFont(font);
            }
            builder.cacheCellStyle.put(cellStyle_p, poiStyle);
        }

        return poiStyle;
    }

    @SuppressWarnings("unchecked")
    public T alignment(final Alignment alignment) {
        if (Alignment.AlignmentType.V == alignment.getAlignmentType())
            cellStyle_p.setVerticalAlignment(alignment.getCode());
        else
            cellStyle_p.setAlignment(alignment.getCode());

        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T defaultBorder() {
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
    public T pattern(final short i) {
        cellStyle_p.setFillPattern(i);
        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T foregroundColor(final short i) {
        cellStyle_p.setFillForegroundColor(i);
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

    @SuppressWarnings("unchecked")
    public T title() {
        alignment(CENTER);
        alignment(VERTICAL_CENTER);
        defaultBorder();
        addFont().heightInPoints((short) 11).boldweight(BOLDWEIGHT_BOLD)
                .buildFont();
        wrapText();

        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T header() {
        alignment(CENTER);
        alignment(VERTICAL_CENTER);
        wrapText();
        defaultBorder();
        addFont().heightInPoints((short) 10)
                .color(IndexedColors.BLACK.getIndex())
                .boldweight(BOLDWEIGHT_BOLD).buildFont();

        return (T) this;
    }

    /**
     * Header style with the blue color
     *
     * @return this
     */
    @SuppressWarnings("unchecked")
    public T headerWithColor() {
        header().foregroundColor(BLUE.index).pattern(SOLID_FOREGROUND);
        return (T) this;
    }

    @SuppressWarnings("unchecked")
    public T total() {
        alignment(CENTER);
        defaultBorder();
        addFont().heightInPoints((short) 10)
                .color(IndexedColors.BLACK.getIndex()).italic()
                .boldweight(BOLDWEIGHT_BOLD).buildFont();

        return (T) this;
    }

    /**
     * Total style with the aqua color
     *
     * @return this
     */
    @SuppressWarnings("unchecked")
    public T totalWithColor() {
        total().foregroundColor(AQUA.index).pattern(SOLID_FOREGROUND);
        return (T) this;
    }

}
