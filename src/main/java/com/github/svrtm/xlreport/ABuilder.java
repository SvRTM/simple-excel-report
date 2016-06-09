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

import static com.github.svrtm.xlreport.Row.RowOperation.CREATE;
import static com.github.svrtm.xlreport.Row.RowOperation.GET;
import static com.github.svrtm.xlreport.Row.RowOperation.GET_CREATE;
import static java.lang.String.format;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.github.svrtm.xlreport.Report.AHeader;

/**
 * @author Artem.Smirnov
 */
public abstract class ABuilder<HB> {
    AHeader header;

    Class<? extends Font> fontClass;
    Class<? extends CellStyle> cellStyleClass;

    Workbook wb;
    Sheet sheet;
    int maxrow = -1;

    List<CellRangeAddress> regionsList;
    Map<Font_p, Font> cacheFont;

    Map<CellStyle_p, CellStyle> cacheCellStyle;

    Map<String, Short> cacheDataFormat;

    //
    //
    protected void init(final Workbook wb, final int maxrow) {
        this.maxrow = maxrow;
        this.wb = wb;
        this.sheet = wb.createSheet();

        this.fontClass = wb.getFontAt((short) 0).getClass();
        this.cellStyleClass = wb.getCellStyleAt((short) 0).getClass();

        this.cacheFont = new HashMap<Font_p, Font>(2);
        this.cacheCellStyle = new HashMap<CellStyle_p, CellStyle>();
        this.cacheDataFormat = new HashMap<String, Short>(2);
    }

    /**
     * used in {@link ABuilder#addRow(int, ABuilder.ICallback)}
     *
     * @author Artem.Smirnov
     */
    public interface ICallback {
        /**
         * @param row
         *            - new row
         */
        void step(final Row<?> row);
    }

    /**
     * The index of a new row is the number of the last row on the sheet
     * plus 1
     *
     * @return
     */
    public Row<HB> addRow() {
        return new Row<HB>(this);
    }

    /**
     * Create a new row within the sheet and return the high level
     * representation
     *
     * @param rowIndex
     *            - logical row
     * @return
     */
    public Row<HB> addRow(final int rowIndex) {
        return new Row<HB>(this, rowIndex, CREATE);
    }

    /**
     * Create a new row
     *
     * @param nRow
     *            - number of new created rows
     * @param callback
     * @return
     */
    @SuppressWarnings("unchecked")
    public HB addRow(final int nRow, final ICallback callback) {
        for (int i = 0; i < nRow; i++)
            callback.step(addRow());
        return (HB) this;
    }

    /**
     * Returns the row 0-based
     *
     * @param i
     *            - row to get
     * @return
     */
    public Row<HB> row(final int i) {
        return new Row<HB>(this, i, GET);
    }

    Row<HB> rowOrCreateIfAbsent(final int i) {
        return new Row<HB>(this, i, GET_CREATE);
    }

    /**
     * Returns the number of physically defined rows (NOT the number of rows
     * in the sheet)
     *
     * @return the number of physically defined rows in this sheet
     */
    public int getIndexOfLastRow() {
        return sheet.getPhysicalNumberOfRows();
    }

    /**
     * Adds a merged region of cells (hence those cells form one)
     *
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     * @return
     */
    @SuppressWarnings("unchecked")
    public HB mergedRegion(final int firstRow, final int lastRow,
                           final int firstCol, final int lastCol) {
        if (regionsList == null)
            regionsList = new ArrayList<CellRangeAddress>();

        CellRangeAddress region;
        region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        regionsList.add(region);
        sheet.addMergedRegion(region);
        return (HB) this;
    }

    /**
     * Set the width
     *
     * @param columnIndex
     *            - the column to set (0-based)
     * @param width
     *            - the width
     * @return
     */
    @SuppressWarnings("unchecked")
    public HB whereiColumnWidthIs(final int columnIndex, final int width) {
        sheet.setColumnWidth(columnIndex, width * 256);
        return (HB) this;
    }

    /**
     * Set the width
     *
     * @param widthColumns
     *            - [columnIndex, width]
     * @return
     */
    @SuppressWarnings("unchecked")
    public HB whereiColumnsWidthIs(final int widthColumns[][]) {
        for (final int[] is : widthColumns)
            sheet.setColumnWidth(is[0], is[1] * 256);

        return (HB) this;
    }

    /**
     * Only for {@link HSSFWorkbook}
     *
     * @param i
     *            - the palette index, between 0x8 to 0x40 inclusive
     * @param red
     *            - the RGB red component, between 0 and 255 inclusive
     * @param green
     *            - the RGB green component, between 0 and 255 inclusive
     * @param blue
     *            - the RGB blue component, between 0 and 255 inclusive
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB replaceColor(final short i, final byte red, final byte green,
                           final byte blue) {
        if (!(wb instanceof HSSFWorkbook))
            throw new ReportBuilderException(
                    format("This operation is not supported for workbook %1",
                            wb.getClass().getSimpleName()));

        final HSSFWorkbook hssfwb = (HSSFWorkbook) wb;
        final HSSFPalette palette = hssfwb.getCustomPalette();
        palette.setColorAtIndex(i, red, green, blue);

        return (HB) this;
    }

    @SuppressWarnings("unchecked")
    public HB withAutoSizeColumn(final int column) {
        sheet.autoSizeColumn(column);
        return (HB) this;
    }
}
