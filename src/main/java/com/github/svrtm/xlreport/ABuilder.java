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

import static com.github.svrtm.xlreport.Row.RowOperation.CREATE;
import static com.github.svrtm.xlreport.Row.RowOperation.CREATE_and_GET;
import static com.github.svrtm.xlreport.Row.RowOperation.GET;

import java.lang.reflect.Constructor;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.github.svrtm.xlreport.Row.RowOperation;

/**
 * @author Artem.Smirnov
 */
abstract class ABuilder<HB, TR extends Row<HB, TR, ?, ?>> {
    AHeader<?, ?> header;
    Class<?> rowClass;

    Workbook wb;
    Sheet sheet;
    int maxrow = -1;

    List<CellRangeAddress> regionsList;
    Map<Font_p, Font> cacheFont;

    Map<CellStyle_p, CellStyle> cacheCellStyle;

    Map<String, Short> cacheDataFormat;

    //
    //
    void init(final Workbook wb, final int maxrow, final Class<?> rowClass) {
        this.maxrow = maxrow;
        this.wb = wb;
        this.sheet = wb.createSheet();

        this.cacheFont = new HashMap<Font_p, Font>(2);
        this.cacheCellStyle = new HashMap<CellStyle_p, CellStyle>();
        this.cacheDataFormat = new HashMap<String, Short>(2);

        this.rowClass = rowClass;
    }

    /**
     * used in {@link ABuilder#addNewnRows(int, ABuilder.INewRow)}
     *
     * @author Artem.Smirnov
     */
    public interface INewRow<TR> {
        /**
         * @param row
         *            - new row
         */
        void iRow(final TR row);
    }

    /**
     * Create a new row.
     * The index of a new row is the number of the last row on the sheet
     * plus 1.
     *
     * @return a new instance of the class of implementation <code>Row</code>.
     * @see com.github.svrtm.xlreport.Row
     * @see com.github.svrtm.xlreport.Row97
     * @see com.github.svrtm.xlreport.Row07
     * @see com.github.svrtm.xlreport.Row07XL
     */
    public TR addNewRow() {
        return newRowInstance();
    }

    /**
     * Create a new rows with empty cells.
     *
     * @param nRows
     *            - number of new created rows
     * @param nCells
     *            - number of new empty cells
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB addNewRowsWithEmptyCells(final int nRows, final int nCells) {
        for (int iRow = 0; iRow < nRows; iRow++) {
            final TR r = addNewRow();
            for (int iCell = 0; iCell < nCells; iCell++)
                r.prepareNewCell(iCell).createCell();
            r.configureRow();
        }
        return (HB) this;
    }

    /**
     * Create a new row within the sheet and return the high level
     * representation.
     *
     * @param rowIndex
     *            - logical row
     * @return a new instance of the class of implementation <code>Row</code>.
     * @see com.github.svrtm.xlreport.Row
     * @see com.github.svrtm.xlreport.Row97
     * @see com.github.svrtm.xlreport.Row07
     * @see com.github.svrtm.xlreport.Row07XL
     */
    public TR addRow(final int rowIndex) {
        return newRowInstance(rowIndex, CREATE);
    }

    /**
     * Create a new row
     *
     * @param nRow
     *            - number of new created rows
     * @param callback
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB addNewnRows(final int nRow, final INewRow<TR> callback) {
        for (int i = 0; i < nRow; i++)
            callback.iRow(addNewRow());
        return (HB) this;
    }

    /**
     * Returns the row 0-based.
     *
     * @param i
     *            - row to get.
     * @return an instance of the class of implementation <code>Row</code>.
     * @see com.github.svrtm.xlreport.Row
     * @see com.github.svrtm.xlreport.Row97
     * @see com.github.svrtm.xlreport.Row07
     * @see com.github.svrtm.xlreport.Row07XL
     */
    public TR row(final int i) {
        return newRowInstance(i, GET);
    }

    TR rowOrCreateIfAbsent(final int i) {
        return newRowInstance(i, CREATE_and_GET);
    }

    /**
     * Returns the number of physically defined rows (NOT the number of rows
     * in the sheet).
     *
     * @return the number of physically defined rows in this sheet.
     */
    public int indexOfLastRow() {
        return sheet.getPhysicalNumberOfRows();
    }

    /**
     * Adds a merged region of cells (hence those cells form one)
     *
     * @param firstRow
     *            Index of first row
     * @param lastRow
     *            Index of last row (inclusive), must be equal to or larger than
     *            {@code firstRow}
     * @param firstCol
     *            Index of first column
     * @param lastCol
     *            Index of last column (inclusive), must be equal to or larger
     *            than {@code firstCol}
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB mergedRegion(final int firstRow, final int lastRow,
                           final int firstCol, final int lastCol) {
        if (regionsList == null)
            regionsList = new ArrayList<CellRangeAddress>(1);

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
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB iColumnWidthIs(final int columnIndex, final int width) {
        sheet.setColumnWidth(columnIndex, width * 256);
        return (HB) this;
    }

    /**
     * Set the width
     *
     * @param widthColumns
     *            - [columnIndex, width]
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB columnWidths(final int widthColumns[][]) {
        for (final int[] is : widthColumns)
            sheet.setColumnWidth(is[0], is[1] * 256);

        return (HB) this;
    }

    /**
     * This process can be relatively slow on large sheets, so this should
     * normally only be called once per column, at the end of your processing.
     *
     * @param column
     *            the column index
     * @return this
     */
    @SuppressWarnings("unchecked")
    public HB withAutoSizeColumn(final int column) {
        sheet.autoSizeColumn(column);
        return (HB) this;
    }

    // ********************************************************************** //
    // ********************************************************************** //

    private TR newRowInstance(final int i, final RowOperation oper) {
        try {
            @SuppressWarnings("unchecked")
            final Constructor<TR> c = (Constructor<TR>) rowClass
                    .getDeclaredConstructor(ABuilder.class, int.class,
                            RowOperation.class);
            return c.newInstance(this, i, oper);
        }
        catch (final Exception e) {
            throw new ReportBuilderException(e);
        }
    }

    private TR newRowInstance() {
        try {
            @SuppressWarnings("unchecked")
            final Constructor<TR> c = (Constructor<TR>) rowClass
                    .getDeclaredConstructor(ABuilder.class);
            return c.newInstance(this);
        }
        catch (final Exception e) {
            throw new ReportBuilderException(e);
        }
    }
}
