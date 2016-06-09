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
import static com.github.svrtm.xlreport.Row.RowOperation.GET_CREATE;
import static java.lang.String.format;
import static java.util.Arrays.sort;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import com.github.svrtm.xlreport.Cell.CellOperation;

/**
 * @author Artem.Smirnov
 */
final public class Row<HB> {
    ABuilder<HB> builder;
    org.apache.poi.ss.usermodel.Row poiRow;

    protected Map<Integer, Cell<HB>> cells = new HashMap<Integer, Cell<HB>>();

    enum RowOperation {
        CREATE, GET, GET_CREATE
    }

    /**
     * used in {@link Row#addAndBuildCells(int, int, ICallback)
     * Row#addAndBuildCells(ICallback, int...) * }
     *
     * @author Artem.Smirnov
     */
    public interface ICallback {
        /**
         * @param newCell
         *            - new cell
         */
        void step(final Cell<?> newCell);
    }

    Row(final ABuilder<HB> aBuilder) {
        this.builder = aBuilder;
        final int rowIndex = getPhysicalNumberOfRows();

        poiRow = aBuilder.sheet.createRow(rowIndex);
    }

    Row(final ABuilder<HB> aBuilder, final int i,
        final RowOperation rowOperation) {
        this.builder = aBuilder;
        poiRow = aBuilder.sheet.getRow(i);
        if (poiRow != null)
            return;

        if (CREATE == rowOperation)
            poiRow = aBuilder.sheet.createRow(i);
        else if (GET_CREATE == rowOperation)
            poiRow = aBuilder.sheet.createRow(i);
        else
                                             // GET
                                             throw new ReportBuilderException(
                                                     format("A row of number %d can't be found. Please, create a number before using it",
                                                             i));
    }

    @SuppressWarnings("unchecked")
    public HB buildRow() {
        return (HB) builder;
    }

    /**
     * Get the cell representing a given column (logical cell) 0-based
     *
     * @param i
     *            - 0 based column number
     * @return
     */
    public Cell<HB> cell(final int i) {
        return createCell(i, CellOperation.GET);
    }

    /**
     * Returns a cell if was absent creates it
     *
     * @param i
     *            - 0 based column number
     * @return
     */
    public Cell<HB> cellOrCreateIfAbsent(final int i) {
        return createCell(i, CellOperation.GET_CREATE);
    }

    /**
     * Create new cells within the row
     *
     * @param i
     *            - the column number this cell represents
     * @return
     */
    public Cell<HB> addCell(final int i) {
        return createCell(i, CellOperation.CREATE);
    }

    /**
     * Addition of new cells in the current row <br>
     * Method call <b>buildCell</b> is not necessary.
     *
     * @param firstCell
     * @param lastCell
     * @param callback
     * @return
     */
    public Row<HB> addAndBuildCells(final int firstCell, final int lastCell,
                                    final ICallback callback) {
        for (int i = firstCell; i <= lastCell; i++)
            callbackCell(i, callback);
        return this;
    }

    /**
     * Addition of new cells in the current row <br>
     * Method call <b>buildCell</b> is not necessary.
     *
     * @param callback
     * @param indexesCells
     *            indexes of cells
     * @return
     */
    public Row<HB> addAndBuildCells(final ICallback callback,
                                    final int... indexesCells) {
        for (final int i : indexesCells)
            callbackCell(i, callback);
        return this;
    }

    /**
     * Create new cells within the row
     *
     * @param indexesCells
     *            - the column numbers
     * @return
     */
    public Cells<HB> addCells(final int... indexesCells) {
        return new Cells<HB>(this, indexesCells);
    }

    public Cells<HB> addCells(final int n) {
        final int[] indexesCells = new int[n];
        for (int i = 0; i < n; i++)
            indexesCells[i] = i;
        return new Cells<HB>(this, indexesCells);
    }

    public Cells<HB> cells() {
        final Set<Integer> keys = cells.keySet();
        if (keys.size() == 0)
            throw new ReportBuilderException("Cells were not found!");

        final int[] indexesCells = new int[keys.size()];
        int i = 0;
        for (final Integer key : keys)
            indexesCells[i++] = key;
        sort(indexesCells);

        return new Cells<HB>(this, indexesCells);
    }

    /**
     * Set the row's height in points.
     *
     * @param height
     * @return - the height in points. <code>-1</code> resets to the default
     *         height
     */
    public Row<HB> whereHeightInPointsIs(final float height) {
        poiRow.setHeightInPoints(height);
        return this;
    }

    private Cell<HB> createCell(final int i,
                                final CellOperation cellOperation) {
        Cell<HB> cell = cells.get(i);
        if (cell == null) {
            cell = new Cell<HB>(this, i, cellOperation);
            cells.put(i, cell);
        }

        return cell;
    }

    private void callbackCell(final int i, final ICallback callback) {
        final Cell<HB> newCell = addCell(i);
        callback.step(newCell);
        newCell.buildCell();
    }

    /**
     * @return Check limit the number of rows and returns the number of
     *         physically defined rows (NOT the number of rows in the sheet)
     */
    private int getPhysicalNumberOfRows() {
        int rowIndex = builder.sheet.getPhysicalNumberOfRows();
        if (builder.maxrow == -1 || rowIndex >= 0 && rowIndex <= builder.maxrow)
            return rowIndex;

        builder.sheet = builder.wb.createSheet();
        builder.header.sheet = builder.sheet;
        builder.header.builder();

        rowIndex = builder.sheet.getPhysicalNumberOfRows();
        return rowIndex;
    }
}
