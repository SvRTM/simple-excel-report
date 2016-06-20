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
import static java.lang.String.format;
import static java.util.Arrays.sort;

import java.lang.reflect.Constructor;
import java.lang.reflect.ParameterizedType;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import com.github.svrtm.xlreport.Cell.CellOperation;
import com.github.svrtm.xlreport.Cell.INewCell;

/**
 * @author Artem.Smirnov
 */
abstract class Row<HB, TR extends Row<HB, TR, TC, TCs>, TC extends Cell<HB, TR, TC>, TCs extends Cells<HB, TR, TCs>> {
    ABuilder<HB, TR> builder;
    org.apache.poi.ss.usermodel.Row poiRow;

    Map<Integer, TC> cells = new HashMap<Integer, TC>();

    enum RowOperation {
        CREATE, GET, CREATE_and_GET
    }

    Row(final ABuilder<HB, TR> aBuilder) {
        this.builder = aBuilder;
        final int rowIndex = getPhysicalNumberOfRows();

        poiRow = aBuilder.sheet.createRow(rowIndex);
    }

    Row(final ABuilder<HB, TR> aBuilder, final int i,
        final RowOperation rowOperation) {
        this.builder = aBuilder;
        poiRow = aBuilder.sheet.getRow(i);
        if (poiRow != null)
            return;

        if (CREATE == rowOperation)
            poiRow = aBuilder.sheet.createRow(i);
        else if (CREATE_and_GET == rowOperation)
            poiRow = aBuilder.sheet.createRow(i);
        else // GET
            throw new ReportBuilderException(
                    format("A row of number %d can't be found. Please, create a number before using it",
                            i));
    }

    @SuppressWarnings("unchecked")
    public HB configureRow() {
        return (HB) builder;
    }

    /**
     * Get the cell representing a given column (logical cell) 0-based
     *
     * @param i
     *            - 0 based column number
     * @return
     */
    public TC cell(final int i) {
        return createCell(i, CellOperation.GET);
    }

    /**
     * Returns a cell if was absent creates it
     *
     * @param i
     *            - 0 based column number
     * @return
     */
    public TC cellOrCreateIfAbsent(final int i) {
        return createCell(i, CellOperation.CREATE_and_GET);
    }

    /**
     * Create new cells within the row
     *
     * @param i
     *            - the column number this cell represents
     * @return
     */
    public TC prepareNewCell(final int i) {
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
    @SuppressWarnings("unchecked")
    public TR addAndConfigureCells(final int firstCell, final int lastCell,
                                   final INewCell<TC> callback) {
        for (int i = firstCell; i <= lastCell; i++)
            callbackCell(i, callback);
        return (TR) this;
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
    @SuppressWarnings("unchecked")
    public TR addAndConfigureCells(final INewCell<TC> callback,
                                   final int... indexesCells) {
        for (final int i : indexesCells)
            callbackCell(i, callback);
        return (TR) this;
    }

    /**
     * Create new cells within the row
     *
     * @param indexesCells
     *            - the column numbers
     * @return
     */
    /**
     * Create new cells within the row
     *
     * @param indexesCells
     *            - the column numbers
     * @return
     */
    public TCs addCells(final int... indexesCells) {
        return createCells(indexesCells);
    }

    public TCs addCells(final int n) {
        final int[] indexesCells = new int[n];
        for (int i = 0; i < n; i++)
            indexesCells[i] = i;

        return createCells(indexesCells);
    }

    public TCs cells() {
        final Set<Integer> keys = cells.keySet();
        if (keys.size() == 0)
            throw new ReportBuilderException("Cells were not found!");

        final int[] indexesCells = new int[keys.size()];
        int i = 0;
        for (final Integer key : keys)
            indexesCells[i++] = key;
        sort(indexesCells);

        return createCells(indexesCells);
    }

    /**
     * Set the row's height in points.
     *
     * @param height
     * @return - the height in points. <code>-1</code> resets to the default
     *         height
     */
    @SuppressWarnings("unchecked")
    public TR whereHeightInPointsIs(final float height) {
        poiRow.setHeightInPoints(height);
        return (TR) this;
    }

    private TC createCell(final int i, final CellOperation cellOperation) {
        TC cell = cells.get(i);
        if (cell == null) {
            final ParameterizedType pt = (ParameterizedType) getClass()
                    .getGenericSuperclass();
            final ParameterizedType paramType = (ParameterizedType) pt
                    .getActualTypeArguments()[2];
            @SuppressWarnings("unchecked")
            final Class<TC> cellClass = (Class<TC>) paramType.getRawType();
            try {
                final Constructor<TC> c = cellClass.getDeclaredConstructor(
                        getClass(), int.class, CellOperation.class);
                cell = c.newInstance(this, i, cellOperation);
            }
            catch (final Exception e) {
                throw new ReportBuilderException(e);
            }
            cells.put(i, cell);
        }

        return cell;
    }

    private void callbackCell(final int i, final INewCell<TC> callback) {
        final TC newCell = prepareNewCell(i);
        callback.iCell(newCell);
        newCell.createCell();
    }

    private TCs createCells(final int... indexesCells) {
        final ParameterizedType pt = (ParameterizedType) getClass()
                .getGenericSuperclass();
        final ParameterizedType paramType = (ParameterizedType) pt
                .getActualTypeArguments()[3];
        @SuppressWarnings("unchecked")
        final Class<TCs> cellsClass = (Class<TCs>) paramType.getRawType();
        try {
            final Constructor<TCs> c = cellsClass
                    .getDeclaredConstructor(getClass(), int[].class);
            return c.newInstance(this, indexesCells);
        }
        catch (final Exception e) {
            throw new ReportBuilderException(e);
        }
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
        builder.header.prepareHeader();

        rowIndex = builder.sheet.getPhysicalNumberOfRows();
        return rowIndex;
    }
}
