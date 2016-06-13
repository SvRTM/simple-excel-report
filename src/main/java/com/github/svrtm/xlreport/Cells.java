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

import static java.lang.String.format;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @author Artem.Smirnov
 */
public abstract class Cells<HB, TR extends Row<HB, TR, ?, TCs>, TCs extends Cells<HB, TR, TCs>>
        extends ACell<HB, TCs> {
    private int columnWidth = -1;
    private int incrementValue = -1;
    private final int[] indexesCells;

    Cells(final TR row, final int[] indexesCells) {
        super(row);
        this.indexesCells = indexesCells;
    }

    @SuppressWarnings("unchecked")
    public TR configureCells() {
        if (columnWidth == -1 && incrementValue == -1 && cellStyle == null)
            return (TR) row;

        int increment = incrementValue;
        for (final int i : indexesCells) {
            if (row.cells.get(i) == null)
                row.prepareNewCell(i).createCell();

            final org.apache.poi.ss.usermodel.Row poiRow = row.poiRow;
            final Cell poiCell = poiRow.getCell(i);
            if (poiCell == null)
                throw new ReportBuilderException(
                        format("A cell of number %d [row:%d] can't be found. Please, create a cell before using it",
                                i, poiRow.getRowNum()));

            if (columnWidth != -1)
                poiCell.getSheet().setColumnWidth(i, columnWidth);
            if (incrementValue != -1)
                poiCell.setCellValue(increment++);
            if (enableAutoSize)
                setAutoSizeColumn(poiCell);

            if (cellStyle != null)
                if (row.cells.get(i).cellStyle == null) {
                    // Apply a style to the cell
                    final CellStyle poiStyle = cellStyle.getStyle();
                    poiCell.setCellStyle(poiStyle);

                    final List<Cell> mergedCells = findMergedCells(poiCell);
                    if (mergedCells != null)
                        for (final Cell mergedCell : mergedCells)
                            mergedCell.setCellStyle(poiStyle);
                }
        }

        return (TR) row;
    }

    @SuppressWarnings("unchecked")
    public TCs whereColumnWidthIs(final int columnWidth) {
        this.columnWidth = columnWidth * 256;
        return (TCs) this;
    }

    @SuppressWarnings("unchecked")
    public TCs withInitialValueOfIncrement(final int start) {
        incrementValue = start;
        return (TCs) this;
    }

}
