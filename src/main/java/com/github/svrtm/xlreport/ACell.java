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

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author Artem.Smirnov
 */
abstract class ACell<HB, TC extends ACell<HB, TC>> {
    private static int AUTOSIZE_MIN_LENGTH = 7;

    final Row<HB, ?, ?, ?> row;
    final ABuilder<HB, ?> builder;

    ACellStyle<TC, ?, ?> cellStyle;

    boolean enableAutoSize;

    ACell(final Row<HB, ?, ?, ?> row) {
        this.row = row;
        this.builder = row.builder;
    }

    List<Cell> findMergedCells(final Cell poiCell) {
        if (builder.regionsList == null)
            return null;

        List<Cell> mergedCells;
        int rowIndex = poiCell.getRowIndex();
        for (final CellRangeAddress region : builder.regionsList)
            if (region.isInRange(rowIndex, poiCell.getColumnIndex())) {
                final int firstCol = region.getFirstColumn();
                final int lastCol = region.getLastColumn();
                final int firstRow = region.getFirstRow();
                final int lastRow = region.getLastRow();
                mergedCells = new ArrayList<Cell>(region.getNumberOfCells());
                for (rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
                    final Row<HB, ?, ?, ?> row;
                    row = builder.rowOrCreateIfAbsent(rowIndex);
                    for (int colIdx = firstCol; colIdx <= lastCol; colIdx++) {
                        final Cell mergedCell;
                        mergedCell = row.cellOrCreateIfAbsent(colIdx).poiCell;
                        mergedCells.add(mergedCell);
                    }
                }

                return mergedCells;
            }

        return null;
    }

    /**
     * Adjusts the column width to fit the contents.<br>
     * This process can be relatively slow on large sheets, so this should
     * normally only be called once per column, at the end of your processing.
     *
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC useAutoSizeColumn() {
        enableAutoSize = true;
        return (TC) this;
    }

    void setAutoSizeColumn(final Cell poiCell) {
        boolean enableAutoSize = false;
        final int cellType = poiCell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_STRING: {
                final String value = poiCell.getStringCellValue();
                if (value != null && value.length() >= AUTOSIZE_MIN_LENGTH) {
                    final String tr = value.trim();
                    for (int i = 0; i < tr.length(); i++)
                        if (Character.isWhitespace(tr.charAt(i)) == false) {
                            enableAutoSize = true;
                            break;
                        }
                }
                break;
            }

            case Cell.CELL_TYPE_BLANK:
                break;

            default:
                enableAutoSize = true;
        }
        if (enableAutoSize)
            builder.sheet.autoSizeColumn(poiCell.getColumnIndex());
    }
}
