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
public abstract class ACell<HB, TC> {
    private static int AUTOSIZE_MIN_LENGTH = 7;

    final Row<HB, ?> row;
    final ABuilder<HB, ?> builder;

    /**
     * This problem is caused by the fact that sometimes javac's implementation
     * of JLS3 15.12.2.8 ignores recursive bounds, sometimes not (as in this
     * case). When recursive bounds contains wildcards, such bounds are included
     * when computing uninferred type variables. This makes subsequent subtyping
     * test (Integer <: Comparable<? super T> where T is a type-variable to be
     * inferred).
     * Will be fixed after 6369605 [JDK 1.7]
     * (http://bugs.sun.com/bugdatabase/view_bug.do?bug_id=6369605)
     */
    // protected ACellStyle<? extends ACell<?, HB>> cellStyle;
    @SuppressWarnings("rawtypes")
    ACellStyle cellStyle;

    boolean enableAutoSize;

    ACell(final Row<HB, ?> row) {
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
                    final Row<HB, ?> row =
                                         builder.rowOrCreateIfAbsent(rowIndex);
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
     * @return
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
