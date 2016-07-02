/**
 * <pre>
 * Copyright © 2016 Artem Smirnov
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

/**
 * @author Artem.Smirnov
 */
final public class Cells97<HB> extends Cells<HB, Row97<HB>, Cells97<HB>> {

    Cells97(final Row97<HB> row, final int[] indexesCells) {
        super(row, indexesCells);
    }

    /**
     * Prepare style for group of cells.
     * Style is applied to cells without styles. Will be ignored by all of the
     * cells with styles
     *
     * @return an instance of the class <code>CellsStyle97</code>.
     * @see com.github.svrtm.xlreport.CellsStyle97
     */
    @SuppressWarnings("unchecked")
    public CellsStyle97<HB> prepareStyle() {
        if (cellStyle == null)
            cellStyle = new CellsStyle97<HB>(this);
        return (CellsStyle97<HB>) cellStyle;
    }

}
