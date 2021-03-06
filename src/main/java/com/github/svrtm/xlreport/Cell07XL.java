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
final public class Cell07XL<HB> extends Cell<HB, Row07XL<HB>, Cell07XL<HB>> {

    Cell07XL(final Row07XL<HB> row, final int i,
             final CellOperation cellOperation) {
        super(row, i, cellOperation);
    }

    /**
     * Prepare style for the current Cell.
     *
     * @return an instance of the class <code>CellStyle07XL</code>.
     * @see com.github.svrtm.xlreport.CellStyle07XL
     */
    @SuppressWarnings("unchecked")
    public CellStyle07XL<HB> prepareStyle() {
        if (cellStyle == null)
            cellStyle = new CellStyle07XL<HB>(this);
        return (CellStyle07XL<HB>) cellStyle;
    }
}
