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

/**
 * @author Artem.Smirnov
 */
final public class CellStyle<HB> extends ACellStyle<CellStyle<HB>, Cell<HB>> {

    CellStyle(final Cell<HB> cell) {
        super(cell);
    }

    @Override
    public Cell<HB> createStyle() {
        final org.apache.poi.ss.usermodel.CellStyle poiStyle = getStyle();
        cell.poiCell.setCellStyle(poiStyle);

        return cell;
    }

}
