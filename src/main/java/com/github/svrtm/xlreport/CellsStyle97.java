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
final public class CellsStyle97<HB> extends
        ACellStyle<Cells97<HB>, CellsStyle97<HB>, Font<CellsStyle97<HB>>> {

    CellsStyle97(final Cells97<HB> cells) {
        super(cells);
    }

    @Override
    public Cells97<HB> createStyle() {
        return cell;
    }
}
