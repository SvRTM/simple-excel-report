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
public interface CellStyle {
    public enum Alignment {
        /** */
        CENTER(org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER, AlignmentType.H),
        /** */
        RIGHT(org.apache.poi.ss.usermodel.CellStyle.ALIGN_RIGHT, AlignmentType.H),
        /** */
        LEFT(org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT, AlignmentType.H),
        /** */
        VERTICAL_CENTER(org.apache.poi.ss.usermodel.CellStyle.VERTICAL_CENTER, AlignmentType.V);

        enum AlignmentType {
            H, V
        };

        short idx;
        private final AlignmentType alignmentType;

        private Alignment(final short idx, final AlignmentType alignmentType) {
            this.idx = idx;
            this.alignmentType = alignmentType;
        }

        public AlignmentType getAlignmentType() {
            return alignmentType;
        }
    }

    public enum FillPattern {
        /** */
        NO_FILL(org.apache.poi.ss.usermodel.CellStyle.NO_FILL),
        /** */
        SOLID_FOREGROUND(org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND),
        /** */
        FINE_DOTS(org.apache.poi.ss.usermodel.CellStyle.FINE_DOTS),
        /** */
        ALT_BARS(org.apache.poi.ss.usermodel.CellStyle.ALT_BARS),
        /** */
        SPARSE_DOTS(org.apache.poi.ss.usermodel.CellStyle.SPARSE_DOTS),
        /** */
        THICK_HORZ_BANDS(org.apache.poi.ss.usermodel.CellStyle.THICK_HORZ_BANDS),
        /** */
        THICK_VERT_BANDS(org.apache.poi.ss.usermodel.CellStyle.THICK_VERT_BANDS),
        /** */
        THICK_BACKWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THICK_BACKWARD_DIAG),
        /** */
        THICK_FORWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THICK_FORWARD_DIAG),
        /** */
        BIG_SPOTS(org.apache.poi.ss.usermodel.CellStyle.BIG_SPOTS),
        /** */
        BRICKS(org.apache.poi.ss.usermodel.CellStyle.BRICKS),
        /** */
        THIN_HORZ_BANDS(org.apache.poi.ss.usermodel.CellStyle.THIN_HORZ_BANDS),
        /** */
        THIN_VERT_BANDS(org.apache.poi.ss.usermodel.CellStyle.THIN_VERT_BANDS),
        /** */
        THIN_BACKWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THIN_BACKWARD_DIAG),
        /** */
        THIN_FORWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THIN_FORWARD_DIAG),
        /** */
        SQUARES(org.apache.poi.ss.usermodel.CellStyle.SQUARES),
        /** */
        DIAMONDS(org.apache.poi.ss.usermodel.CellStyle.DIAMONDS);

        short idx;

        private FillPattern(final short idx) {
            this.idx = idx;
        }
    }

}
