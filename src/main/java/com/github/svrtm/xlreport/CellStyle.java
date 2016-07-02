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

    /**
     * Specifies whether the object in the cell is right-aligned, left-aligned,
     * centered (horizontal / vertical).
     *
     * @author Artem.Smirnov
     */
    public enum Alignment {
        /** Center horizontal alignment */
        CENTER(org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER, AlignmentType.H),
        /** Right-justified horizontal alignment */
        RIGHT(org.apache.poi.ss.usermodel.CellStyle.ALIGN_RIGHT, AlignmentType.H),
        /** Left-justified horizontal alignment */
        LEFT(org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT, AlignmentType.H),
        /** Center-aligned vertical alignment */
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

        AlignmentType getAlignmentType() {
            return alignmentType;
        }
    }

    /**
     * The enumeration value indicating the style of fill pattern being used for
     * a cell format.
     *
     * @author Artem.Smirnov
     */
    public enum FillPattern {
        /** No background */
        NO_FILL(org.apache.poi.ss.usermodel.CellStyle.NO_FILL),
        /** Solidly filled */
        SOLID_FOREGROUND(org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND),
        /** Small fine dots */
        FINE_DOTS(org.apache.poi.ss.usermodel.CellStyle.FINE_DOTS),
        /** Wide dots */
        ALT_BARS(org.apache.poi.ss.usermodel.CellStyle.ALT_BARS),
        /** Sparse dots */
        SPARSE_DOTS(org.apache.poi.ss.usermodel.CellStyle.SPARSE_DOTS),
        /** Thick horizontal bands */
        THICK_HORZ_BANDS(org.apache.poi.ss.usermodel.CellStyle.THICK_HORZ_BANDS),
        /** Thick vertical bands */
        THICK_VERT_BANDS(org.apache.poi.ss.usermodel.CellStyle.THICK_VERT_BANDS),
        /** Thick backward facing diagonals */
        THICK_BACKWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THICK_BACKWARD_DIAG),
        /** Thick forward facing diagonals */
        THICK_FORWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THICK_FORWARD_DIAG),
        /** Large spots */
        BIG_SPOTS(org.apache.poi.ss.usermodel.CellStyle.BIG_SPOTS),
        /** Brick-like layout */
        BRICKS(org.apache.poi.ss.usermodel.CellStyle.BRICKS),
        /** Thin horizontal bands */
        THIN_HORZ_BANDS(org.apache.poi.ss.usermodel.CellStyle.THIN_HORZ_BANDS),
        /** Thin vertical bands */
        THIN_VERT_BANDS(org.apache.poi.ss.usermodel.CellStyle.THIN_VERT_BANDS),
        /** Thin backward diagonal */
        THIN_BACKWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THIN_BACKWARD_DIAG),
        /** Thin forward diagonal */
        THIN_FORWARD_DIAG(org.apache.poi.ss.usermodel.CellStyle.THIN_FORWARD_DIAG),
        /** Squares */
        SQUARES(org.apache.poi.ss.usermodel.CellStyle.SQUARES),
        /** Diamonds */
        DIAMONDS(org.apache.poi.ss.usermodel.CellStyle.DIAMONDS);

        short idx;

        private FillPattern(final short idx) {
            this.idx = idx;
        }
    }

}
