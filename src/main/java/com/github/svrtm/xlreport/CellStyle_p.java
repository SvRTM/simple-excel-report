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

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @author Artem.Smirnov
 */
final class CellStyle_p {

    private short verticalAlignment;
    private short alignment;

    private short borderBottom;
    private short bottomBorderColor;
    private short borderLeft;
    private short leftBorderColor;
    private short borderRight;
    private short rightBorderColor;
    private short borderTop;
    private short topBorderColor;

    private boolean wrapText;

    private short fillPattern;
    private short fillForegroundColor;

    private short fmt;

    Font_p font_p;

    /**
     * @param verticalAlignment
     *            the verticalAlignment to set
     */
    public void setVerticalAlignment(final short verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    /**
     * @param alignment
     *            the alignment to set
     */
    public void setAlignment(final short alignment) {
        this.alignment = alignment;
    }

    /**
     * @param borderBottom
     *            the borderBottom to set
     */
    public void setBorderBottom(final short borderBottom) {
        this.borderBottom = borderBottom;
    }

    /**
     * @param bottomBorderColor
     *            the bottomBorderColor to set
     */
    public void setBottomBorderColor(final short bottomBorderColor) {
        this.bottomBorderColor = bottomBorderColor;
    }

    /**
     * @param borderLeft
     *            the borderLeft to set
     */
    public void setBorderLeft(final short borderLeft) {
        this.borderLeft = borderLeft;
    }

    /**
     * @param leftBorderColor
     *            the leftBorderColor to set
     */
    public void setLeftBorderColor(final short leftBorderColor) {
        this.leftBorderColor = leftBorderColor;
    }

    /**
     * @param borderRight
     *            the borderRight to set
     */
    public void setBorderRight(final short borderRight) {
        this.borderRight = borderRight;
    }

    /**
     * @param rightBorderColor
     *            the rightBorderColor to set
     */
    public void setRightBorderColor(final short rightBorderColor) {
        this.rightBorderColor = rightBorderColor;
    }

    /**
     * @param borderTop
     *            the borderTop to set
     */
    public void setBorderTop(final short borderTop) {
        this.borderTop = borderTop;
    }

    /**
     * @param topBorderColor
     *            the topBorderColor to set
     */
    public void setTopBorderColor(final short topBorderColor) {
        this.topBorderColor = topBorderColor;
    }

    /**
     * @param wrapText
     *            the wrapText to set
     */
    public void setWrapText(final boolean wrapText) {
        this.wrapText = wrapText;
    }

    /**
     * @param fillPattern
     *            the fillPattern to set
     */
    public void setFillPattern(final short fillPattern) {
        this.fillPattern = fillPattern;
    }

    /**
     * @param fillForegroundColor
     *            the fillForegroundColor to set
     */
    public void setFillForegroundColor(final short fillForegroundColor) {
        this.fillForegroundColor = fillForegroundColor;
    }

    /**
     * @return
     *         get the index of the format
     */
    public short getDataFormat() {
        return fmt;
    }

    /**
     * set the data format (must be a valid format)
     *
     * @param dataFormat
     *            the
     */
    public void setDataFormat(final short fmt) {
        this.fmt = fmt;
    }

    /*
     * (non-Javadoc)
     * @see java.lang.Object#hashCode()
     */
    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + alignment;
        result = prime * result + borderBottom;
        result = prime * result + borderLeft;
        result = prime * result + borderRight;
        result = prime * result + borderTop;
        result = prime * result + bottomBorderColor;
        result = prime * result + fillForegroundColor;
        result = prime * result + fillPattern;
        result = prime * result + fmt;
        result = prime * result + (font_p == null ? 0 : font_p.hashCode());
        result = prime * result + leftBorderColor;
        result = prime * result + rightBorderColor;
        result = prime * result + topBorderColor;
        result = prime * result + verticalAlignment;
        result = prime * result + (wrapText ? 1231 : 1237);
        return result;
    }

    /*
     * (non-Javadoc)
     * @see java.lang.Object#equals(java.lang.Object)
     */
    @Override
    public boolean equals(final Object obj) {
        if (this == obj)
            return true;
        if (obj == null)
            return false;
        if (!(obj instanceof CellStyle_p))
            return false;
        final CellStyle_p other = (CellStyle_p) obj;
        if (alignment != other.alignment)
            return false;
        if (borderBottom != other.borderBottom)
            return false;
        if (borderLeft != other.borderLeft)
            return false;
        if (borderRight != other.borderRight)
            return false;
        if (borderTop != other.borderTop)
            return false;
        if (bottomBorderColor != other.bottomBorderColor)
            return false;
        if (fillForegroundColor != other.fillForegroundColor)
            return false;
        if (fillPattern != other.fillPattern)
            return false;
        if (fmt != other.fmt)
            return false;
        if (font_p == null) {
            if (other.font_p != null)
                return false;
        }
        else if (!font_p.equals(other.font_p))
            return false;
        if (leftBorderColor != other.leftBorderColor)
            return false;
        if (rightBorderColor != other.rightBorderColor)
            return false;
        if (topBorderColor != other.topBorderColor)
            return false;
        if (verticalAlignment != other.verticalAlignment)
            return false;
        if (wrapText != other.wrapText)
            return false;
        return true;
    }

    public void copyTo(final CellStyle poiStyle) {
        poiStyle.setVerticalAlignment(verticalAlignment);
        poiStyle.setAlignment(alignment);

        poiStyle.setBorderBottom(borderBottom);
        poiStyle.setBottomBorderColor(bottomBorderColor);
        poiStyle.setBorderLeft(borderLeft);
        poiStyle.setLeftBorderColor(leftBorderColor);
        poiStyle.setBorderRight(borderRight);
        poiStyle.setRightBorderColor(rightBorderColor);
        poiStyle.setBorderTop(borderTop);
        poiStyle.setTopBorderColor(topBorderColor);

        poiStyle.setWrapText(wrapText);

        poiStyle.setFillPattern(fillPattern);
        poiStyle.setFillForegroundColor(fillForegroundColor);

        poiStyle.setDataFormat(fmt);
    }
}