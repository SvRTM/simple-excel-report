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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * @author Artem.Smirnov
 */
final class CellStyle_p {
    private Short verticalAlignment;
    private Short alignment;

    private Short borderBottom;
    private Short bottomBorderColor;
    private Short borderLeft;
    private Short leftBorderColor;
    private Short borderRight;
    private Short rightBorderColor;
    private Short borderTop;
    private Short topBorderColor;

    private Boolean wrapText;

    private Short fillPattern;
    private Short fillForegroundColor;
    private XSSFColor xssfFgColor;

    private Short fmt;

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
     * Sets the Red Green Blue or Alpha Red Green Blue
     *
     * @param fgRgb
     *            the fillForegroundColor to set
     */
    public void setFillForegroundColor(final byte[] fgRgb) {
        xssfFgColor = new XSSFColor(fgRgb);
    }

    /**
     * @return
     *         get the index of the format
     */
    public Short getDataFormat() {
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
        result =
               prime * result + (alignment == null ? 0 : alignment.hashCode());
        result = prime * result
                 + (borderBottom == null ? 0 : borderBottom.hashCode());
        result = prime * result
                 + (borderLeft == null ? 0 : borderLeft.hashCode());
        result = prime * result
                 + (borderRight == null ? 0 : borderRight.hashCode());
        result =
               prime * result + (borderTop == null ? 0 : borderTop.hashCode());
        result = prime * result
                 + (bottomBorderColor == null ? 0
                                              : bottomBorderColor.hashCode());
        result = prime * result
                 + (fillForegroundColor == null ? 0
                                                : fillForegroundColor
                                                        .hashCode());
        result = prime * result
                 + (fillPattern == null ? 0 : fillPattern.hashCode());
        result = prime * result + (fmt == null ? 0 : fmt.hashCode());
        result = prime * result + (font_p == null ? 0 : font_p.hashCode());
        result = prime * result
                 + (leftBorderColor == null ? 0 : leftBorderColor.hashCode());
        result = prime * result
                 + (rightBorderColor == null ? 0 : rightBorderColor.hashCode());
        result = prime * result
                 + (topBorderColor == null ? 0 : topBorderColor.hashCode());
        result = prime * result
                 + (verticalAlignment == null ? 0
                                              : verticalAlignment.hashCode());
        result = prime * result + (wrapText == null ? 0 : wrapText.hashCode());
        result = prime * result
                 + (xssfFgColor == null ? 0 : xssfFgColor.hashCode());
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
        if (getClass() != obj.getClass())
            return false;
        final CellStyle_p other = (CellStyle_p) obj;
        if (alignment == null) {
            if (other.alignment != null)
                return false;
        }
        else if (!alignment.equals(other.alignment))
            return false;
        if (borderBottom == null) {
            if (other.borderBottom != null)
                return false;
        }
        else if (!borderBottom.equals(other.borderBottom))
            return false;
        if (borderLeft == null) {
            if (other.borderLeft != null)
                return false;
        }
        else if (!borderLeft.equals(other.borderLeft))
            return false;
        if (borderRight == null) {
            if (other.borderRight != null)
                return false;
        }
        else if (!borderRight.equals(other.borderRight))
            return false;
        if (borderTop == null) {
            if (other.borderTop != null)
                return false;
        }
        else if (!borderTop.equals(other.borderTop))
            return false;
        if (bottomBorderColor == null) {
            if (other.bottomBorderColor != null)
                return false;
        }
        else if (!bottomBorderColor.equals(other.bottomBorderColor))
            return false;
        if (fillForegroundColor == null) {
            if (other.fillForegroundColor != null)
                return false;
        }
        else if (!fillForegroundColor.equals(other.fillForegroundColor))
            return false;
        if (fillPattern == null) {
            if (other.fillPattern != null)
                return false;
        }
        else if (!fillPattern.equals(other.fillPattern))
            return false;
        if (fmt == null) {
            if (other.fmt != null)
                return false;
        }
        else if (!fmt.equals(other.fmt))
            return false;
        if (font_p == null) {
            if (other.font_p != null)
                return false;
        }
        else if (!font_p.equals(other.font_p))
            return false;
        if (leftBorderColor == null) {
            if (other.leftBorderColor != null)
                return false;
        }
        else if (!leftBorderColor.equals(other.leftBorderColor))
            return false;
        if (rightBorderColor == null) {
            if (other.rightBorderColor != null)
                return false;
        }
        else if (!rightBorderColor.equals(other.rightBorderColor))
            return false;
        if (topBorderColor == null) {
            if (other.topBorderColor != null)
                return false;
        }
        else if (!topBorderColor.equals(other.topBorderColor))
            return false;
        if (verticalAlignment == null) {
            if (other.verticalAlignment != null)
                return false;
        }
        else if (!verticalAlignment.equals(other.verticalAlignment))
            return false;
        if (wrapText == null) {
            if (other.wrapText != null)
                return false;
        }
        else if (!wrapText.equals(other.wrapText))
            return false;
        if (xssfFgColor == null) {
            if (other.xssfFgColor != null)
                return false;
        }
        else if (!xssfFgColor.equals(other.xssfFgColor))
            return false;
        return true;
    }

    public void copyTo(final CellStyle poiStyle) {
        if (verticalAlignment != null)
            poiStyle.setVerticalAlignment(verticalAlignment);
        if (alignment != null)
            poiStyle.setAlignment(alignment);

        if (borderBottom != null)
            poiStyle.setBorderBottom(borderBottom);
        if (bottomBorderColor != null)
            poiStyle.setBottomBorderColor(bottomBorderColor);
        if (borderLeft != null)
            poiStyle.setBorderLeft(borderLeft);
        if (leftBorderColor != null)
            poiStyle.setLeftBorderColor(leftBorderColor);
        if (borderRight != null)
            poiStyle.setBorderRight(borderRight);
        if (rightBorderColor != null)
            poiStyle.setRightBorderColor(rightBorderColor);
        if (borderTop != null)
            poiStyle.setBorderTop(borderTop);
        if (topBorderColor != null)
            poiStyle.setTopBorderColor(topBorderColor);

        if (wrapText != null)
            poiStyle.setWrapText(wrapText);

        if (fillPattern != null)
            poiStyle.setFillPattern(fillPattern);
        if (xssfFgColor == null) {
            if (fillForegroundColor != null)
                poiStyle.setFillForegroundColor(fillForegroundColor);
        }
        else
            ((XSSFCellStyle) poiStyle).setFillForegroundColor(xssfFgColor);

        if (fmt != null)
            poiStyle.setDataFormat(fmt);
    }
}
