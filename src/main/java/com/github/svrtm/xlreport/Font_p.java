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

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * @author Artem.Smirnov
 */
final class Font_p {
    private Short fontHeightInPoints;
    private Short color;
    private XSSFColor xssfColor;
    private Boolean italic;
    private Short boldweight;
    private String name;

    /**
     * @param fontHeightInPoints
     *            the fontHeightInPoints to set
     */
    public void setFontHeightInPoints(final short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }

    /**
     * @param color
     *            the color to set
     */
    public void setColor(final short color) {
        this.color = color;
    }

    /**
     * Sets the Red Green Blue or Alpha Red Green Blue
     *
     * @param rgb
     *            the color to set
     */
    public void setColor(final byte[] rgb) {
        xssfColor = new XSSFColor(rgb);
    }

    /**
     * @param italic
     *            the italic to set
     */
    public void setItalic(final boolean italic) {
        this.italic = italic;
    }

    /**
     * @param boldweight
     *            the boldweight to set
     */
    public void setBoldweight(final short boldweight) {
        this.boldweight = boldweight;
    }

    /**
     * set the name for the font (i.e. Arial).
     *
     * @param name
     *            - value representing the name of the font to use
     */
    public void setName(final String name) {
        this.name = name;
    }

    /*
     * (non-Javadoc)
     * @see java.lang.Object#hashCode()
     */
    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result
                 + (boldweight == null ? 0 : boldweight.hashCode());
        result = prime * result + (color == null ? 0 : color.hashCode());
        result = prime * result
                 + (fontHeightInPoints == null ? 0
                                               : fontHeightInPoints.hashCode());
        result = prime * result + (italic == null ? 0 : italic.hashCode());
        result = prime * result + (name == null ? 0 : name.hashCode());
        result =
               prime * result + (xssfColor == null ? 0 : xssfColor.hashCode());
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
        final Font_p other = (Font_p) obj;
        if (boldweight == null) {
            if (other.boldweight != null)
                return false;
        }
        else if (!boldweight.equals(other.boldweight))
            return false;
        if (color == null) {
            if (other.color != null)
                return false;
        }
        else if (!color.equals(other.color))
            return false;
        if (fontHeightInPoints == null) {
            if (other.fontHeightInPoints != null)
                return false;
        }
        else if (!fontHeightInPoints.equals(other.fontHeightInPoints))
            return false;
        if (italic == null) {
            if (other.italic != null)
                return false;
        }
        else if (!italic.equals(other.italic))
            return false;
        if (name == null) {
            if (other.name != null)
                return false;
        }
        else if (!name.equals(other.name))
            return false;
        if (xssfColor == null) {
            if (other.xssfColor != null)
                return false;
        }
        else if (!xssfColor.equals(other.xssfColor))
            return false;
        return true;
    }

    public void copyTo(final org.apache.poi.ss.usermodel.Font poiFont) {
        if (fontHeightInPoints != null)
            poiFont.setFontHeightInPoints(fontHeightInPoints);
        if (xssfColor == null) {
            if (color != null)
                poiFont.setColor(color);
        }
        else
            ((XSSFFont) poiFont).setColor(xssfColor);
        if (italic != null)
            poiFont.setItalic(italic);
        if (boldweight != null)
            poiFont.setBoldweight(boldweight);
        if (name != null)
            poiFont.setFontName(name);
    }

}
