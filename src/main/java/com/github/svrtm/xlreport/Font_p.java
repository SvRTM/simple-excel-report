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
final class Font_p {
    private short fontHeightInPoints;
    private short color;
    private boolean italic;
    private short boldweight;

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

    /*
     * (non-Javadoc)
     * @see java.lang.Object#hashCode()
     */
    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + boldweight;
        result = prime * result + color;
        result = prime * result + fontHeightInPoints;
        result = prime * result + (italic ? 1231 : 1237);
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
        if (!(obj instanceof Font_p))
            return false;
        final Font_p other = (Font_p) obj;
        if (boldweight != other.boldweight)
            return false;
        if (color != other.color)
            return false;
        if (fontHeightInPoints != other.fontHeightInPoints)
            return false;
        if (italic != other.italic)
            return false;
        return true;
    }

    public void copyTo(final org.apache.poi.ss.usermodel.Font poiFont) {
        poiFont.setFontHeightInPoints(fontHeightInPoints);
        poiFont.setColor(color);
        poiFont.setItalic(italic);
        poiFont.setBoldweight(boldweight);
    }
}
