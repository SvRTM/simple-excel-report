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

import static com.github.svrtm.xlreport.Cell.CellOperation.CREATE;
import static com.github.svrtm.xlreport.Cell.CellOperation.CREATE_and_GET;
import static java.lang.String.format;
import static java.util.Locale.ENGLISH;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.text.NumberFormat;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 * @author Artem.Smirnov
 */
public abstract class Cell<HB, TR extends Row<HB, TR, TC, ?>, TC extends Cell<HB, TR, TC>>
        extends ACell<HB, TC> {
    private boolean decimalFormat;
    private BigDecimal bigDecimalValue;
    private final CreationHelper creationHelper;
    final org.apache.poi.ss.usermodel.Cell poiCell;

    /**
     * used in {@link Row#addAndConfigureCells(int, int, INewCell)}
     *
     * @author Artem.Smirnov
     */
    public interface INewCell<TC> {
        /**
         * @param cell
         *            - new cell
         */
        void iCell(final TC cell);
    }

    enum CellOperation {
        CREATE, GET, CREATE_and_GET
    }

    Cell(final TR row, final int i, final CellOperation cellOperation) {
        super(row);
        creationHelper = builder.sheet.getWorkbook().getCreationHelper();

        final org.apache.poi.ss.usermodel.Row poiRow = row.poiRow;
        if (CREATE == cellOperation)
            poiCell = poiRow.createCell(i);
        else
            poiCell = CREATE_and_GET == cellOperation
                      && poiRow.getCell(i) == null ? poiRow.createCell(i)
                                                   : poiRow.getCell(i);
    }

    /**
     * Finalization of the implementation <code>Cell</code>.
     *
     * @return an instance of the implementation <code>Row</code>
     * @see com.github.svrtm.xlreport.Row#prepareNewCell(int)
     */
    @SuppressWarnings("unchecked")
    public TR createCell() {
        if (decimalFormat)
            valueWithDecimalFormat();
        else if (bigDecimalValue != null)
            withValue(bigDecimalValue.toString());
        if (enableAutoSize)
            setAutoSizeColumn(poiCell);

        return (TR) row;
    }

    /**
     * Set a string value for the cell.
     *
     * @param value
     *            value to set the cell to. For formulas we'll set the formula
     *            string, for String cells we'll set its value. For other types
     *            we will change the cell to a string cell and set its value. If
     *            value is null then we will change the cell to a Blank cell.
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withValue(final String value) {
        poiCell.setCellValue(value);
        return (TC) this;
    }

    /**
     * Set a string value for the cell.
     *
     * @param value
     *            value to set the cell to. For formulas we'll set the
     *            'pre-evaluated result string, for String cells we'll set its
     *            value. For other types we will change the cell to a string
     *            cell and set its value. If value is null then we will change
     *            the cell to a Blank cell.
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withRichText(final String value) {
        poiCell.setCellValue(creationHelper.createRichTextString(value));
        return (TC) this;
    }

    /**
     * Set a rich string value for the cell.
     *
     * @param value
     *            value to set the cell to. For formulas we'll set the formula
     *            string, for String cells we'll set its value. For other types
     *            we will change the cell to a string cell and set its value. If
     *            value is null then we will change the cell to a Blank cell.
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withRichText(final RichTextString value) {
        poiCell.setCellValue(value);
        return (TC) this;
    }

    /**
     * Set a numeric value for the cell.
     * Converts integer to a double.
     *
     * @param value
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withValue(final int value) {
        poiCell.setCellValue(value);
        return (TC) this;
    }

    /**
     * Set a numeric value for the cell.
     *
     * @param value
     *            the numeric value to set this cell to. For formulas we'll set
     *            the precalculated value, for numerics we'll set its value. For
     *            other types we will change the cell to a numeric cell and set
     *            its value.
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withValue(final double value) {
        poiCell.setCellValue(value);
        return (TC) this;
    }

    /**
     * Set a numeric value for the cell.
     * Sets the scale of BigDecimal to value in 2 and sets the mode of rounding
     * RoundingMode.Down. The BigDecimal converted to a double.
     *
     * @param value
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withValue(final BigDecimal value) {
        bigDecimalValue = value;
        return (TC) this;
    }

    /**
     * Set a string value for the cell.
     * Converts the numeric value of the BigInteger to its equivalent string
     * representation.
     *
     * @param value
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TC withValue(final BigInteger value) {
        withValue(value.toString());
        return (TC) this;
    }

    /**
     * The value can be of type String, Double, Integer, Long, BigDecimal or
     * BigInteger.
     *
     * @param value
     * @return this
     * @see com.github.svrtm.xlreport.Cell#withRichText(String)
     * @see com.github.svrtm.xlreport.Cell#withValue(double)
     * @see com.github.svrtm.xlreport.Cell#withValue(int)
     * @see com.github.svrtm.xlreport.Cell#withValue(BigDecimal)
     * @see com.github.svrtm.xlreport.Cell#withValue(BigInteger)
     */
    @SuppressWarnings("unchecked")
    public TC withValue(final Object value) {
        if (value == null)
            throw new ReportBuilderException("Value cannot be null");

        else if (value instanceof String)
            withValue((String) value);
        else if (value instanceof Double)
            withValue(((Double) value).doubleValue());
        else if (value instanceof Integer)
            withValue(((Integer) value).intValue());
        else if (value instanceof Long)
            withValue(((Long) value).longValue());
        else if (value instanceof BigDecimal)
            withValue((BigDecimal) value);
        else if (value instanceof BigInteger)
            withValue((BigInteger) value);
        else
            throw new ReportBuilderException(format("Type is not supported: %s",
                    value.getClass().getName()));
        return (TC) this;
    }

    /**
     * Specialization of format.
     * <p>
     * Cell will be of type <b>String</b> with fixed number of decimal places
     * (use a pattern "0.00").<br/>
     * Or the <b>Number</b> type if is used by
     * {@link ACellStyle#dataFormat(String)}.
     *
     * @return this
     * @see ACellStyle#dataFormat(String)
     */
    @SuppressWarnings("unchecked")
    public TC useDecimalFormat() {
        decimalFormat = true;
        return (TC) this;
    }

    private void valueWithDecimalFormat() {
        final double number;
        if (bigDecimalValue != null) {
            bigDecimalValue = bigDecimalValue.setScale(2, RoundingMode.DOWN);
            number = bigDecimalValue.doubleValue();
        }
        else if (CELL_TYPE_NUMERIC == poiCell.getCellType())
            number = poiCell.getNumericCellValue();
        else
            return;

        if (cellStyle != null
            && cellStyle.cellStyle_p.getDataFormat() != null) {
            poiCell.setCellValue(number);
            return;
        }

        final NumberFormat numFormat = NumberFormat.getInstance(ENGLISH);
        numFormat.setGroupingUsed(false);
        numFormat.setMaximumFractionDigits(2);
        numFormat.setRoundingMode(RoundingMode.DOWN);

        final StringBuilder sb = new StringBuilder(numFormat.format(number));
        if (!hasFractionalPart(number))
            sb.append(".00");
        withValue(sb.toString());
    }

    private boolean hasFractionalPart(final double a) {
        return a - Math.floor(a) > 0;
    }

}
