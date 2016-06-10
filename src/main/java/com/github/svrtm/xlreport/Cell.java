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
import static com.github.svrtm.xlreport.Cell.CellOperation.GET_CREATE;
import static java.lang.String.format;
import static java.util.Locale.ENGLISH;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.text.NumberFormat;

import org.apache.poi.ss.usermodel.CreationHelper;

/**
 * @author Artem.Smirnov
 */
final public class Cell<HB> extends ACell<Cell<HB>, HB> {
    private boolean decimalFormat;
    private BigDecimal bigDecimalValue;
    private final CreationHelper creationHelper;
    final org.apache.poi.ss.usermodel.Cell poiCell;

    enum CellOperation {
        CREATE, GET, GET_CREATE
    }

    Cell(final Row<HB> row, final int i, final CellOperation cellOperation) {
        super(row);
        creationHelper = builder.sheet.getWorkbook().getCreationHelper();

        final org.apache.poi.ss.usermodel.Row poiRow = row.poiRow;
        if (CREATE == cellOperation)
            poiCell = poiRow.createCell(i);
        else
            poiCell = GET_CREATE == cellOperation
                      && poiRow.getCell(i) == null ? poiRow.createCell(i)
                                                   : poiRow.getCell(i);
    }

    public Row<HB> buildCell() {
        if (decimalFormat)
            valueWithDecimalFormat();
        else if (bigDecimalValue != null)
            withValue(bigDecimalValue.toString());
        if (enableAutoSize)
            setAutoSizeColumn(poiCell);

        return row;
    }

    @SuppressWarnings("unchecked")
    public CellStyle<HB> withStyle() {
        if (cellStyle == null)
            cellStyle = new CellStyle<HB>(this);
        return (CellStyle<HB>) cellStyle;
    }

    public Cell<HB> withValue(final String value) {
        poiCell.setCellValue(value);
        return this;
    }

    public Cell<HB> withRichText(final String value) {
        poiCell.setCellValue(creationHelper.createRichTextString(value));
        return this;
    }

    public Cell<HB> withValue(final int value) {
        poiCell.setCellValue(value);
        return this;
    }

    public Cell<HB> withValue(final double value) {
        poiCell.setCellValue(value);
        return this;
    }

    public Cell<HB> withValue(final BigDecimal value) {
        bigDecimalValue = value;
        return this;
    }

    public Cell<HB> withValue(final BigInteger value) {
        withValue(value.toString());
        return this;
    }

    /**
     * The value can be of type String, Long, BigDecimal or BigInteger
     *
     * @param value
     * @return
     */
    public Cell<HB> withValue(final Object value) {
        if (value instanceof String)
            withValue((String) value);
        else if (value instanceof Long)
            withValue(((Long) value).longValue());
        else if (value instanceof BigDecimal)
            withValue((BigDecimal) value);
        else if (value instanceof BigInteger)
            withValue((BigInteger) value);
        else
            throw new ReportBuilderException(format("Type is not supported: %s",
                    value.getClass().getName()));
        return this;
    }

    /**
     * Specialization of format
     * <p>
     * Cell will be of type <b>String</b> with fixed number of decimal places
     * (use a pattern "0.00").<br/>
     * Or the <b>Number</b> type if is used by ACellStyle#dataFormat
     *
     * @return this
     * @see ACellStyle#dataFormat(String)
     */
    public Cell<HB> toDecimalFormat() {
        decimalFormat = true;
        return this;
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

        if (cellStyle != null && cellStyle.cellStyle_p.getDataFormat() != 0) {
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
