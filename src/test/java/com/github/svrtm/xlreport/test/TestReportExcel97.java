package com.github.svrtm.xlreport.test;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.github.svrtm.xlreport.ACellStyle.Alignment;
import com.github.svrtm.xlreport.Report;
import com.github.svrtm.xlreport.Report.AHeader;
import com.github.svrtm.xlreport.Report.Body;
import com.github.svrtm.xlreport.Report.Final;
import com.github.svrtm.xlreport.Report.Header;
import com.github.svrtm.xlreport.Row;


public class TestReportExcel97 {

    @Test
    public void testBigReport() throws Exception  {
        final int nCell = 6;
        final int nRow = 150000;

        Header hr = Report.EXCEL97.createHeader();
        Body bd =
        hr.build(new Report.AHeader() {
            public void builder() {
                mergedRegion(0, 0, 0, nCell-1)
                .addRow()
                    .addCell(0)
                        .withStyle()
                        .buildStyle()
                    .buildCell()
                    .whereHeightInPointsIs(11)
                .buildRow();
                int nLastRow = getIndexOfLastRow();

                addRow(2, new ICallback<AHeader>() {
                    public void step(Row<AHeader> row) {
                        for (int i = 0; i < nCell; i++) {
                            row.addCell(i).buildCell();
                        }
                    }
                })
                .mergedRegion(nLastRow, nLastRow + 1, 0, 0)
                .mergedRegion(nLastRow, nLastRow + 1, 1, 1)
                .mergedRegion(nLastRow, nLastRow, 2, 3)
                .mergedRegion(nLastRow, nLastRow, 4, 5);

                row(nLastRow)
                    .addCell(0)
                        .withValue("-=0=-").buildCell()
                    .addCell(1)
                        .withValue("-=1=-").buildCell()
                    .addCell(2)
                        .withValue("-=2=-").buildCell()
                    .addCell(4)
                        .withValue("-=3=-").buildCell()
                    .cells()
                        .withStyle()
                            .header()
                        .buildStyle()
                        .whereColumnWidthIs(16)
                    .buildCells()
                    .whereHeightInPointsIs(45)
                .buildRow()
                .row(nLastRow+1)
                    .addCell(2)
                        .withValue("-=2.1=-").buildCell()
                    .addCell(3)
                        .withValue("-=2.2=-").buildCell()
                    .addCell(4)
                        .withValue("-=3.1=-").buildCell()
                    .addCell(5)
                        .withValue("-=3.2=-").buildCell()
                    .cells()
                        .withStyle()
                            .header()
                        .buildStyle()
                        .whereColumnWidthIs(16)
                    .buildCells()
                    .whereHeightInPointsIs(45)
                .buildRow()
                .addRow()
                    .addCells(nCell)
                        .withInitialValueOfIncrement(1)
                    .buildCells()
                    .cells()
                        .withStyle()
                            .header()
                        .buildStyle()
                    .buildCells()
                .buildRow()
                .whereiColumnWidthIs(1, 35)
            ;
            }
        });

        for (int i = 0; i < nRow; i++) {
            bd
                .addRow()
                    .addCell(0)
                        .withValue(i)
                    .buildCell()
                    .addCell(1)
                        .withValue("str -> " + i)
                    .buildCell()
                    .addCell(2)
                        .withValue(123.45 + i)
                        .toDecimalFormat()
                    .buildCell()
                    .addCell(3)
                        .withValue(345.55 + i)
                    .buildCell()
                    .addCell(4)
                        .withValue(456.54 + i)
                    .buildCell()
                    .addCell(5)
                        .withValue(567.75 + i)
                    .buildCell()
                    .cells()
                        .withStyle()
                            .defaultBorder()
                            .alignment(Alignment.RIGHT)
                        .buildStyle()
                    .buildCells()
                ;
        }
        bd.withAutoSizeColumn(1).withAutoSizeColumn(2).withAutoSizeColumn(4);

        int iLastRow = bd.getIndexOfLastRow();
        Final report =
        bd
            .mergedRegion(iLastRow, iLastRow, 0, 1)
            .addRow()
                .addCell(0)
                    .withValue("-==-")
                .buildCell()
                .addCell(1)
                .buildCell()
                .addCell(2)
                    .withValue(1234567.23)
                    .toDecimalFormat()
                    .withStyle()
                        .defaultBorder()
                        .alignment(Alignment.RIGHT)
                    .buildStyle()
                .buildCell()
                .addCell(3)
                    .withValue(786541.55)
                    .withStyle()
                        .defaultBorder()
                        .alignment(Alignment.RIGHT)
                    .buildStyle()
                .buildCell()
                .addCell(4)
                    .withValue(4453306.45)
                    .withStyle()
                        .defaultBorder()
                        .alignment(Alignment.RIGHT)
                    .buildStyle()
                .buildCell()
                .addCell(5)
                    .withValue(500567.75)
                    .withStyle()
                        .defaultBorder()
                        .alignment(Alignment.RIGHT)
                    .buildStyle()
                .buildCell()
                .cells()
                    .withStyle()
                        .defaultBorder()
                        .alignment(Alignment.CENTER)
                    .buildStyle()
                .buildCells()
           .buildRow()
       .build()
;
        Workbook wb = report.getWorkbook();
        FileOutputStream out = new FileOutputStream("target/workbook.xls");
        wb.write(out);
        out.close();
    }
}
