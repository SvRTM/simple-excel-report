package com.github.svrtm.xlreport.test;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.github.svrtm.xlreport.Cell.INewCell;
import com.github.svrtm.xlreport.Cell97;
import com.github.svrtm.xlreport.CellStyle.Alignment;
import com.github.svrtm.xlreport.CellStyle.FillPattern;
import com.github.svrtm.xlreport.Excel_97;
import com.github.svrtm.xlreport.Excel_97.Body;
import com.github.svrtm.xlreport.Excel_97.Header;
import com.github.svrtm.xlreport.Final;


public class TestReportExcel_97 {

    @Test
    public void testBigReport() throws Exception  {
        final int nMergedCells = 6;
        final int countRows = 150000;

        Body bd = Excel_97.instanceBody(
        new Excel_97.Header() {
            public void prepareHeader() {
                mergedRegion(0, 0, 0, nMergedCells-1)
                .addNewRow()
                    .prepareNewCell(0)
                        .prepareStyle()
                        .createStyle()
                    .createCell()
                    .whereHeightInPointsIs(11)
                .configureRow();
                int nLastRow = indexOfLastRow();

                addNewRowsWithEmptyCells(2, nMergedCells)
                .mergedRegion(nLastRow, nLastRow + 1, 0, 0)
                .mergedRegion(nLastRow, nLastRow + 1, 1, 1)
                .mergedRegion(nLastRow, nLastRow, 2, 3)
                .mergedRegion(nLastRow, nLastRow, 4, 5);

                row(nLastRow)
                    .prepareNewCell(0)
                        .withValue("-=0=-").createCell()
                    .prepareNewCell(1)
                        .withValue("-=1=-").createCell()
                    .prepareNewCell(2)
                        .withValue("-=2=-").createCell()
                    .prepareNewCell(4)
                        .withValue("-=3=-").createCell()
                    .cells()
                        .prepareStyle()
                            .header()
                        .createStyle()
                        .whereColumnWidthIs(16)
                    .configureCells()
                    .whereHeightInPointsIs(45)
                .configureRow()
                .row(nLastRow+1)
                    .prepareNewCell(2)
                        .withValue("-=2.1=-").createCell()
                    .prepareNewCell(3)
                        .withValue("-=2.2=-").createCell()
                    .prepareNewCell(4)
                        .withValue("-=3.1=-").createCell()
                    .prepareNewCell(5)
                        .withValue("-=3.2=-").createCell()
                    .cells()
                        .prepareStyle()
                            .header()
                        .createStyle()
                        .whereColumnWidthIs(16)
                    .configureCells()
                    .whereHeightInPointsIs(45)
                .configureRow()
                .addNewRow()
                    .addCells(nMergedCells)
                        .withInitialValueOfIncrement(1)
                    .configureCells()
                    .cells()
                        .prepareStyle()
                            .header()
                        .createStyle()
                    .configureCells()
                .configureRow()
                .iColumnWidthIs(1, 35)
            ;
            }
        });

        for (int i = 0; i < countRows; i++) {
            bd
                .addNewRow()
                    .prepareNewCell(0)
                        .withValue(i)
                    .createCell()
                    .prepareNewCell(1)
                        .withValue("str -> " + i)
                    .createCell()
                    .prepareNewCell(2)
                        .withValue(123.45 + i)
                        .useDecimalFormat()
                    .createCell()
                    .prepareNewCell(3)
                        .withValue(345.55 + i)
                    .createCell()
                    .prepareNewCell(4)
                        .withValue(456.54 + i)
                    .createCell()
                    .prepareNewCell(5)
                        .withValue(567.75 + i)
                    .createCell()
                    .cells()
                        .prepareStyle()
                            .defaultEdging()
                            .alignment(Alignment.RIGHT)
                        .createStyle()
                    .configureCells()
                ;
        }
        bd.withAutoSizeColumn(1).withAutoSizeColumn(2).withAutoSizeColumn(4);

        int iLastRow = bd.indexOfLastRow();
        Final report =
        bd
            .mergedRegion(iLastRow, iLastRow, 0, 1)
            .addNewRow()
                .prepareNewCell(0)
                    .withValue("-==-")
                .createCell()
                .prepareNewCell(1)
                .createCell()
                .prepareNewCell(2)
                    .withValue(1234567.23)
                    .useDecimalFormat()
                    .prepareStyle()
                        .defaultEdging()
                        .alignment(Alignment.RIGHT)
                    .createStyle()
                .createCell()
                .prepareNewCell(3)
                    .withValue(786541.55)
                    .prepareStyle()
                        .defaultEdging()
                        .alignment(Alignment.RIGHT)
                    .createStyle()
                .createCell()
                .prepareNewCell(4)
                    .withValue(4453306.45)
                    .prepareStyle()
                        .defaultEdging()
                        .alignment(Alignment.RIGHT)
                    .createStyle()
                .createCell()
                .prepareNewCell(5)
                    .withValue(500567.75)
                    .prepareStyle()
                        .defaultEdging()
                        .alignment(Alignment.RIGHT)
                    .createStyle()
                .createCell()
                .cells()
                    .prepareStyle()
                        .defaultEdging()
                        .alignment(Alignment.CENTER)
                    .createStyle()
                .configureCells()
           .configureRow()
        .instanceFinal();

        Workbook wb = report.getWorkbook();
        FileOutputStream out = new FileOutputStream("target/bigReport_97.xls");
        wb.write(out);
        out.close();
    }

    @Test
    public void testColorReport() throws Exception {
        final String[] headerValueCells = { "Header 1", "Header 2",
                                            "Header 3", "Header 4",
                                            //
                                            "Header 2.1", "Header 2.2",
                                            "Header 2.3" };

        final int firstDataCell = 4;
        final int secondDataCell = 3;
        final int iColumns[] = { 0, 1, 2, 3, 5, 6, 7 };

        Body bd = Excel_97.instanceBody(
        new Excel_97.Header() {
            public void prepareHeader() {
                replaceColor(IndexedColors.MAROON, (byte) 203, (byte) 234, (byte) 220)
                .replaceColor(IndexedColors.PINK, (byte) 0xC5, (byte) 0xD9, (byte) 0xF1)
                .mergedRegion(0, 0, 0, firstDataCell + secondDataCell)
                .mergedRegion(1, 1, 0, firstDataCell - 1)
                .mergedRegion(1, 1, firstDataCell + 1, firstDataCell + 1 + secondDataCell - 1)
                .addNewRow()
                    .prepareNewCell(0)
                        .withValue("Title 1")
                        .prepareStyle()
                            .title()
                            .headerWithForegroundColor(IndexedColors.MAROON)
                            .addFont()
                                .heightInPoints((short) 18)
                                .color(IndexedColors.DARK_RED)
                            .configureFont()
                        .createStyle()
                    .createCell()
                    .whereHeightInPointsIs(65)
                    .cells()
                        .prepareStyle()
                            .title()
                            .header()
                        .createStyle()
                    .configureCells()
                .configureRow()
                .addNewRow()
                    .prepareNewCell(0)
                        .withValue("Title 2")
                    .createCell()
                    .prepareNewCell(firstDataCell + 1)
                        .withValue("Title 3")
                    .createCell()
                    .whereHeightInPointsIs(30)
                    .cells()
                        .prepareStyle()
                            .title()
                            .headerWithForegroundColor(IndexedColors.PINK)
                        .createStyle()
                    .configureCells()
                .configureRow()
                .iColumnWidthIs(0, 28)
                .iColumnWidthIs(1, 14)
                .iColumnWidthIs(2, 18)
                .iColumnWidthIs(3, 20)
                .iColumnWidthIs(5, 25)
                .iColumnWidthIs(6, 25)
                .iColumnWidthIs(7, 25)
                //
                .addNewRow()
                    .addAndConfigureCells(new INewCell<Cell97<Header>>() {
                        private int i = 0;

                        public void iCell(final Cell97<Header> newCell) {
                            newCell.withValue(headerValueCells[i++]);
                        }
                    }, iColumns)
                    .cells()
                        .prepareStyle()
                            .header()
                        .createStyle()
                    .configureCells()
                    .whereHeightInPointsIs(45)
                .configureRow()
                //
                .addNewRow()
                    .addCells(iColumns)
                        .withInitialValueOfIncrement(1)
                        .prepareStyle()
                            .header()
                        .createStyle()
                    .configureCells()
                .configureRow()
                //
                ;
            }
        });

        for (int i = 0; i < 100; i++)
            bd
            .addNewRow()
                .prepareNewCell(0)
                    .withValue("A" + (5+i))
                    .prepareStyle()
                        .alignment(Alignment.CENTER)
                        .defaultEdging()
                        .fillForegroundColor( i % 2 == 0 ? IndexedColors.GREY_25_PERCENT : IndexedColors.GREY_40_PERCENT)
                        .fillPattern(FillPattern.SOLID_FOREGROUND)
                    .createStyle()
                .createCell()
                .prepareNewCell(1)
                    .withValue("B" + (5+i))
                    .prepareStyle()
                        .defaultEdging()
                        .fillForegroundColor( i % 2 != 0 ? IndexedColors.GREY_25_PERCENT : IndexedColors.GREY_40_PERCENT)
                        .fillPattern(FillPattern.SOLID_FOREGROUND)
                    .createStyle()
                .createCell()
                .prepareNewCell(2)
                    .withValue("C" + (5+i))
                    .prepareStyle()
                        .defaultEdging()
                        .fillForegroundColor( i % 2 == 0 ? IndexedColors.GREY_25_PERCENT : IndexedColors.GREY_40_PERCENT)
                        .fillPattern(FillPattern.SOLID_FOREGROUND)
                    .createStyle()
                .createCell()
                .prepareNewCell(3)
                    .withValue(i)
                    .useDecimalFormat()
                .createCell()
                //
                .prepareNewCell(5)
                    .withValue("F" + (5+i))
                .createCell()
                .prepareNewCell(6)
                    .withValue(i)
                    .useDecimalFormat()
                .createCell()
                .prepareNewCell(7)
                    .withValue("H" + (5+i))
                .createCell()
                .cells()
                    .prepareStyle()
                        .defaultEdging()
                        .alignment(Alignment.RIGHT)
                    .createStyle()
                .configureCells()
            .configureRow();

        final Final report = bd.instanceFinal();
        final Workbook wb = report.getWorkbook();
        final FileOutputStream fout = new FileOutputStream("target/colorReport_97.xls");
        wb.write(fout);
        fout.close();
    }
}
