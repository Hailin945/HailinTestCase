package com.wd.hssf;


import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * POI HSSF 边框样式
 * @author Hailin
 */
public class ExcelHSSFCellBorderStyleTest {
    static HSSFWorkbook workbook;

    public static void main(String[] args) {
        workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet();

        HSSFRow row[] = new HSSFRow[5];
        for (int i = 0; i < 5; i++) {
            row[i] = sheet.createRow(i);
        }

        HSSFCell cell[][] = new HSSFCell[5][3];
        for (int i = 0; i < 5; i++) {
            for (int j = 0; j < 3; j++) {
                cell[i][j] = row[i].createCell((short) j);
            }
        }

        setStyle(cell[0][0], "DASH_DOT", HSSFCellStyle.BORDER_DASH_DOT);
        setStyle(cell[0][1], "DASH_DOT_DOT", HSSFCellStyle.BORDER_DASH_DOT_DOT);
        setStyle(cell[0][2], "DASHED", HSSFCellStyle.BORDER_DASHED);

        setStyle(cell[1][0], "DOTTED", HSSFCellStyle.BORDER_DOTTED);
        setStyle(cell[1][1], "DOUBLE", HSSFCellStyle.BORDER_DOUBLE);
        setStyle(cell[1][2], "HAIR", HSSFCellStyle.BORDER_HAIR);

        setStyle(cell[2][0], "MEDIUM", HSSFCellStyle.BORDER_MEDIUM);
        setStyle(cell[2][1], "MEDIUM_DASH_DOT", HSSFCellStyle.BORDER_MEDIUM_DASH_DOT);
        setStyle(cell[2][2], "MEDIUM_DASH_DOT_DOT", HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);

        setStyle(cell[3][0], "MEDIUM_DASHED", HSSFCellStyle.BORDER_MEDIUM_DASHED);
        setStyle(cell[3][1], "NONE", HSSFCellStyle.BORDER_NONE);
        setStyle(cell[3][2], "SLANTED_DASH_DOT", HSSFCellStyle.BORDER_SLANTED_DASH_DOT);

        setStyle(cell[4][0], "THICK", HSSFCellStyle.BORDER_THICK);
        setStyle(cell[4][1], "THIN", HSSFCellStyle.BORDER_THIN);

        FileOutputStream out = null;
        try {
            out = new FileOutputStream("F:/Temp/ExcelHSSFCellBorderStyle.xls");
            workbook.write(out);
        } catch (IOException e) {
            System.out.println(e.toString());
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                System.out.println(e.toString());
            }
        }
    }

    public static void setStyle(HSSFCell cell, String bn, short border) {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setBorderBottom(border);
        style.setBottomBorderColor(HSSFColor.ORANGE.index);
        cell.setCellStyle(style);

        cell.setCellValue(bn);
    }
}