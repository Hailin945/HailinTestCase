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
 * POI HSSF 单元格背景色样式
 * @author Hailin
 */
public class ExcelHSSFCellBackgroundStyleTest2 {
    static HSSFWorkbook workbook;

    public static void main(String[] args) {
        workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet();

        HSSFRow row[] = new HSSFRow[5];
        for (int i = 0; i < 5; i++) {
            row[i] = sheet.createRow(i);
        }

        HSSFCell cell[][] = new HSSFCell[5][4];
        for (int i = 0; i < 5; i++) {
            for (int j = 0; j < 4; j++) {
                cell[i][j] = row[i].createCell((short) j);
            }
        }

        setStyle(cell[0][0], "NO_FILL", HSSFCellStyle.NO_FILL);
        setStyle(cell[0][1], "SOLID_FOREGROUND", HSSFCellStyle.SOLID_FOREGROUND);
        setStyle(cell[0][2], "FINE_DOTS", HSSFCellStyle.FINE_DOTS);
        setStyle(cell[0][3], "ALT_BARS", HSSFCellStyle.ALT_BARS);

        setStyle(cell[1][0], "SPARSE_DOTS", HSSFCellStyle.SPARSE_DOTS);
        setStyle(cell[1][1], "THICK_HORZ_BANDS", HSSFCellStyle.THICK_HORZ_BANDS);
        setStyle(cell[1][2], "THICK_VERT_BANDS", HSSFCellStyle.THICK_VERT_BANDS);
        setStyle(cell[1][3], "THICK_BACKWARD_DIAG", HSSFCellStyle.THICK_BACKWARD_DIAG);

        setStyle(cell[2][0], "THICK_FORWARD_DIAG", HSSFCellStyle.THICK_FORWARD_DIAG);
        setStyle(cell[2][1], "BIG_SPOTS", HSSFCellStyle.BIG_SPOTS);
        setStyle(cell[2][2], "BRICKS", HSSFCellStyle.BRICKS);
        setStyle(cell[2][3], "THIN_HORZ_BANDS", HSSFCellStyle.THIN_HORZ_BANDS);

        setStyle(cell[3][0], "THIN_VERT_BANDS", HSSFCellStyle.THIN_VERT_BANDS);
        setStyle(cell[3][1], "THIN_BACKWARD_DIAG", HSSFCellStyle.THIN_BACKWARD_DIAG);
        setStyle(cell[3][2], "THIN_FORWARD_DIAG", HSSFCellStyle.THIN_FORWARD_DIAG);
        setStyle(cell[3][3], "SQUARES", HSSFCellStyle.SQUARES);

        setStyle(cell[4][0], "DIAMONDS", HSSFCellStyle.DIAMONDS);

        FileOutputStream out = null;
        try {
            out = new FileOutputStream("F:/Temp/ExcelHSSFCellBackgroundStyle2.xls");
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

    public static void setStyle(HSSFCell cell, String fps, short fp) {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
        style.setFillPattern(fp);
        cell.setCellStyle(style);

        cell.setCellValue(fps);
    }
}