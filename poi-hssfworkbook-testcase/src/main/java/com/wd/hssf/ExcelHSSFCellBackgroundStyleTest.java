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
public class ExcelHSSFCellBackgroundStyleTest {
    public static void main(String[] args) {
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(1);//第二行

        HSSFCell cell = row.createCell(0);//2,1格
        cell.setCellValue("sample");//写入sample

        HSSFCellStyle style = workbook.createCellStyle();//创建个workbook的HSSFCellStyle格式对象style

        //设定格式
        style.setFillBackgroundColor(HSSFColor.WHITE.index);
        style.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
        style.setFillPattern(HSSFCellStyle.THICK_HORZ_BANDS);

        cell.setCellStyle(style);//对2,1格写入上面的格式

        FileOutputStream out = null;
        try {
            out = new FileOutputStream("F:/Temp/ExcelHSSFCellBackgroundStyle.xls");
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
}