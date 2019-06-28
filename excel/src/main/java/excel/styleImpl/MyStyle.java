package com.mbc.common.utils.excel.styleImpl;

import com.mbc.common.utils.excel.base.ExcelStyleBase;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by 符柱成 on 2017/8/25.
 */
public class MyStyle extends ExcelStyleBase {

    @Override
    public CellStyle setHeaderStyle(Workbook wb) {

        // 设置字体
        Font font = wb.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 10);
        //字体加粗
        font.setBold(true);
        font.setFontName("宋体");
        font.setColor((short) 10);
        // 生成一个样式
        CellStyle headerStyle = wb.createCellStyle();
        // 设置背景色
        headerStyle.setFillForegroundColor((short) 13);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置字体
        headerStyle.setFont(font);
        //设置顶边框;
        headerStyle.setBorderTop(BorderStyle.THIN);
        //设置右边框;
        headerStyle.setBorderRight(BorderStyle.THIN);
        //设置左边框;
        headerStyle.setBorderLeft(BorderStyle.THIN);
        //设置底边框;
        headerStyle.setBorderBottom(BorderStyle.THIN);
        //设置自动换行;
        headerStyle.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置垂直对齐的样式为居中对齐;
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return headerStyle;
    }

    @Override
    public CellStyle setDataStyle(Workbook wb) {
        // 设置字体
        Font font = wb.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 10);
        //字体加粗
//        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");
        font.setColor((short) 32767);
        // 生成一个样式
        CellStyle dataStyle = wb.createCellStyle();
        // 设置背景色
        dataStyle.setFillForegroundColor((short) 70);
        dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置字体
        dataStyle.setFont(font);
        //设置顶边框;
        dataStyle.setBorderTop(BorderStyle.THIN);
        //设置右边框;
        dataStyle.setBorderRight(BorderStyle.THIN);
        //设置左边框;
        dataStyle.setBorderLeft(BorderStyle.THIN);
        //设置底边框;
        dataStyle.setBorderBottom(BorderStyle.THIN);
        //设置自动换行;
        dataStyle.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置垂直对齐的样式为居中对齐;
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return dataStyle;
    }

    @Override
    public void setRowHigh() {
        setMyRowHigh((short) (2 * 256));
    }

    @Override
    public void setColumnWidth() {
        setMyColumnWidth((short) 30);
    }

    @Override
    public Map<Integer, Integer> setMySpecifiedHighAndWidth() {
        Map<Integer, Integer> map = new HashMap();
        map.put(1, 3000);
        map.put(2, 3000);
        return map;
    }

    @Override
    public CellStyle setHeaderStyle(Workbook wb, Sheet sheet) {
        return null;
    }
}
