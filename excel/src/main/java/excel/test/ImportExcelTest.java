package com.mbc.common.utils.excel.test;

import com.mbc.common.utils.excel.model.PhoneModelExcelTest;
import com.mbc.common.utils.excel.util.ImportExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class ImportExcelTest {

    public static void main(String[] args) throws Exception {
        //导入
        String keyValue = "手机名称:phoneName,颜色:color,售价:price";
        List<PhoneModelExcelTest> list = ImportExcel.readExcel(
                "test.xls",
                new FileInputStream("/Users/yjj/Desktop/test.xls"),
                ImportExcel.getMap(keyValue),
                "com.mbc.common.utils.excel.model.PhoneModelExcelTest", 1);
        System.out.println(list.size());
        //导出
        ImportExcel.exportExcel(
                "导出数据",
                new FileOutputStream("/Users/yjj/Desktop/testOut.xls"),
                ImportExcel.getMap(keyValue), list,
                "com.mbc.common.utils.excel.model.PhoneModelExcelTest", null, null, null);
    }


}
