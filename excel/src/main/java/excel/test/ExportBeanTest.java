package com.mbc.common.utils.excel.test;


import com.mbc.common.utils.excel.base.PoiExcelBase;
import com.mbc.common.utils.excel.base.PoiInterface;
import com.mbc.common.utils.excel.base.StyleInterface;
import com.mbc.common.utils.excel.model.StudentModelExcelTest;
import com.mbc.common.utils.excel.styleImpl.TestStyle;
import com.mbc.common.utils.excel.util.PoiBeanFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by 符柱成 on 2017/8/24.
 */
public class ExportBeanTest {
    public static void main(String[] args) throws IOException {
        List<String> listName = new ArrayList();
        listName.add("id");
        listName.add("名字");
        listName.add("性别");
        List<String> listId = new ArrayList();
        listId.add("id");
        listId.add("sex");
        listId.add("name");
        List<StudentModelExcelTest> list = new ArrayList();
        list.add(new StudentModelExcelTest(111, "张三", "男"));
        list.add(new StudentModelExcelTest(111, "李四", "男"));
        list.add(new StudentModelExcelTest(111, "王五", "女"));

        FileOutputStream exportXls = null;
        if (PoiExcelBase.EXCEL_VERSION_07 == 0) {
            exportXls = new FileOutputStream("/Users/yjj/Desktop/abc.xls");
        } else {
            exportXls = new FileOutputStream("/Users/yjj/Desktop/abc.xls");
        }
        /*
            （一）去工厂拿导出工具
         */
        PoiInterface<StudentModelExcelTest> poiInterface = PoiBeanFactory.getInstance().getPoiUtil(PoiExcelBase.EXPORT_SIMPLE_EXCEL);
        /*
            （二）自定义样式（可无）
         */
        StyleInterface myStyle = new TestStyle();
        /*
            （三）根据需求选择接口方法（返回码：1是成功，0为失败）
         */
        //导出默认样式EXCEL文件（根据headersId来导出对应字段，）--根据headersId筛选要导出的字段
        //int flag = poiInterface.exportBeanExcel(PoiExcelBase.EXCEL_VERSION_07,"测试POI导出EXCEL文档",listName,listId,list,exportXls);

        //导出自定义样式Excel文件--根据headersId筛选要导出的字段
        int flag = poiInterface.exportStyleBeanExcel(PoiExcelBase.EXCEL_VERSION_07, "测试POI导出EXCEL文档", listName, listId, list, exportXls, myStyle);
        //默认导出dtolist的所有数据--默认导出dtolist的所有数据
        // int flag = poiInterface.exportStyleBeanExcel(PoiExcelBase.EXCEL_VERSION_07,"测试POI导出EXCEL文档",listName,list,exportXls,myStyle);
        System.out.println("flag  : " + flag);
        exportXls.close();
    }
}
