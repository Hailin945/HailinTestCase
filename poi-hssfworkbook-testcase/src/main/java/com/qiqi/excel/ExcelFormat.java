package com.qiqi.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelFormat {
	
	/**
	 * 点击运行
	 * @param args
	 */
	public static void main(String[] args) {

		// 读取Excel
		List<ExcelDTO> readExcel = readExcel();
		createExcel(readExcel);
		System.out.println("导出成功！");
	}
	
	
	

	private static List<ExcelDTO> readExcel() {
		List<ExcelDTO> list = new ArrayList<>();
		HSSFWorkbook workbook = null;

		try {
			// 读取Excel文件
			InputStream inputStream = new FileInputStream("F:\\Temp\\order.xls");
			workbook = new HSSFWorkbook(inputStream);
			inputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		// 循环工作表
		for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = workbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}
			// 循环行
			for (int rowNum = 3; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow == null) {
					continue;
				}

				// 将单元格中的内容存入集合
				ExcelDTO excelDTO = new ExcelDTO();

				HSSFCell cell = hssfRow.getCell(2);
				if (cell == null) {
					continue;
				}
				excelDTO.setCompanyName(cell.getStringCellValue());

				cell = hssfRow.getCell(17);
				if (cell == null) {
					continue;
				}
				excelDTO.setOrderType(cell.getStringCellValue());


				list.add(excelDTO);
			}
		}
		
		Map<String, List<ExcelDTO>> map = new HashMap<>();
		for (ExcelDTO excelDTO : list) {
			String companyName = excelDTO.getCompanyName();
			if (!map.containsKey(companyName)) {
				map.put(companyName, new ArrayList<>());
			}
//			System.out.println(excelDTO.getCompanyName()+"----"+excelDTO.getOrderType());
		}
		
		for (ExcelDTO excelDTO : list) {
			map.get(excelDTO.getCompanyName()).add(excelDTO);
		}
		
		List<ExcelDTO> resultList = new ArrayList<>();
		
		Set<String> keySet = map.keySet();
		for (String name : keySet) {
			List<ExcelDTO> excelDTOs = map.get(name);
			int cancelTotalNum = 0;
			int completedTotalNum = 0;
			int totalNum = excelDTOs.size();
			for (ExcelDTO excelDTO : excelDTOs) {
				if ("已取消".equals(excelDTO.getOrderType())) {
					cancelTotalNum ++;
				}
				if ("已完成".equals(excelDTO.getOrderType())) {
					completedTotalNum ++;
				}
			}
			ExcelDTO result = new ExcelDTO();
			result.setCompanyName(name);
			result.setCancenlTotalNum(cancelTotalNum);
			result.setCompletedTotalNum(completedTotalNum);
			result.setTotalNum(totalNum); 
			resultList.add(result);
		}
		
		ExcelDTO neiMeng = new ExcelDTO("内蒙地区业务", 0, 0, 0);
		ExcelDTO daTong = new ExcelDTO("大同地区业务", 0, 0, 0);
		ExcelDTO lingShi = new ExcelDTO("灵石地区业务", 0, 0, 0);
		ExcelDTO ningXia = new ExcelDTO("宁夏地区业务", 0, 0, 0);
		ExcelDTO siChuan = new ExcelDTO("四川地区业务", 0, 0, 0);
		
		Iterator<ExcelDTO> iterator = resultList.iterator();
		while (iterator.hasNext()) {
			ExcelDTO excel = iterator.next();
			if (excel.getCompanyName().contains("内蒙") || excel.getCompanyName().contains("包头") || excel.getCompanyName().contains("鄂尔多斯") || excel.getCompanyName().contains("土右旗")) {
				neiMeng.setCancenlTotalNum(neiMeng.getCancenlTotalNum() + excel.getCancenlTotalNum());
				neiMeng.setCompletedTotalNum(neiMeng.getCompletedTotalNum() + excel.getCompletedTotalNum());
				neiMeng.setTotalNum(neiMeng.getTotalNum() + excel.getTotalNum());
				iterator.remove();
			} else if (excel.getCompanyName().contains("大同")) {
				daTong.setCancenlTotalNum(neiMeng.getCancenlTotalNum() + excel.getCancenlTotalNum());
				daTong.setCompletedTotalNum(neiMeng.getCompletedTotalNum() + excel.getCompletedTotalNum());
				daTong.setTotalNum(neiMeng.getTotalNum() + excel.getTotalNum());
				iterator.remove();
			} else if (excel.getCompanyName().contains("灵石")) {
				lingShi.setCancenlTotalNum(neiMeng.getCancenlTotalNum() + excel.getCancenlTotalNum());
				lingShi.setCompletedTotalNum(neiMeng.getCompletedTotalNum() + excel.getCompletedTotalNum());
				lingShi.setTotalNum(neiMeng.getTotalNum() + excel.getTotalNum());
				iterator.remove();
			} else if (excel.getCompanyName().contains("宁夏")) {
				ningXia.setCancenlTotalNum(neiMeng.getCancenlTotalNum() + excel.getCancenlTotalNum());
				ningXia.setCompletedTotalNum(neiMeng.getCompletedTotalNum() + excel.getCompletedTotalNum());
				ningXia.setTotalNum(neiMeng.getTotalNum() + excel.getTotalNum());
				iterator.remove();
			} else if (excel.getCompanyName().contains("四川") || excel.getCompanyName().contains("贵州") || excel.getCompanyName().contains("成都") || excel.getCompanyName().contains("峨眉山")) {
				siChuan.setCancenlTotalNum(neiMeng.getCancenlTotalNum() + excel.getCancenlTotalNum());
				siChuan.setCompletedTotalNum(neiMeng.getCompletedTotalNum() + excel.getCompletedTotalNum());
				siChuan.setTotalNum(neiMeng.getTotalNum() + excel.getTotalNum());
				iterator.remove();
			}
		}
		
		ExcelDTO qiTa = new ExcelDTO("其他", 0, 0, 0);
		Iterator<ExcelDTO> it = resultList.iterator();
		while (it.hasNext()) {
			ExcelDTO excel = it.next();
			if (excel.getTotalNum() < 30) {
				qiTa.setCancenlTotalNum(neiMeng.getCancenlTotalNum() + excel.getCancenlTotalNum());
				qiTa.setCompletedTotalNum(neiMeng.getCompletedTotalNum() + excel.getCompletedTotalNum());
				qiTa.setTotalNum(neiMeng.getTotalNum() + excel.getTotalNum());
				it.remove();
			}
		}
		
		resultList.add(neiMeng);
		resultList.add(daTong);
		resultList.add(lingShi);
		resultList.add(ningXia);
		resultList.add(siChuan);
		resultList.add(qiTa);
		
		
		return resultList;
	}
	
	 /**
     * 创建Excel
     */
    private static void createExcel(List<ExcelDTO> list) {
    	
        // 创建一个Excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建一个工作表
        HSSFSheet sheet = workbook.createSheet("order");
        // 添加表头行
        HSSFRow hssfRow = sheet.createRow(0);
        // 设置单元格格式居中
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        // 添加表头内容
        HSSFCell headCell = hssfRow.createCell(0);
        headCell.setCellValue("公司名称");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(1);
        headCell.setCellValue("总订单数");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(2);
        headCell.setCellValue("已取消");
        headCell.setCellStyle(cellStyle);
        
        headCell = hssfRow.createCell(3);
        headCell.setCellValue("进行中");
        headCell.setCellStyle(cellStyle);
        
        headCell = hssfRow.createCell(4);
        headCell.setCellValue("已完成");
        headCell.setCellStyle(cellStyle);

        // 添加数据内容
        for (int i = 0; i < list.size(); i++) {
            hssfRow = sheet.createRow((int) i + 1);
            ExcelDTO excelDTO = list.get(i);

            // 创建单元格，并设置值
            HSSFCell cell = hssfRow.createCell(0);
            cell.setCellValue(excelDTO.getCompanyName());
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(1);
            cell.setCellValue(excelDTO.getTotalNum());
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(2);
            cell.setCellValue(excelDTO.getCancenlTotalNum());
            cell.setCellStyle(cellStyle);
            
            cell = hssfRow.createCell(3);
            cell.setCellValue(excelDTO.getTotalNum() - excelDTO.getCancenlTotalNum() - excelDTO.getCompletedTotalNum());
            cell.setCellStyle(cellStyle);
            
            cell = hssfRow.createCell(4);
            cell.setCellValue(excelDTO.getCompletedTotalNum());
            cell.setCellStyle(cellStyle);
        }

        // 保存Excel文件
        try {
            OutputStream outputStream = new FileOutputStream("F:\\Temp\\abc.xls");
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
   
	
}
