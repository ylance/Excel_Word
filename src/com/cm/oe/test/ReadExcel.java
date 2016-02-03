package com.cm.oe.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReadExcel {
	/*public static void main(String[] args) throws IOException {
		ReadExcel readExcel  = new ReadExcel();
		System.out.println(readExcel.getZH("C:\\Users\\王宁\\Desktop\\新建文件夹\\宏基站-室外站-D频段-上海贝尔-24111.xls"));
	}*/
	public List<String> getZH(String excelPath) throws IOException {
		FileInputStream fis = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		Sheet sht = wb.getSheetAt(0);
		List<String> zhs = new ArrayList<String>();
		for (int i = 0; i < rowNumber(wb); i++) {
			Row row = sht.getRow(i);
			String zh = row.getCell(2).toString();
			zhs.add(zh);
		}
		wb.close();
		fis.close();
		return zhs;
	}

	public int rowNumber(HSSFWorkbook wb) {
		Sheet sht = wb.getSheetAt(0);
		int rowNums = 0;
		int rowCount = sht.getPhysicalNumberOfRows();
		for (int i = 0; i < rowCount; i++) {
			Row row = sht.getRow(i);
			//System.out.println(row.getCell(0));
			if (row.getCell(0) == null) {
				return rowNums;
			} else if (row.getCell(0).getCellType() != 3) {
				rowNums++;
			}
		}
		return rowNums;
	}

	public List<String> getExcelvalues(Row row) {
		List<String> line = new ArrayList<String>();
		for (int i = row.getFirstCellNum(); i < row.getPhysicalNumberOfCells(); i++) {
			Cell cell = row.getCell(i);
			switch (cell.getCellType()) {
			case HSSFCell.CELL_TYPE_STRING:
				line.add(cell.getStringCellValue());
				break;
			case HSSFCell.CELL_TYPE_NUMERIC:
				HSSFDataFormatter df = new HSSFDataFormatter();
				String cellf = df.formatCellValue(cell);
				line.add(cellf);
			default:
				break;
			}
		}
		return line;
	}

}
