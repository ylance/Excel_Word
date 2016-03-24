package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReadExcel {
	
	private static final Exception NullPointerException = null;
	private String excelPath ;
	
	public ReadExcel(String excelPath) {
		this.excelPath = excelPath;
	}
/*	public static void main(String[] args) throws IOException {
		String excelPath="E:\\Desktop\\宏基站-室外站-D频段-上海贝尔-24111.xls";
		ReadExcel readExcel  = new ReadExcel(excelPath);
		System.out.println(readExcel.getRow("SXZH001TL"));
	}*/
	
	public Integer getRow(String zh) throws Exception{
		List<String> zhs = getZH();
		int i = 1;
		if(zhs.contains(zh)){
			i=zhs.indexOf(zh);
		}
		return i;
	}
	public List<String> getZH() throws Exception {
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

	public int rowNumber(HSSFWorkbook wb) throws Exception {
		Sheet sht = wb.getSheetAt(0);
		int rowNums = 0;
		int rowCount = sht.getPhysicalNumberOfRows();
		for (int i = 0; i < rowCount; i++) {
			Row row = sht.getRow(i);
			//System.out.println(row.getCell(0));
			if (row.getCell(0) == null||row.getCell(0).toString().trim().equals("")) {
				return rowNums;
			} else if (row.getCell(0).getCellType() != 3 && !row.getCell(0).toString().trim().equals("")) {
				rowNums++;
			}
		}
		if(rowNums ==1){
			JOptionPane.showMessageDialog(null, "所选汇总表没有填写内容！");
			throw NullPointerException;
		}
		return rowNums;
	}

	public List<String> getExcelvalues(Row row) throws Exception {
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
			if(row.getCell(i).toString().trim().equals(" ")||row.getCell(i).getCellType() ==3||row.getCell(i)==null){
				JOptionPane.showMessageDialog(null, "汇总表第"+row.getRowNum()+"行"+cell.getColumnIndex()+"列为空，请填写数值！");
				throw NullPointerException;
			}
		}
		return line;
	}

}
