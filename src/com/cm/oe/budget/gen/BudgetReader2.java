package com.cm.oe.budget.gen;




import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;



public class BudgetReader2 {
	//TODO:添加资源关闭语句
	private String excelPath;
	private String excelPath2 ;
	
	public BudgetReader2(String excelPath, String excelPath2){
		this.excelPath = excelPath;
		this.excelPath2 = excelPath2;
	}
	
	public String getZhFrom4GYsb() throws IOException{
		//预算表中  第三行 B列的名称为： 单项工程名称:SXZH001TL新建、共址2G、共址其他运营商的(F)（D)宏站基站
		//path1: 预算表   path2 3g4g基础信息
		String results = "";
		String result2 = "";
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(excelPath2));
		
		HSSFFormulaEvaluator e= new HSSFFormulaEvaluator(wb);
		HSSFFormulaEvaluator e2= new HSSFFormulaEvaluator(wb2);
		
		//TODO: 将此处的文件名替换为从参数读取的文件名
		String [] strArray = new String[2];
		String[] splits = excelPath.split("/");
		String[] splits2 = excelPath2.split("/");
		strArray[0] = splits[splits.length-1];
		strArray[1] = splits2[splits2.length-1];
		HSSFFormulaEvaluator[] evals = new HSSFFormulaEvaluator[2];
		evals[0] = e;
		evals[1] = e2;
		HSSFFormulaEvaluator.setupEnvironment(strArray, evals); 
		Sheet sheet = wb.getSheetAt(7);
		Row r = null;
		r = sheet.getRow(2);
		Cell cell2 = r.getCell(1);
		if(cell2.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
			results = e.evaluate(cell2).getStringValue();
		}else if(cell2.getCellType()==HSSFCell.CELL_TYPE_STRING){
			results = cell2.getStringCellValue();
		}
		//TL TLD TLFD    SXZH017TL   前面有7位
		if(results.contains("TLFD")){
			int end = results.indexOf("TLFD");
			end = end+4;
			int begin = end-10;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TLD")){
			int end = results.indexOf("TLD");
			end = end+3;
			int begin = end-10;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TL")){
			int end = results.indexOf("TL");
			end = end+2;
			int begin = end-10;
			result2 = results.substring(begin+1, end);
		}
		wb2.close();
		wb.close();
		fise.close();
		return result2;
	}
	
	public List<String> readExcel(String zh) throws IOException{
		List<String> values = new ArrayList<String>();
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		
		Sheet sheet3 = wb.getSheetAt(1);
		Row r = null;	
		int linenum = 0;
		Cell cell = null;
		for(int i=8; i<sheet3.getPhysicalNumberOfRows(); i++){
			r = sheet3.getRow(i);
			cell = r.getCell(3);
			if(cell.toString().contains(zh)){
				linenum = i;
				break;
			}
		}
		r = sheet3.getRow(linenum);
		for(int i=4;i<=10;i++){
			cell = r.getCell(i);
			values.add(cell.toString());
		}
		
		wb.close();
		fise.close();
		return values;
	}

	public static void main(String[] args) {
		String path1 = "E:\\Desktop\\预算表\\4G工程基站预算输出表3.xls";
		String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
		BudgetReader2 ub = new BudgetReader2(path1, path2);
		try {
			String zh = ub.getZhFrom4GYsb();
			List<String> datas = ub.readExcel(zh);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

