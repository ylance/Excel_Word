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
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
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
	
	public boolean isNumeric(String s) {  
        if (s != null && !"".equals(s.trim()))  
            return s.matches("^[0-9]*$");  
        else  
            return false;  
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
		String[] paths = this.excelPath.split("\\\\");
		strArray[0] = paths[paths.length-1];
		String[] paths1 = this.excelPath2.split("\\\\");
		strArray[1] = paths1[paths1.length-1];
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
			HSSFDataFormatter df = new HSSFDataFormatter();
			String cellf = df.formatCellValue(cell2);
			results = cellf;
		}
		//TL TLD TLFD    SXZH017TL   前面有7位   TL- conditions
		//TL-1、TL-2、TL-3  TLD-1 TLD-2 TLD-3
		if(results.contains("TLFD")){
			int end = results.indexOf("TLFD");
			end = end+4;
			int begin = end-12;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TL-")){
			int end = results.indexOf("TL-");
			int count = 11;
			boolean flag = true;
			end = end+3;
			//System.out.println(results.substring(end-1, end));
			for(int k=1; k<5; k++){
				end = end+1;
				//System.out.println(results.substring(end-1, end));
				flag = isNumeric(results.substring(end-1, end));
				if(flag){
					count++;
				}else{
					break;
				}
			}
			int begin = end-count;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TLD-")){
			int end = results.indexOf("TLD-");
			int count = 12;
			boolean flag = true;
			end = end+4;
			//System.out.println(results.substring(end-1, end));
			for(int k=1; k<5; k++){
				end = end+1;
				//System.out.println(results.substring(end-1, end));
				flag = isNumeric(results.substring(end-1, end));
				if(flag){
					count++;
				}else{
				    break;
				}
			}
			int begin = end-count;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TLD")){
			int end = results.indexOf("TLD");
			end = end+3;
			int begin = end-11;
			result2 = results.substring(begin+1, end);
		}else if(results.contains("TL")){
//			System.out.println(results);
			int end = results.indexOf("TL");
//			System.out.println(end);
			end = end+2;
//			System.out.println(results.substring(end-1, end));
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
			String value = null;
			if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
				HSSFDataFormatter df = new HSSFDataFormatter();
				String cellf = df.formatCellValue(cell);
				value = cellf;
			}else if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
				value = cell.getStringCellValue();
			}else if(cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
				value = "0";
			}
			values.add(value);
		}
		//System.out.println(values);
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
			System.out.println(zh);
			List<String> datas = ub.readExcel(zh);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

