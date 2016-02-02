package com.cm.oe.ui;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.cm.oe.budget.gen.BudgetReader2;

public class ReadB3 {

	public Map<String, String> read(String path1, String path2) throws FileNotFoundException, IOException {
		String a = null;
		ExcelFileFilter filter = new ExcelFileFilter();
		List<File> xlsFiles = new ArrayList<File>();
		Map<String, String> map = new HashMap<String, String>();
		// 确定路径下 文件的个数
		File f = new File(path1);
		// System.out.println(path1);

		File[] files = f.listFiles();
		for (int j = 0; j < files.length; j++) {
			if (filter.accept(files[j])) {
				xlsFiles.add(files[j]);
			}
		}
		for (int i = 0; i < xlsFiles.size(); i++) {
			String zh = null;
			String path = xlsFiles.get(i).getPath();
			// System.out.println(path);
			BudgetReader2 br = new BudgetReader2(path, path2);
			zh = br.getZhFrom4GYsb();
			map.put(zh, path);
		}
		return map;
	}

	/*
	 * public Map<String, String> read1(String path1,String path2 ) throws
	 * FileNotFoundException, IOException{ int i ; String a = null; Map<String,
	 * String> map =new HashMap<String,String>(); //确定路径下 文件的个数 File f= new
	 * File(path1); File[] files= f.listFiles(); for(i=0;i<files.length;i++){
	 * String zh =null; String path= files[i].getPath(); HSSFWorkbook wb = new
	 * HSSFWorkbook(new FileInputStream(path));
	 * zh=wb.getSheet("B3甲").getRow(2).getCell(1).getStringCellValue();
	 * if(zh.contains("TLFD")){ int end = zh.indexOf("TLFD"); end = end+4; int
	 * begin = end-10; zh = zh.substring(begin+1, end); }else
	 * if(zh.contains("TLD")){ int end = zh.indexOf("TLD"); end = end+3; int
	 * begin = end-10; zh = zh.substring(begin+1, end); }else
	 * if(zh.contains("TL")){ int end = zh.indexOf("TL"); end = end+2; int begin
	 * = end-10; zh = zh.substring(begin+1, end); } System.out.println(zh);
	 * System.out.println(path); map.put(zh,path); } return map; }
	 */
}