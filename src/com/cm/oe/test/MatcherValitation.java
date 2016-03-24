package com.cm.oe.test;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.jar.Attributes.Name;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.cm.oe.ui.ExcelFileFilter;

public class MatcherValitation {
	private static final Exception NullPointerException = null;

	public static void main(String[] args) throws Exception {
		String path = "E:/Desktop/拉远站-室外站-E频段-大唐-422553.xls";
		MatcherValitation mv = new MatcherValitation();
		mv.valitationFileAndFiles(path);
	}

	public void valitationFileAndFiles(String path) throws Exception {
		if (path.endsWith(".xls")) {
			if (valitation(path)) {
				JOptionPane.showMessageDialog(null, "检验通过！");
			} else {
				throw NullPointerException;
			}
		} else {
			files(path);
		}
	}

	public void files(String path) throws Exception {
		ExcelFileFilter filter = new ExcelFileFilter();
		List<File> xlsFiles = new ArrayList<File>();
		File f = new File(path);
		File[] files = f.listFiles();
		for (int j = 0; j < files.length; j++) {
			if (filter.accept(files[j])) {
				xlsFiles.add(files[j]);
			}
		}
		String filePath = null;
		int k = 0;
		for (int i = 0; i < xlsFiles.size(); i++) {
			filePath = xlsFiles.get(i).getPath();
			String name = xlsFiles.get(i).getName();

			if (valitation(filePath)) {
				// JOptionPane.showMessageDialog(null, "文件"+name+"检验通过！");
				k++;
			} else {
				// JOptionPane.showMessageDialog(null, "文件"+name+"出现错误！");
			}
		}
		if (k == xlsFiles.size()) {
			JOptionPane.showMessageDialog(null, "文件全部检验通过！");
		}
	}

	public boolean valitation(String path) throws Exception {
		File file = new File(path);
		String name = file.getName();
		String jizName = name.split("-")[0];
		FileInputStream fis = new FileInputStream(file);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sht = wb.getSheetAt(0);
		int rownums = 0;
		for (int i = 0; i < sht.getPhysicalNumberOfRows(); i++) {
			if (sht.getRow(i).getCell(0).getCellType() != 3 || !sht.getRow(i).getCell(0).toString().equals("")) {
				rownums++;
			}
		}
		boolean flag = true;
		if (rownums == 1) {
			JOptionPane.showMessageDialog(null, "文件" + name + "为空，请填写数值！");
			flag = false;
		}
		Row row = null;
		for (int i = 1; i < rownums; i++) {
			row = sht.getRow(i);
			if (sht.getRow(0).getCell(26).toString().trim().equals("天线方位角")) {
				if (MatcherXYZ(row.getCell(26).toString().trim())) {
					if (MatcherXYZ(row.getCell(27).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天下挂高(格式如：x/y/z)！");
					}
					if (MatcherXYZ(row.getCell(28).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的总下倾角(格式如：x/y/z)！");
					}
					if (row.getCell(30).toString().trim().equals("S111")
							|| row.getCell(30).toString().trim().equals("s111")) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的配置(S111或者s111)！");
					}
					if (row.getCell(31).toString().trim().equals("3") || row.getCell(31).toString().trim().equals("3.0")
							|| Double.parseDouble(row.getCell(31).toString()) == 3.0) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确RRU数量！");
					}
				} else if (MatcherXY(row.getCell(26).toString().trim())) {
					if (MatcherXY(row.getCell(27).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天下挂高(格式如：x/y)！");
					}
					if (MatcherXY(row.getCell(28).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的总下倾角(格式如：x/y)！");
					}
					if (row.getCell(30).toString().trim().equals("S11")
							|| row.getCell(30).toString().trim().equals("s11")) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的配置(S11或者s11)！");
					}
					if (row.getCell(31).toString().trim().equals("2") || row.getCell(31).toString().trim().equals("2.0")
							|| Integer.parseInt(row.getCell(31).toString()) == 2
							|| Double.parseDouble(row.getCell(31).toString()) == 2.0) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确RRU数量！");
					}
				} else if (MatcherX(row.getCell(26).toString().trim())) {
					if (MatcherX(row.getCell(27).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天下挂高(格式如：x)！");
					}
					if (MatcherX(row.getCell(28).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的总下倾角(格式如：x)！");
					}
					if (row.getCell(30).toString().trim().equals("S1")
							|| row.getCell(30).toString().trim().equals("s1")) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的配置(S1或者s1)！");
					}
					if (row.getCell(31).toString().trim().equals("1") || row.getCell(31).toString().trim().equals("1.0")
							|| Double.parseDouble(row.getCell(31).toString()) == 1.0) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确RRU数量！");
					}
				} else {
					flag = false;
					JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天线方位角(格式如：x)！");
				}
			} else if (sht.getRow(0).getCell(24).toString().trim().equals("天线方位角")) {
				if (MatcherXYZ(row.getCell(24).toString().trim())) {
					if (MatcherXYZ(row.getCell(25).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天下挂高(格式如：x/y/z)！");
					}
					if (MatcherXYZ(row.getCell(26).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的总下倾角(格式如：x/y/z)！");
					}
					if (row.getCell(28).toString().trim().equals("S111")
							|| row.getCell(28).toString().trim().equals("s111")) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的配置(S111或者s111)！");
					}
					if (row.getCell(29).toString().trim().equals("3") || row.getCell(29).toString().trim().equals("3.0")
							|| Double.parseDouble(row.getCell(29).toString()) == 3.0) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确RRU数量！");
					}
				} else if (MatcherXY(row.getCell(24).toString().trim())) {
					if (MatcherXY(row.getCell(25).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天下挂高(格式如：x/y)！");
					}
					if (MatcherXY(row.getCell(26).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的总下倾角(格式如：x/y)！");
					}
					if (row.getCell(28).toString().trim().equals("S11")
							|| row.getCell(28).toString().trim().equals("s11")) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的配置(S11或者s11)！");
					}
					if (row.getCell(29).toString().trim().equals("2") || row.getCell(29).toString().trim().equals("2.0")
							|| Integer.parseInt(row.getCell(29).toString()) == 2
							|| Double.parseDouble(row.getCell(29).toString()) == 2.0) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确RRU数量！");
					}
				} else if (MatcherX(row.getCell(24).toString().trim())) {
					if (MatcherX(row.getCell(25).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天下挂高(格式如：x)！");
					}
					if (MatcherX(row.getCell(26).toString().trim())) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的总下倾角(格式如：x)！");
					}
					if (row.getCell(28).toString().trim().equals("S1")
							|| row.getCell(28).toString().trim().equals("s1")) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的配置(S1或者s1)！");
					}
					if (row.getCell(29).toString().trim().equals("1") || row.getCell(29).toString().trim().equals("1.0")
							|| Double.parseDouble(row.getCell(29).toString()) == 1.0) {

					} else {
						flag = false;
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确RRU数量！");
					}
				} else {
					flag = false;
					JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行，请输入正确的天线方位角(格式如：x)！");
				}
			} else if (sht.getRow(0).getCell(24).toString().trim().equals("配置")) {

			} else {
				flag = false;
				JOptionPane.showConfirmDialog(null, "文件" + name + "不正确，请选择正确的拉远站宏基站小基站信源站文件！");
			}
			if (jizName.equals("宏基站") || jizName.equals("拉远站")) {
				for (int j = 0; j < 38; j++) {
					Cell cell = row.getCell(j);
					if (cell != row.getCell(26) && cell != row.getCell(27) && cell != row.getCell(28)
							&& cell != row.getCell(30) && cell != row.getCell(31)) {
						if (cell.getCellType() == 3 || row.getCell(j).toString().equals("")) {
							JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行"
									+ sht.getRow(0).getCell(j).toString().trim() + "为空，请填写数值！");
							flag = false;
						}
					}
				}
			} else if (jizName.equals("小基站")) {
				for (int j = 0; j < 36; j++) {
					Cell cell = row.getCell(j);
					if (cell != row.getCell(24) && cell != row.getCell(25) && cell != row.getCell(26)
							&& cell != row.getCell(28) && cell != row.getCell(29)) {
						if (cell.getCellType() == 3 || row.getCell(j).toString().equals("")) {
							JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行"
									+ sht.getRow(0).getCell(j).toString().trim() + "为空，请填写数值！");
							flag = false;
						}
					}
				}
			} else if (jizName.equals("信源站")) {
				for (int j = 0; j < 32; j++) {
					Cell cell = row.getCell(j);
					if (cell.getCellType() == 3 || row.getCell(j).toString().equals("")) {
						JOptionPane.showMessageDialog(null, "汇总表" + name + "第" + i + "行"
								+ sht.getRow(0).getCell(j).toString().trim() + "为空，请填写数值！");
						flag = false;
					}
				}
			}
		}
		wb.close();
		fis.close();
		return flag;

	}

	public boolean MatcherXYZ(String s) throws Exception {
		Pattern pattern = Pattern.compile("^\\d{1,3}/\\d{1,3}/\\d{1,3}$");
		Matcher matcher = pattern.matcher(s);
		if (matcher.find()) {
			return true;
		} else {
			return false;
		}
	}

	public boolean MatcherXY(String s) throws Exception {
		Pattern pattern = Pattern.compile("^\\d{1,3}/\\d{1,3}$");
		Matcher matcher = pattern.matcher(s);
		if (matcher.find()) {
			return true;
		} else {
			return false;
		}

	}

	public boolean MatcherX(String s) throws Exception {
		Pattern pattern = Pattern.compile("^\\d{1,3}$");
		Matcher matcher = pattern.matcher(s);
		if (matcher.find()) {
			return true;
		} else {
			return false;
		}
	}
}
