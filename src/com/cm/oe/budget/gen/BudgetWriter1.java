package com.cm.oe.budget.gen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JOptionPane;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSignedTwipsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTextScale;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHint;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import com.cm.oe.test.ReadExcel;
import com.cm.oe.test.ReadExcelTable;

public class BudgetWriter1 {
	private static final Exception NullPointerException = null;
	public String path;
	public String zhao;
	public String path2;
	private String tablePath;
	private String excelPath;
	private String output;

	public BudgetWriter1(String path, String zhao, String path2, String tablePath, String excelPath, String output) {
		this.zhao = zhao;
		this.path = path;
		this.path2 = path2;
		this.tablePath = tablePath;
		this.excelPath = excelPath;
		this.output = output;

	}

	public static void main(String[] args) throws Exception {
		String path = "E:\\Desktop\\预算表";
		String path2 = "testfiles/3G4G工程基站预算基础信息表.xls";
		String tablePath = "testfiles/tables.xls";
		String excelPath = "E:\\Desktop\\拉远站-室外站-F频段-华为-341358.xls";
		String output = "E:\\Desktop\\";
		String zhao = "SXZH002TL";
		BudgetWriter1 bw = new BudgetWriter1(path, zhao, path2, tablePath, excelPath, output);
		bw.testReadByDoc();
	}

	@SuppressWarnings("resource")
	public void testReadByDoc() throws Exception {
		BudgetReader1 ub = new BudgetReader1(path, path2);
		BudgetReader2 ub2 = new BudgetReader2(path, path2);
		BudgetReader3 ub3 = new BudgetReader3(path, path2);
		Map<String, List<String>> data_b3 = ub.getB3(path);
		Map<String, Map<String, String>> data_all = ub.get3G4Gjcxx(path2);
		Map<String, List<String>> datas_map = ub.getB3JData(data_all, data_b3, path, path2);
		ReadExcelTable ret = new ReadExcelTable(tablePath, excelPath);
		ReadExcel re = new ReadExcel(excelPath);
		Map<Integer, List<String>> excelmap = new HashMap<Integer, List<String>>();
		Map<Integer, List<String>> BBUtablemap = ret.readBBUinExcel();
		Map<Integer, List<String>> RRUtablemap = ret.readRRUinExcel();
		Map<Integer, List<String>> Antennatablemap = ret.readAntennaIntables();
		Map<Integer, String> AZHoLa = ret.anzhuang();
		FileInputStream fise = new FileInputStream(excelPath);
		HSSFWorkbook wb = new HSSFWorkbook(fise);
		Sheet sheet = wb.getSheetAt(0);

		int rowNums = re.rowNumber(wb);
		// System.out.println(rowNums);
		FileOutputStream fos = null;
		Row r = null;
		for (int i = 0; i < rowNums; i++) {
			List<String> line = new ArrayList<String>();
			r = sheet.getRow(i);
			String zmzh = re.getExcelvalues(r).get(1).trim() + re.getExcelvalues(r).get(2).trim();
			String dizhi = re.getExcelvalues(r).get(0).trim();
			line.add(dizhi);
			line.add(zmzh);
			for (int j = 3; j < re.getExcelvalues(r).size(); j++) {
				line.add(re.getExcelvalues(r).get(j));
			}
			excelmap.put(i, line);
		}
		// System.out.println(excelmap.size());
		// System.out.println(excelmap);

		int i = re.getRow(zhao);
		File file = new File(excelPath);
		InputStream is = null;
		String[] strs = file.getName().split("-");
		String name = strs[0];
		String neiwai = strs[1];
		String pinduan = strs[2];
		String shebei = strs[3];
		if (name.trim().equals("宏基站")) {
			if (neiwai.trim().equals("室内站")) {
				if (pinduan.trim().equals("F频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/宏基站室内F频段无天线.docx");
					} else {

						is = new FileInputStream("testfiles/宏基站室内F频段.docx");
					}
				} else if (pinduan.trim().equals("E频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/宏基站室内E频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/宏基站室内E频段.docx");
					}
				} else if (pinduan.trim().equals("D频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/宏基站室内D频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/宏基站室内D频段.docx");
					}
				}

			} else if (neiwai.trim().equals("室外站")) {
				if (pinduan.trim().equals("F频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/宏基站室外F频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/宏基站室外F频段.docx");
					}

				} else if (pinduan.trim().equals("E频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/宏基站室外E频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/宏基站室外E频段.docx");
					}
				} else if (pinduan.trim().equals("D频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/宏基站室外D频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/宏基站室外D频段.docx");
					}
				}
			}
		} else if (name.trim().equals("小基站")) {
			if (neiwai.trim().equals("室内站")) {
				if (pinduan.trim().equals("F频段")) {
					is = new FileInputStream("testfiles/小基站室内F频段.docx");
				} else if (pinduan.trim().equals("E频段")) {
					is = new FileInputStream("testfiles/小基站室内E频段.docx");
				} else if (pinduan.trim().equals("D频段")) {
					is = new FileInputStream("testfiles/小基站室内D频段.docx");
				}
			} else if (neiwai.trim().equals("室外站")) {
				if (pinduan.trim().equals("F频段")) {
					is = new FileInputStream("testfiles/小基站室外F频段.docx");
				} else if (pinduan.trim().equals("E频段")) {
					is = new FileInputStream("testfiles/小基站室外E频段.docx");
				} else if (pinduan.trim().equals("D频段")) {
					is = new FileInputStream("testfiles/小基站室外D频段.docx");
				}
			}
		} else if (name.trim().equals("拉远站")) {
			if (neiwai.trim().equals("室内站")) {
				if (pinduan.trim().equals("F频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/拉远站室内F频段无天线.docx");
					} else {

						is = new FileInputStream("testfiles/拉远站室内F频段.docx");
					}
				} else if (pinduan.trim().equals("E频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/拉远站室内E频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/拉远站室内E频段.docx");
					}
				} else if (pinduan.trim().equals("D频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/拉远站室内D频段无天线.docx");
					} else {

						is = new FileInputStream("testfiles/拉远站室内D频段.docx");
					}
				}
			} else if (neiwai.trim().equals("室外站")) {
				if (pinduan.trim().equals("F频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/拉远站室外F频段无天线.docx");
					} else {

						is = new FileInputStream("testfiles/拉远站室外F频段.docx");
					}
				} else if (pinduan.trim().equals("E频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/拉远站室外E频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/拉远站室外E频段.docx");

					}
				} else if (pinduan.trim().equals("D频段")) {
					if (excelmap.get(i).get(19).toString().equals("AAU3213")
							|| excelmap.get(i).get(19).toString().equals("BTS3205E")
							|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
							|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
						is = new FileInputStream("testfiles/拉远站室外D频段无天线.docx");
					} else {
						is = new FileInputStream("testfiles/拉远站室外D频段.docx");
					}
				}
			}
		} else if (name.trim().equals("信源站")) {
			if (neiwai.trim().equals("室内站")) {
				if (pinduan.trim().equals("F频段")) {
					is = new FileInputStream("testfiles/信源站室内F频段.docx");
				} else if (pinduan.trim().equals("E频段")) {
					is = new FileInputStream("testfiles/信源站室内E频段.docx");
				} else if (pinduan.trim().equals("D频段")) {
					is = new FileInputStream("testfiles/信源站室内D频段.docx");
				}
			} else if (neiwai.trim().equals("室外站")) {
				if (pinduan.trim().equals("F频段")) {
					is = new FileInputStream("testfiles/信源站室外F频段.docx");
				} else if (pinduan.trim().equals("E频段")) {
					is = new FileInputStream("testfiles/信源站室外E频段.docx");
				} else if (pinduan.trim().equals("D频段")) {
					is = new FileInputStream("testfiles/信源站室外D频段.docx");
				}
			}
		}
		XWPFDocument doc = new XWPFDocument(is);
		List<XWPFTable> tables = doc.getTables();
		if (name.trim().equals("信源站")) {
			XWPFTable table0 = tables.get(0);
			setCellText(table0.getRow(0).getCell(1), excelmap.get(i).get(22).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(1).getCell(1), pinduan.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(2).getCell(1), name.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(23).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(24).toString().trim(), "FFFFFF", 21);
		} else if (name.trim().equals("拉远站") || name.trim().equals("宏基站")) {
			XWPFTable table0 = tables.get(0);
			setCellText(table0.getRow(0).getCell(1), excelmap.get(i).get(24).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(1).getCell(1), pinduan.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(2).getCell(1), name.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(0).getCell(3), shebei.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(25).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(26).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(29).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(30).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(4).getCell(3), excelmap.get(i).get(28).toString().trim(), "FFFFFF", 21);
			// TODO
			/*if (MatcherXYZ(excelmap.get(i).get(25).toString().trim())) {
				
				if (MatcherXYZ(excelmap.get(i).get(26).toString().trim())) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线挂高(格式如：x/y/z)！");
					throw NullPointerException;
				}
				if (MatcherXYZ(excelmap.get(i).get(27).toString().trim())) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的总下倾角(格式如：x/y/z)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(29).toString().trim().equals("S111")
						|| excelmap.get(i).get(29).toString().trim().equals("s111")) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的配置(S111或者s111)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(30).toString().trim().equals("3")
						|| excelmap.get(i).get(30).toString().trim().equals("3.0")
						|| Integer.parseInt(excelmap.get(i).get(30)) == 3
						|| Double.parseDouble(excelmap.get(i).get(30)) == 3.0) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确RRU数量！");
					throw NullPointerException;
				}
			} else if (MatcherXY(excelmap.get(i).get(25).toString().trim())) {
				setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(25).toString().trim(), "FFFFFF", 21);
				if (MatcherXY(excelmap.get(i).get(26).toString().trim())) {
					setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(26).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线挂高(格式如：x/y)！");
					throw NullPointerException;
				}
				if (MatcherXY(excelmap.get(i).get(27).toString().trim())) {
					setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的总下倾角(格式如：x/y)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(29).toString().trim().equals("S11")
						|| excelmap.get(i).get(29).toString().trim().equals("s11")) {
					setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(29).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的配置(S11或者s11)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(30).toString().trim().equals("2")
						|| excelmap.get(i).get(30).toString().trim().equals("2.0")
						|| Integer.parseInt(excelmap.get(i).get(30)) == 2
						|| Double.parseDouble(excelmap.get(i).get(30)) == 2.0) {
					setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(30).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正RRU数量！");
					throw NullPointerException;
				}
			} else if (matcherX(excelmap.get(i).get(25).toString().trim())) {
				setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(25).toString().trim(), "FFFFFF", 21);
				if (matcherX(excelmap.get(i).get(26).toString().trim())) {
					setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(26).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线挂高(格式如：x)！");
					throw NullPointerException;
				}
				if (matcherX(excelmap.get(i).get(27).toString().trim())) {
					setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的总下倾角(格式如：x)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(29).toString().trim().equals("S1")
						|| excelmap.get(i).get(29).toString().trim().equals("s1")) {
					setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(29).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的配置(S1或者s1)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(30).toString().trim().equals("1")
						|| excelmap.get(i).get(30).toString().trim().equals("1.0")
						|| Integer.parseInt(excelmap.get(i).get(30)) == 1
						|| Double.parseDouble(excelmap.get(i).get(30)) == 1.0) {
					setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(30).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正RRU数量！");
					throw NullPointerException;
				}
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线方位角(格式如：x/y/z,或者x/y,或者x)！");
				throw NullPointerException;
			}*/
			
		} else if (name.trim().equals("小基站")) {
			XWPFTable table0 = tables.get(0);
			// TODO excelmap里的顺序需要重新改过
			setCellText(table0.getRow(0).getCell(1), excelmap.get(i).get(22).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(1).getCell(1), pinduan.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(2).getCell(1), name.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(0).getCell(3), shebei.toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(4).getCell(3), excelmap.get(i).get(26).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(23).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(24).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(25).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
			setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(28).toString().trim(), "FFFFFF", 21);
			/*if (MatcherXYZ(excelmap.get(i).get(23).toString().trim())) {
				
				if (MatcherXYZ(excelmap.get(i).get(24).toString().trim())) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线挂高(格式如：x/y/z)！");
					throw NullPointerException;
				}
				if (MatcherXYZ(excelmap.get(i).get(25).toString().trim())) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的总下倾角(格式如：x/y/z)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(27).toString().trim().equals("S111")
						|| excelmap.get(i).get(27).toString().trim().equals("s111")) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的配置(S111或者s111)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(28).toString().trim().equals("3")
						|| excelmap.get(i).get(28).toString().trim().equals("3.0")
						|| Integer.parseInt(excelmap.get(i).get(28)) == 3
						|| Double.parseDouble(excelmap.get(i).get(28)) == 3.0) {
					
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确RRU数量！");
					throw NullPointerException;
				}
			} else if (MatcherXY(excelmap.get(i).get(23).toString().trim())) {
				setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(23).toString().trim(), "FFFFFF", 21);
				if (MatcherXY(excelmap.get(i).get(24).toString().trim())) {
					setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(24).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线挂高(格式如：x/y)！");
					throw NullPointerException;
				}
				if (MatcherXY(excelmap.get(i).get(25).toString().trim())) {
					setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(25).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的总下倾角(格式如：x/y)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(27).toString().trim().equals("S11")
						|| excelmap.get(i).get(27).toString().trim().equals("s11")) {
					setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的配置(S11或者s11)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(28).toString().trim().equals("2")
						|| excelmap.get(i).get(28).toString().trim().equals("2.0")
						|| Integer.parseInt(excelmap.get(i).get(28)) == 2
						|| Double.parseDouble(excelmap.get(i).get(28)) == 2.0) {
					setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(28).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确RRU数量！");
					throw NullPointerException;
				}
			} else if (matcherX(excelmap.get(i).get(23).toString().trim())) {
				setCellText(table0.getRow(1).getCell(3), excelmap.get(i).get(23).toString().trim(), "FFFFFF", 21);
				if (matcherX(excelmap.get(i).get(24).toString().trim())) {
					setCellText(table0.getRow(2).getCell(3), excelmap.get(i).get(24).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线挂高(格式如：x)！");
					throw NullPointerException;
				}
				if (matcherX(excelmap.get(i).get(25).toString().trim())) {
					setCellText(table0.getRow(3).getCell(3), excelmap.get(i).get(25).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的总下倾角(格式如：x)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(27).toString().trim().equals("S1")
						|| excelmap.get(i).get(27).toString().trim().equals("s1")) {
					setCellText(table0.getRow(3).getCell(1), excelmap.get(i).get(27).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的配置(S1或者s1)！");
					throw NullPointerException;
				}
				if (excelmap.get(i).get(28).toString().trim().equals("1")
						|| excelmap.get(i).get(28).toString().trim().equals("1.0")
						|| Integer.parseInt(excelmap.get(i).get(28)) == 1
						|| Double.parseDouble(excelmap.get(i).get(28)) == 1.0) {
					setCellText(table0.getRow(4).getCell(1), excelmap.get(i).get(28).toString().trim(), "FFFFFF", 21);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确RRU数量！");
					throw NullPointerException;
				}
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线方位角(格式如：x/y/z,或者x/y,或者x)！");
				throw NullPointerException;
			}*/
		}

		// BBU表格生成
		XWPFTable tableBBU = tables.get(2);
		// XWPFTableRow tBBURow = tableBBU.createRow();
		tableBBU.addRow(tableBBU.getRow(0), 0);
		XWPFTableRow tBBURow = tableBBU.getRow(1);
		// tBBURow.setHeight(11);
		// System.out.println(BBUtablemap);
		for (int j = 0; j < BBUtablemap.get(i).size(); j++) {
			setCellText(tBBURow.getCell(j), BBUtablemap.get(i).get(j), "FFFFFF", 21);
			if (tBBURow.getCell(j).getParagraphs().size() > 1) {
				tBBURow.getCell(j).removeParagraph(1);
			}
		}

		if (excelmap.get(0).get(19).toString().trim().equals("RRU型号")) {
			XWPFTable tableRRU = tables.get(3);
			XWPFTableRow row = tableRRU.getRow(0);
			// System.out.println(RRUtablemap.get(4));
			if (RRUtablemap.get(i).get(8).toString().equals("工作频带宽度 ")) {
				mergeCellsHorizontal(tableRRU, 0, 5, 7);
				// XWPFTableRow tRRURow = tableRRU.createRow();
				tableRRU.addRow(tableRRU.getRow(0), 0);
				XWPFTableRow tRRURow = tableRRU.getRow(1);
				mergeCellsHorizontal(tableRRU, 1, 5, 7);
				// tRRURow.setHeight(11);
				setCellText(tRRURow.getCell(0), RRUtablemap.get(i).get(0), "FFFFFF", 21);
				setCellText(tRRURow.getCell(1), RRUtablemap.get(i).get(1), "FFFFFF", 21);
				setCellText(tRRURow.getCell(2), RRUtablemap.get(i).get(2), "FFFFFF", 21);
				setCellText(tRRURow.getCell(3), RRUtablemap.get(i).get(3), "FFFFFF", 21);
				setCellText(tRRURow.getCell(4), RRUtablemap.get(i).get(4), "FFFFFF", 21);
				setCellText(tRRURow.getCell(5), RRUtablemap.get(i).get(5), "FFFFFF", 21);
				setCellText(tRRURow.getCell(8), RRUtablemap.get(i).get(6), "FFFFFF", 21);
				setCellText(tRRURow.getCell(9), RRUtablemap.get(i).get(7), "FFFFFF", 21);
				if (tRRURow.getCell(9).getParagraphs().size() > 1) {
					tRRURow.getCell(9).removeParagraph(1);
				}
			}
			 System.out.println(RRUtablemap);
			if (RRUtablemap.get(i).get(8).toString().equals("功耗")) {
				mergeCellsHorizontal(tableRRU, 0, 5, 6);
				XWPFTableCell cell = row.getCell(5);
				setCellText(cell, "供电方式", "FFFFFF", 21);
				tableRRU.addRow(tableRRU.getRow(0), 0);
				XWPFTableRow tRRURow = tableRRU.getRow(1);
				mergeCellsHorizontal(tableRRU, 1, 5, 6);
				// tRRURow.setHeight(11);

				setCellText(tRRURow.getCell(0), RRUtablemap.get(i).get(0), "FFFFFF", 21);
				setCellText(tRRURow.getCell(1), RRUtablemap.get(i).get(1), "FFFFFF", 21);
				setCellText(tRRURow.getCell(2), RRUtablemap.get(i).get(2), "FFFFFF", 21);
				setCellText(tRRURow.getCell(3), RRUtablemap.get(i).get(3), "FFFFFF", 21);
				setCellText(tRRURow.getCell(4), RRUtablemap.get(i).get(4), "FFFFFF", 21);
				setCellText(tRRURow.getCell(5), RRUtablemap.get(i).get(9), "FFFFFF", 21);
				setCellText(tRRURow.getCell(7), RRUtablemap.get(i).get(5), "FFFFFF", 21);
				setCellText(tRRURow.getCell(8), RRUtablemap.get(i).get(6), "FFFFFF", 21);
				setCellText(tRRURow.getCell(9), RRUtablemap.get(i).get(7), "FFFFFF", 21);
				if (tRRURow.getCell(9).getParagraphs().size() > 1) {
					tRRURow.getCell(9).removeParagraph(1);
				}
			}
		}
		// System.out.println(excelmap.get(0).get(23));
		if (excelmap.get(0).get(21).toString().trim().equals("天线型号")) {
			if (!excelmap.get(i).get(19).toString().equals("AAU3213")
					&& !excelmap.get(i).get(19).toString().equals("BTS3205E")
					&& !excelmap.get(i).get(19).toString().equals("BBOK RRU")
					&& !excelmap.get(i).get(19).toString().equals("Easymacro")) {
				XWPFTable tableAnn = tables.get(4);
				tableAnn.addRow(tableAnn.getRow(0), 0);
				XWPFTableRow tAnnRow = tableAnn.getRow(1);
				for (int j = 0; j < Antennatablemap.get(i).size(); j++) {
					setCellText(tAnnRow.getCell(j), Antennatablemap.get(i).get(j).trim(), "FFFFFF", 21);
					if (tAnnRow.getCell(j).getParagraphs().size() > 1) {
						tAnnRow.getCell(j).removeParagraph(1);
					}
				}
			}
			// CTTblBorders borders =
			// tableAnn.getCTTbl().getTblPr().addNewTblBorders();
			// genBorders(borders);
			// tAnnRow.getCell(2).removeParagraph(0);
			// tAnnRow.getCell(2).setText("asdfsdaf");

			// tAnnRow.getCell(0).removeParagraph(0);
			// setRowHeight(tAnnRow, "567", STHeightRule.AT_LEAST);
			// tAnnRow.setHeight(11);

			// System.out.println(tableAnn.getRow(0).getCell(2).getText());
			// System.out.println(tAnnRow.getCell(2).getText()+"!!!!!!!!!!!");
		}

		XWPFTable table = tables.get(1);
		CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
		genBorders(borders);
		for (String key : datas_map.keySet()) {

			XWPFTableRow tableOneRowTwo = table.createRow();
			tableOneRowTwo.setHeight(11);
			setCellText(tableOneRowTwo.getCell(0), key, "FFFFFF", 21);
			setCellText(tableOneRowTwo.getCell(1), datas_map.get(key).get(0), "FFFFFF", 21);
			setCellText(tableOneRowTwo.getCell(2), datas_map.get(key).get(1), "FFFFFF", 21);
		}

		String zh = ub2.getZhFrom4GYsb();
		List<String> datas = ub2.readExcel(zh);

		if (name.toString().equals("信源站") || name.toString().equals("小基站")) {
			table = tables.get(5);
		} else if (name.toString().equals("宏基站") || name.toString().equals("拉远站")) {
			if (excelmap.get(i).get(19).toString().equals("AAU3213")
					|| excelmap.get(i).get(19).toString().equals("BTS3205E")
					|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
					|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
				table = tables.get(5);
			} else {
				table = tables.get(6);
			}
		}
		// borders = table.getCTTbl().getTblPr().addNewTblBorders();
		// genBorders(borders);
		table.addRow(table.getRow(0), 0);
		XWPFTableRow tableOneRowTwo = table.getRow(1);
		// XWPFTableRow tableOneRowTwo = table.createRow();
		// tableOneRowTwo.setHeight(11);
		int k = 0;
		for (String values : datas) {
			if (values.equals("") || values == null) {
				tableOneRowTwo.getCell(k).removeParagraph(0);
			}
			tableOneRowTwo.getCell(k).addParagraph();
			setCellText(tableOneRowTwo.getCell(k), values, "FFFFFF", 21);
			if (tableOneRowTwo.getCell(k).getParagraphs().size() > 1) {
				tableOneRowTwo.getCell(k).removeParagraph(1);
			}
			k++;
		}

		datas = ub3.readExcel(zh);
		if (name.toString().trim().equals("信源站") || name.toString().trim().equals("小基站")) {
			table = tables.get(6);
		} else if (name.toString().trim().equals("宏基站") || name.toString().trim().equals("拉远站")) {
			if (excelmap.get(i).get(19).toString().equals("AAU3213")
					|| excelmap.get(i).get(19).toString().equals("BTS3205E")
					|| excelmap.get(i).get(19).toString().equals("BBOK RRU")
					|| excelmap.get(i).get(19).toString().equals("Easymacro")) {
				table = tables.get(6);
			} else {
				table = tables.get(7);
			}
		}
		// borders = table.getCTTbl().getTblPr().addNewTblBorders();
		// genBorders(borders);
		// tableOneRowTwo = table.createRow();
		// tableOneRowTwo.setHeight(11);
		table.addRow(table.getRow(0), 0);
		tableOneRowTwo = table.getRow(1);
		k = 0;
		for (String values : datas) {
			if (values.equals("") || values == null) {
				tableOneRowTwo.getCell(k).removeParagraph(0);
			}
			tableOneRowTwo.getCell(k).addParagraph();
			setCellText(tableOneRowTwo.getCell(k), values, "FFFFFF", 21);
			if (tableOneRowTwo.getCell(k).getParagraphs().size() > 1) {
				tableOneRowTwo.getCell(k).removeParagraph(1);
			}
			k++;
		}

		// OPCPackage pack = POIXMLDocument.openPackage(template);
		// XWPFDocument doc = new XWPFDocument(pack);
		List<XWPFParagraph> paragraphs = doc.getParagraphs();
		// System.out.println(paragraphs.size());
		for (XWPFParagraph tmp : paragraphs) {
			// System.out.println(tmp.getParagraphText());
			List<XWPFRun> runs = tmp.getRuns();
			for (XWPFRun aa : runs) {
				 //System.out.println("XWPFRun-Text:" + aa.getText(0));
				if ("city".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(0), 0);
				}
				if ("option2".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(1), 0);
				}
				if ("option3".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(2), 0);
				}
				if ("bh".equals(aa.getText(0)) || "option4".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(3), 0);
				}
				if ("option5".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(4), 0);
				}
				if ("option6".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(5), 0);
				}
				if ("option7".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(6), 0);
				}
				if ("option8".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(7), 0);
				}
				if ("option9".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(8), 0);
				}
				if ("option10".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(9), 0);
				}
				if ("option11".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(10), 0);
				}
				if ("option12".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(11), 0);
				}
				if ("option13".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(12), 0);
				}
				if ("option15".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(17) + "；", 0);
				}
				if ("option16".equals(aa.getText(0))) {
					aa.setText(shebei, 0);
				}
				if ("option17".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(18), 0);
				}
				if ("option18".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(20), 0);
				}
				if ("hongjizhan".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(13), 0);
				}
				if ("xiaojizhan".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(14), 0);
				}
				if ("layuanzhan".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(15), 0);
				}
				if ("xinyuanzhan".equals(aa.getText(0))) {
					aa.setText(excelmap.get(i).get(16), 0);
				}
				// System.out.println(excelmap.get(i).get(25).toString().trim().equals("挂墙机框"));
				// System.out.println(ret.anzhuangHongandLa(tablePath,
				// excelPath));
				// System.out.println(AZHoLa.get(0));
				if ("option19".equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(22), 0);
					} else if (name.trim().equals("小基站") || name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(20), 0);
					}

				}
				if ("fx".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(31), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(29), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(25), 0);
					}
				}
				if ("gcmc".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(32), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(30), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(26), 0);
					}
				}

				if ("ds".trim().equals(aa.getText(0)) || "ts".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(33), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(31), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(27), 0);
					}
				}
				if ("ds2".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(34), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(32), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(28), 0);
					}
				}
				if ("ds3".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(35), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(33), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(29), 0);
					}
				}

				if ("ds4".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(36), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(34), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(30), 0);
					}
				}

				if ("ds5".trim().equals(aa.getText(0))) {
					if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
						aa.setText(excelmap.get(i).get(37), 0);
					} else if (name.trim().equals("小基站")) {
						aa.setText(excelmap.get(i).get(35), 0);
					} else if (name.trim().equals("信源站")) {
						aa.setText(excelmap.get(i).get(31), 0);
					}
				}

				if ("fs".trim().equals(aa.getText(0))) {
					aa.setText(AZHoLa.get(i), 0);
				}
				if ("sb".trim().equals(aa.getText(0))) {
					aa.setText(shebei, 0);
				}
				if ("jine".trim().equals(aa.getText(0))) {
					aa.setText(ub2.readExcel(zh).get(ub2.readExcel(zh).size() - 1), 0);
				}
				if ("jzh".trim().equals(aa.getText(0))) {
					int hong = Integer.parseInt(excelmap.get(i).get(13));
					int xiao = Integer.parseInt(excelmap.get(i).get(14));
					int la = Integer.parseInt(excelmap.get(i).get(15));
					int xin = Integer.parseInt(excelmap.get(i).get(16));
					String he = String.valueOf(hong + xiao + la + xin);
					aa.setText(he, 0);
				}
			}
		}
		if (name.trim().equals("宏基站") || name.trim().equals("拉远站")) {
			simpleDateHeader2(doc, excelmap.get(i).get(32));
		} else if (name.trim().equals("小基站")) {
			simpleDateHeader2(doc, excelmap.get(i).get(30));
		} else if (name.trim().equals("信源站")) {
			simpleDateHeader2(doc, excelmap.get(i).get(26));
		}

		// System.out.println(doc.getHeaderList());
		fos = new FileOutputStream(output + excelmap.get(i).get(1) + ".docx");
		doc.write(fos);
		fos.flush();
		fos.close();
		is.close();
		wb.close();
		fise.close();

	}

	private void genBorders(CTTblBorders borders) {
		CTBorder hBorder = borders.addNewInsideH();
		hBorder.setVal(STBorder.Enum.forString("thick"));
		hBorder.setSz(new BigInteger("1"));
		hBorder.setColor("000000");
		//
		CTBorder vBorder = borders.addNewInsideV();
		vBorder.setVal(STBorder.Enum.forString("thick"));
		vBorder.setSz(new BigInteger("1"));
		vBorder.setColor("000000");
		//
		CTBorder lBorder = borders.addNewLeft();
		lBorder.setVal(STBorder.Enum.forString("thick"));
		lBorder.setSz(new BigInteger("1"));
		lBorder.setColor("000000");
		//
		CTBorder rBorder = borders.addNewRight();
		rBorder.setVal(STBorder.Enum.forString("thick"));
		rBorder.setSz(new BigInteger("1"));
		rBorder.setColor("000000");
		//
		CTBorder tBorder = borders.addNewTop();
		tBorder.setVal(STBorder.Enum.forString("thick"));
		tBorder.setSz(new BigInteger("1"));
		tBorder.setColor("000000");
		//
		CTBorder bBorder = borders.addNewBottom();
		bBorder.setVal(STBorder.Enum.forString("thick"));
		bBorder.setSz(new BigInteger("1"));
		bBorder.setColor("000000");
	}

	public void setCellText(XWPFTableCell cell, String text, String bgcolor, int width) {
		/*
		 * CTTc cttc = cell.getCTTc(); CTTcPr ctPr = cttc.isSetTcPr() ?
		 * cttc.getTcPr() : cttc.addNewTcPr();
		 * ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		 * cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.LEFT);
		 */
		setCellWidthAndVAlign(cell, "2169", STTblWidth.DXA, STVerticalJc.CENTER);
		XWPFParagraph cellP = cell.getParagraphs().get(0);
		XWPFRun cellR = getOrAddParagraphFirstRun(cellP, false, false);
		// XWPFRun cellR = cellP.createRun();
		/*
		 * cellR.addBreak(); cellR.setFontSize(11); cellR.setText(text,0);
		 * cell.removeParagraph(0);
		 */
		setParagraphRunFontInfo(cellP, cellR, text, "宋体", "Times New Roman", "21", false, false, false, false, null,
				null, 0, 6, 0);
	}

	public void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCell) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one, are set with
				// CONTINUE
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	public void simpleDateHeader2(XWPFDocument document, String gcmc) throws Exception {
		CTHdrFtr headerFooter = CTHdrFtr.Factory.newInstance();
		CTP ctp = headerFooter.addNewP();
		XWPFParagraph codePara = new XWPFParagraph(ctp, document);
		XWPFRun r1 = codePara.createRun();

		r1 = codePara.createRun();
		CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		r1 = codePara.createRun();
		r1.setText(gcmc + "一阶段设计        CMDI");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("宋体");
		fonts.setEastAsia("宋体");
		fonts.setHAnsi("宋体");

		codePara.setAlignment(ParagraphAlignment.RIGHT);
		codePara.setVerticalAlignment(TextAlignment.AUTO);
		codePara.setBorderBottom(Borders.THICK);
		XWPFParagraph[] newparagraphs = new XWPFParagraph[1];
		newparagraphs[0] = codePara;
		List<XWPFHeader> headers = document.getHeaderList();
		for (XWPFHeader header : headers) {
			if (header.getPackagePart().getPartName().toString().equals("/word/header1.xml")) {
				continue;
			}
			if (header.getPackagePart().getPartName().toString().equals("/word/header2.xml")) {
				continue;
			}
			// System.out.println(header.getPackagePart().getPartName());
			header.setHeaderFooter(headerFooter);
		}
	}

	public XWPFRun getOrAddParagraphFirstRun(XWPFParagraph p, boolean isInsert, boolean isNewLine) {
		XWPFRun pRun = null;
		if (isInsert) {
			pRun = p.createRun();
		} else {
			if (p.getRuns() != null && p.getRuns().size() > 0) {
				pRun = p.getRuns().get(0);
			} else {
				pRun = p.createRun();
			}
		}
		if (isNewLine) {
			pRun.addBreak();
		}
		return pRun;
	}

	public CTRPr getRunCTRPr(XWPFParagraph p, XWPFRun pRun) {
		CTRPr pRpr = null;
		if (pRun.getCTR() != null) {
			pRpr = pRun.getCTR().getRPr();
			if (pRpr == null) {
				pRpr = pRun.getCTR().addNewRPr();
			}
		} else {
			pRpr = p.getCTP().addNewR().addNewRPr();
		}
		return pRpr;
	}

	public void setParagraphRunFontInfo(XWPFParagraph p, XWPFRun pRun, String content, String cnFontFamily,
			String enFontFamily, String fontSize, boolean isBlod, boolean isItalic, boolean isStrike, boolean isShd,
			String shdColor, STShd.Enum shdStyle, int position, int spacingValue, int indent) {
		CTRPr pRpr = getRunCTRPr(p, pRun);
		if (StringUtils.isNotBlank(content)) {
			// pRun.setText(content);
			if (content.contains("\n")) {// System.properties("line.separator")
				String[] lines = content.split("\n");
				pRun.setText(lines[0], 0); // set first line into XWPFRun
				for (int i = 1; i < lines.length; i++) {
					// add break and insert new text
					pRun.addBreak();
					pRun.setText(lines[i]);
				}
			} else {
				pRun.setText(content, 0);
			}
		}
		// 设置字体
		CTFonts fonts = pRpr.isSetRFonts() ? pRpr.getRFonts() : pRpr.addNewRFonts();
		if (StringUtils.isNotBlank(enFontFamily)) {
			fonts.setAscii(enFontFamily);
			fonts.setHAnsi(enFontFamily);
		}
		if (StringUtils.isNotBlank(cnFontFamily)) {
			fonts.setEastAsia(cnFontFamily);
			fonts.setHint(STHint.EAST_ASIA);
		}
		// 设置字体大小
		CTHpsMeasure sz = pRpr.isSetSz() ? pRpr.getSz() : pRpr.addNewSz();
		sz.setVal(new BigInteger(fontSize));

		CTHpsMeasure szCs = pRpr.isSetSzCs() ? pRpr.getSzCs() : pRpr.addNewSzCs();
		szCs.setVal(new BigInteger(fontSize));

		// 设置字体样式
		// 加粗
		if (isBlod) {
			pRun.setBold(isBlod);
		}
		// 倾斜
		if (isItalic) {
			pRun.setItalic(isItalic);
		}
		// 删除线
		if (isStrike) {
			pRun.setStrike(isStrike);
		}
		if (isShd) {
			// 设置底纹
			CTShd shd = pRpr.isSetShd() ? pRpr.getShd() : pRpr.addNewShd();
			if (shdStyle != null) {
				shd.setVal(shdStyle);
			}
			if (shdColor != null) {
				shd.setColor(shdColor);
				shd.setFill(shdColor);
			}
		}

		// 设置文本位置
		if (position != 0) {
			pRun.setTextPosition(position);
		}
		if (spacingValue > 0) {
			// 设置字符间距信息
			CTSignedTwipsMeasure ctSTwipsMeasure = pRpr.isSetSpacing() ? pRpr.getSpacing() : pRpr.addNewSpacing();
			ctSTwipsMeasure.setVal(new BigInteger(String.valueOf(spacingValue)));
		}
		if (indent > 0) {
			CTTextScale paramCTTextScale = pRpr.isSetW() ? pRpr.getW() : pRpr.addNewW();
			paramCTTextScale.setVal(indent);
		}
	}

	public CTTcPr getCellCTTcPr(XWPFTableCell cell) {
		CTTc cttc = cell.getCTTc();
		CTTcPr tcPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
		return tcPr;
	}

	public void setCellWidthAndVAlign(XWPFTableCell cell, String width, STTblWidth.Enum typeEnum,
			STVerticalJc.Enum vAlign) {
		CTTcPr tcPr = getCellCTTcPr(cell);
		CTTblWidth tcw = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
		if (width != null) {
			tcw.setW(new BigInteger(width));
		}
		if (typeEnum != null) {
			tcw.setType(typeEnum);
		}
		if (vAlign != null) {
			CTVerticalJc vJc = tcPr.isSetVAlign() ? tcPr.getVAlign() : tcPr.addNewVAlign();
			vJc.setVal(vAlign);
		}
	}

	public void setRowHeight(XWPFTableRow row, String hight, STHeightRule.Enum heigthEnum) {
		CTTrPr trPr = getRowCTTrPr(row);
		CTHeight trHeight;
		if (trPr.getTrHeightList() != null && trPr.getTrHeightList().size() > 0) {
			trHeight = trPr.getTrHeightList().get(0);
		} else {
			trHeight = trPr.addNewTrHeight();
		}
		trHeight.setVal(new BigInteger(hight));
		if (heigthEnum != null) {
			trHeight.setHRule(heigthEnum);
		}
	}

	public CTTrPr getRowCTTrPr(XWPFTableRow row) {
		CTRow ctRow = row.getCtRow();
		CTTrPr trPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
		return trPr;
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

	public boolean matcherX(String s) throws Exception {
		Pattern pattern = Pattern.compile("^\\d{1,3}$");
		Matcher matcher = pattern.matcher(s);
		if (matcher.find()) {
			return true;
		} else {

			return false;
		}
	}
}
