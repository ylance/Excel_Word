package com.cm.oe.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ReadExcelTable {
	private String  excelPath;
	private String tablePath;
	private static final Exception NullPointerException = null;
	public ReadExcelTable(String tablePath,String excelPath) {
		this.tablePath = tablePath;
		this.excelPath  = excelPath;
	}
	ReadExcel re = new ReadExcel(excelPath);
	
	public List<Row> genRowlist() throws Exception {
		List<Row> rowe = new ArrayList<Row>();
		try {
			FileInputStream fise = new FileInputStream(excelPath);
			HSSFWorkbook wbe = new HSSFWorkbook(fise);
			Sheet sht_e = wbe.getSheetAt(0);
			for (int i = 0; i < re.rowNumber(wbe); i++) {
				rowe.add(sht_e.getRow(i));
			}
			wbe.close();
			fise.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return rowe;
	}

	public Map<Integer, List<String>> readBBUinExcel() throws Exception {
		File file = new File(excelPath);
		String[] strs = file.getName().split("-");
		String shebei = strs[3];
		System.out.println(shebei);
		FileInputStream fist = new FileInputStream(tablePath);
		HSSFWorkbook wbt = new HSSFWorkbook(fist);
		Sheet sht_t = wbt.getSheetAt(0);
		List<Row> rowe = genRowlist();

		Row rowt = null;
		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();
		for (int i = 1; i < rowe.size(); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (shebei.toString().trim().equals("华为")) {
					rowt = sht_t.getRow(0);
					tablevalues.add(rowt.getCell(0).toString());
					for (int j = 1; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
			} else if (shebei.toString().trim().equals("大唐")) {
					rowt = sht_t.getRow(3);
					tablevalues.add(rowt.getCell(0).toString());
					for (int j = 1; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
			} else if (shebei.toString().trim().equals("中兴")) {
					rowt = sht_t.getRow(1);
					tablevalues.add(rowt.getCell(0).toString());
					for (int j = 1; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}

					tableMap.put(i, tablevalues);
			} else if (shebei.toString().trim().equals("上海贝尔")) {
					rowt = sht_t.getRow(2);
					tablevalues.add(rowt.getCell(0).toString());
					for (int j = 1; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
			} 
		}
		wbt.close();
		fist.close();
		return tableMap;

	}

	public Map<Integer, List<String>> readRRUinExcel() throws Exception {
		FileInputStream fist = new FileInputStream(tablePath);
		HSSFWorkbook wbt = new HSSFWorkbook(fist);
		Sheet sht_t = wbt.getSheetAt(0);
		List<Row> rowe = genRowlist();
		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();
		Row rowt = null;
		for (int i = 1; i < rowe.size(); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (rowe.get(i).getCell(19).toString().trim().equals("华为")) {
				if (rowe.get(i).getCell(20).toString().trim().equals("AAU3213")) {
					rowt = sht_t.getRow(4);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("RRU3277")) {
					rowt = sht_t.getRow(5);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("DRRU3168e-fa")) {
					rowt = sht_t.getRow(6);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("RRU3278M")) {
					rowt = sht_t.getRow(7);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("DRRU3172-fad")) {
					rowt = sht_t.getRow(8);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("DRRU3161-fae")) {
					rowt = sht_t.getRow(9);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("BTS3205E")) {
					rowt = sht_t.getRow(10);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("BBOK RRU")) {
					rowt = sht_t.getRow(11);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("Easymacro")) {
					rowt = sht_t.getRow(12);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的华为RRU设备型号！");
					throw NullPointerException;
				}
			} else if (rowe.get(i).getCell(19).toString().trim().equals("中兴")) {
				if (rowe.get(i).getCell(20).toString().trim().equals("ZXSDR R8978 S2600W")) {
					rowt = sht_t.getRow(13);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("ZXSDR R8972E S2600W")) {
					rowt = sht_t.getRow(14);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("ZXSDR R8978 M1920A")) {
					rowt = sht_t.getRow(15);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("ZXSDR R8972E M1920A")) {
					rowt = sht_t.getRow(16);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("ZXSDR R8972E S2300W")) {
					rowt = sht_t.getRow(17);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("ZXSDR R8972E M192023A")) {
					rowt = sht_t.getRow(18);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的中兴RRU设备型号！");
					throw NullPointerException;
				}
			} else if (rowe.get(i).getCell(19).toString().trim().equals("大唐")) {
				if (rowe.get(i).getCell(20).toString().trim().equals("TDRU348FA")) {
					rowt = sht_t.getRow(19);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TDRU348D")) {
					rowt = sht_t.getRow(20);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TDRU342D")) {
					rowt = sht_t.getRow(21);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TDRU342FA")) {
					rowt = sht_t.getRow(22);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TDRU342E")) {
					rowt = sht_t.getRow(23);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TDRU341FAE")) {
					rowt = sht_t.getRow(24);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("mTDRU342D")) {
					rowt = sht_t.getRow(25);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的大唐RRU设备型号！");
					throw NullPointerException;
				}
			} else if (rowe.get(i).getCell(19).toString().equals("上海贝尔")) {
				if (rowe.get(i).getCell(20).toString().trim().equals("TD-RRH8X20-25A")) {
					rowt = sht_t.getRow(26);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TD-RRH2×40-25A")) {
					rowt = sht_t.getRow(27);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TD-RRH8x10-1935")) {
					rowt = sht_t.getRow(28);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TD-RRH2x60-1935")) {
					rowt = sht_t.getRow(29);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("TD-RRH2X50-2350")) {
					rowt = sht_t.getRow(30);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(20).toString().trim().equals("9768 MRO B38 TD-LTE 2x5W")) {
					rowt = sht_t.getRow(31);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的上海贝尔RRU设备的型号！");
					throw NullPointerException;
				}
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的RRU品牌！");
				throw NullPointerException;
			}
		}
		wbt.close();
		fist.close();
		return tableMap;
	}

	public Map<Integer, List<String>> readAntennaIntables() throws Exception {
		FileInputStream fist = new FileInputStream(tablePath);
		HSSFWorkbook wbt = new HSSFWorkbook(fist);
		Sheet sht_t = wbt.getSheetAt(0);
		List<Row> rowe = genRowlist();

		Map<Integer, List<String>> tableMap = new HashMap<Integer, List<String>>();
		Row rowt = null;
		for (int i = 1; i < rowe.size(); i++) {
			List<String> tablevalues = new ArrayList<String>();
			if (rowe.get(i).getCell(21).toString().trim().equals("华为")) {
				if (rowe.get(i).getCell(22).toString().trim().equals("ATD-")) {
					rowt = sht_t.getRow(32);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("ATD451601")) {
					rowt = sht_t.getRow(33);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("ATD451602")) {
					rowt = sht_t.getRow(34);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("ATD451603")) {
					rowt = sht_t.getRow(35);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("ATD4516R0")) {
					rowt = sht_t.getRow(36);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("-")) {
					rowt = sht_t.getRow(37);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("ATD451800")) {
					rowt = sht_t.getRow(38);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT0")) {
					rowt = sht_t.getRow(39);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT3")) {
					rowt = sht_t.getRow(40);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT6")) {
					rowt = sht_t.getRow(41);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT9")) {
					rowt = sht_t.getRow(42);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的华为天线型号！");
					throw NullPointerException;
				}
			} else if (rowe.get(i).getCell(21).toString().trim().equals("中兴")) {
				if (rowe.get(i).getCell(22).toString().trim().equals("T-04-52-50-002")) {
					rowt = sht_t.getRow(43);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("T-03-52-52-003")) {
					rowt = sht_t.getRow(44);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("T-DA-02-00-59")) {
					rowt = sht_t.getRow(45);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("T-12-54-18-002")) {
					rowt = sht_t.getRow(46);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "第" + (i + 1) + "行" + "请输入正确的中兴天线型号！");
					throw NullPointerException;
				}
			} else if (rowe.get(i).getCell(21).toString().trim().equals("大唐")
					|| rowe.get(i).getCell(21).toString().trim().equals("上海贝尔")) {
				if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-202616D4T0")) {
					rowt = sht_t.getRow(47);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-202616D4T3")) {
					rowt = sht_t.getRow(48);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-202616D4T6")) {
					rowt = sht_t.getRow(49);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-202616D4T9")) {
					rowt = sht_t.getRow(50);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-2015/2616DE4-BC")) {
					rowt = sht_t.getRow(51);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-1917D4T0")) {
					rowt = sht_t.getRow(52);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TYDA-1917D4T3")) {
					rowt = sht_t.getRow(53);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT0")) {
					rowt = sht_t.getRow(54);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT3")) {
					rowt = sht_t.getRow(55);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT6")) {
					rowt = sht_t.getRow(56);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else if (rowe.get(i).getCell(22).toString().trim().equals("TDJ-172718D-65PT9")) {
					rowt = sht_t.getRow(57);
					for (int j = 0; j < rowt.getPhysicalNumberOfCells(); j++) {
						tablevalues.add(Numberchange(rowt.getCell(j)));
					}
					tableMap.put(i, tablevalues);
				} else {
					JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的大唐（或上海贝尔）天线型号！");
					throw NullPointerException;
				}
			} else if (rowe.get(i).getCell(21).getCellType() == 0) {
				return null;
			} else {
				JOptionPane.showMessageDialog(null, "汇总表第" + (i + 1) + "行，请输入正确的天线品牌！");
				throw NullPointerException;
			}
		}
		wbt.close();
		fist.close();
		return tableMap;
	}

	public Map<Integer, String> anzhuang() throws Exception {
		FileInputStream fist = new FileInputStream(tablePath);
		HSSFWorkbook wbt = new HSSFWorkbook(fist);
		Sheet sht_t = wbt.getSheetAt(0);
		List<Row> rowe = genRowlist();

		Map<Integer, String> tableMap = new HashMap<Integer, String>();
		Row rowt = null;
		for (int i = 1; i < rowe.size(); i++) {
			String s = null;
			if (rowe.get(i).getCell(24).toString().trim().equals("挂墙机框内安装")
					|| rowe.get(i).getCell(22).toString().trim().equals("挂墙机框内安装")) {
				rowt = sht_t.getRow(59);
				s = rowt.getCell(0).toString();
				tableMap.put(i, s);
			} else if (rowe.get(i).getCell(24).toString().trim().equals("自立式机柜安装")
					|| rowe.get(i).getCell(22).toString().trim().equals("自立式机柜安装")) {
				rowt = sht_t.getRow(58);
				s = rowt.getCell(0).toString();
				tableMap.put(i, s);
			} else if (rowe.get(i).getCell(24).toString().trim().equals("嵌入综合柜安装")
					|| rowe.get(i).getCell(22).toString().trim().equals("嵌入综合柜安装")) {
				rowt = sht_t.getRow(60);
				s = rowt.getCell(0).toString();
				tableMap.put(i, s);
			}
		}
		return tableMap;
	}
	
	public String Numberchange(Cell cell){
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
		
		return value;
		
	}
}
