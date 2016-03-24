package com.cm.oe.ui;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

public class ExcelRRU {
	private static String EXCEL_HIDE_SHEET_NAME = "poihide";
	private static String HIDE_SHEET_NAME_YN = "yesOrNOList";
	private static String HIDE_SHEET_NAME_PROVINCE = "provinceList";
	// 设置下拉列表的内容
	private static String[] yesOrNOList = { "是", "否" };
	private static String[] provinceList = { "华为", "中兴", "大唐", "上海贝尔" };
	private static String[] gzProvinceList = { "AAU3213", "RRU3277", "DRRU3168e-fa", "RRU3278M","DRRU3172-fad","DRRU3161-fae","BTS3205E","BBOK RRU","Easymacro" };
	private static String[] hnProvinceList = { "ZXSDR R8978 S2600W", "ZXSDR R8972E S2600W", "ZXSDR R8978 M1920A","ZXSDR R8972E M1920A","ZXSDR R8972E S2300W","ZXSDR R8972E M192023A" };
	private static String[] ssProvinceList = { "TDRU348FA", "TDRU348D", "TDRU342D","TDRU342FA","TDRU342E","TDRU341FAE","mTDRU342D" };
	private static String[] xxProvinceList = { "TD-RRH8X20-25A", "TD-RRH2×40-25A", "TD-RRH8x10-1935","TD-RRH2x60-1935","TD-RRH2X50-2350","9768 MRO B38 TD-LTE 2x5W"};

	public static void creatExcelHidePage(Workbook workbook) {
		Sheet hideInfoSheet = workbook.createSheet(EXCEL_HIDE_SHEET_NAME);// 隐藏一些信息
		// 在隐藏页设置选择信息
		// 第一行设置性别信息
		Row sexRow = hideInfoSheet.createRow(0);
		creatRow(sexRow, yesOrNOList);
		// 第二行设置省份名称列表
		Row provinceNameRow = hideInfoSheet.createRow(1);
		creatRow(provinceNameRow, provinceList);
		// 以下行设置城市名称列表
		Row cityNameRow = hideInfoSheet.createRow(2);
		creatRow(cityNameRow, gzProvinceList);

		cityNameRow = hideInfoSheet.createRow(3);
		creatRow(cityNameRow, hnProvinceList);

		cityNameRow = hideInfoSheet.createRow(4);
		creatRow(cityNameRow, ssProvinceList);

		cityNameRow = hideInfoSheet.createRow(5);
		creatRow(cityNameRow, xxProvinceList);
		// 名称管理
		// 第一行设置性别信息
		creatExcelNameList(workbook, HIDE_SHEET_NAME_YN, 1, yesOrNOList.length,
				false);
		// 第二行设置省份名称列表
		creatExcelNameList(workbook, HIDE_SHEET_NAME_PROVINCE, 2,
				provinceList.length, false);
		// 以后动态大小设置省份对应的城市列表
		creatExcelNameList(workbook, provinceList[0], 3, gzProvinceList.length,
				true);
		creatExcelNameList(workbook, provinceList[1], 4, hnProvinceList.length,
				true);

		creatExcelNameList(workbook, provinceList[2], 5, ssProvinceList.length,
				true);
		creatExcelNameList(workbook, provinceList[3], 6, xxProvinceList.length,
				true);
		// 设置隐藏页标志
		workbook.setSheetHidden(workbook.getSheetIndex(EXCEL_HIDE_SHEET_NAME),
				true);
	}

	private static void creatExcelNameList(Workbook workbook, String nameCode,
			int order, int size, boolean cascadeFlag) {
		Name name;
		name = workbook.createName();
		name.setNameName(nameCode);
		/*System.out.println(nameCode+"bbbbbbbbbbbbbbbbbbbbbbbbbbbbb");*/
		String formula = EXCEL_HIDE_SHEET_NAME + "!"
				+ creatExcelNameList(order, size, cascadeFlag);
		/*System.out.println(nameCode + " ==  " + formula);*/
		name.setRefersToFormula(formula);
	}

	private static String creatExcelNameList(int order, int size,
			boolean cascadeFlag) {
		char start = 'A';
		if (cascadeFlag) {
			if (size <= 25) {
				char end = (char) (start + size - 1);
				return "$" + start + "$" + order + ":$" + end + "$" + order;
			} else {
				char endPrefix = 'A';
				char endSuffix = 'A';
				if ((size - 25) / 26 == 0 || size == 51) {// 26-51之间，包括边界（仅两次字母表计算）
					if ((size - 25) % 26 == 0) {// 边界值
						endSuffix = (char) ('A' + 25);
					} else {
						endSuffix = (char) ('A' + (size - 25) % 26 - 1);
					}
				} else {// 51以上
					if ((size - 25) % 26 == 0) {
						endSuffix = (char) ('A' + 25);
						endPrefix = (char) (endPrefix + (size - 25) / 26 - 1);
					} else {
						endSuffix = (char) ('A' + (size - 25) % 26 - 1);
						endPrefix = (char) (endPrefix + (size - 25) / 26);
					}
				}
				return "$" + start + "$" + order + ":$" + endPrefix + endSuffix
						+ "$" + order;
			}
		} else {
			if (size <= 26) {
				char end = (char) (start + size - 1);
				return "$" + start + "$" + order + ":$" + end + "$" + order;
			} else {
				char endPrefix = 'A';
				char endSuffix = 'A';
				if (size % 26 == 0) {
					endSuffix = (char) ('A' + 25);
					if (size > 52 && size / 26 > 0) {
						endPrefix = (char) (endPrefix + size / 26 - 2);
					}
				} else {
					endSuffix = (char) ('A' + size % 26 - 1);
					if (size > 52 && size / 26 > 0) {
						endPrefix = (char) (endPrefix + size / 26 - 1);
					}
				}
				return "$" + start + "$" + order + ":$" + endPrefix + endSuffix
						+ "$" + order;
			}
		}
	}

	private static void creatRow(Row currentRow, String[] textList) {
		if (textList != null && textList.length > 0) {
			int i = 0;
			for (String cellValue : textList) {
				Cell userNameLableCell = currentRow.createCell(i++);
				userNameLableCell.setCellValue(cellValue);
			}
		}
	}

	public static void setDataValidation(Workbook wb) {
		int sheetIndex = wb.getNumberOfSheets();
		if (sheetIndex > 0) {
			for (int i = 0; i < sheetIndex; i++) {
				Sheet sheet = wb.getSheetAt(i);
				if (!EXCEL_HIDE_SHEET_NAME.equals(sheet.getSheetName())) {
					// 省份选项添加验证数据
					for (int a = 2; a < 1000; a++) {
						// 性别添加验证数据
						/*
						 * DataValidation data_validation_list =
						 * getDataValidationByFormula(HIDE_SHEET_NAME_YN, a, 1);
						 * sheet.addValidationData(data_validation_list);
						 */
						DataValidation data_validation_list2 = getDataValidationByFormula(
								HIDE_SHEET_NAME_PROVINCE, a, 20);
						sheet.addValidationData(data_validation_list2);
						// 城市选项添加验证数据
						DataValidation data_validation_list3 = getDataValidationByFormula(
								"INDIRECT($T$" + a + ")", a, 21);
						sheet.addValidationData(data_validation_list3);
					}
				}
			}
		}
	}

	private static DataValidation getDataValidationByFormula(
			String formulaString, int naturalRowIndex, int naturalColumnIndex) {
		/*System.out.println("formulaString  " + formulaString);*/
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint
				.createFormulaListConstraint(formulaString);
		// 设置数据有效性加载在哪个单元格上。
		// 四个参数分别是：起始行、终止行、起始列、终止列
		int firstRow = naturalRowIndex - 1;
		int lastRow = naturalRowIndex - 1;
		int firstCol = naturalColumnIndex - 1;
		int lastCol = naturalColumnIndex - 1;
		CellRangeAddressList regions = new CellRangeAddressList(firstRow,
				lastRow, firstCol, lastCol);
		// 数据有效性对象
		DataValidation data_validation_list = new HSSFDataValidation(regions,
				constraint);
		data_validation_list.setEmptyCellAllowed(false);
		if (data_validation_list instanceof XSSFDataValidation) {
			data_validation_list.setSuppressDropDownArrow(true);
			data_validation_list.setShowErrorBox(true);
		} else {
			data_validation_list.setSuppressDropDownArrow(false);
		}
		// 设置输入信息提示信息
		data_validation_list.createPromptBox("下拉选择提示", "请使用下拉方式选择合适的值！");
		// 设置输入错误提示信息
		data_validation_list
				.createErrorBox("选择错误提示", "你输入的值未在备选列表中，请下拉选择合适的值！");
		return data_validation_list;
	}
}