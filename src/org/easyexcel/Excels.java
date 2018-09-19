package org.easyexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 封装了Apache POI操作Excel辅助类
 */
public class Excels {

	/**
	 * 创建一个工作簿实例
	 *
	 * @return 工作簿实例
	 */
	public static Workbooks createWorkbook() {
		return new Workbooks(new HSSFWorkbook());
	}

	/**
	 * 创建一个Excel表格辅助者
	 *
	 * @return Excel表格辅助者
	 */
	public static ExcelHelper helper() {
		return new ExcelHelper();
	}

	/**
	 * 创建一个Excel表格辅助者
	 *
	 * @param sheetname Sheet名称
	 * @return Excel表格辅助者
	 */
	public static ExcelHelper helper(String sheetname) {
		return new ExcelHelper(sheetname);
	}
}
