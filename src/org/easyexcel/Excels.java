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
}
