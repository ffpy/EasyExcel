package org.ffpy.easyexcel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Excel工作簿的辅助类
 */
public class Workbooks {
	/** 工作簿 */
	private HSSFWorkbook workbook;

	/**
	 * 获取工作簿
	 *
	 * @return 工作簿
	 */
	public HSSFWorkbook getWorkbook() {
		return workbook;
	}

	/**
	 * @param workbook 工作簿
	 */
	Workbooks(HSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * 创建工作簿，设置默认名字
	 *
	 * @return 工作簿辅助类
	 */
	public Sheets createSheet() {
		return new Sheets(workbook.createSheet());
	}

	/**
	 * 创建工作簿，设置指定名字
	 *
	 * @param sheetname 工作簿名称
	 * @return 工作簿辅助类
	 */
	public Sheets createSheet(String sheetname) {
		return new Sheets(workbook.createSheet(sheetname));
	}

	/**
	 * 写入到输出流中
	 *
	 * @param out 输出流
	 * @throws IOException IO错误   
	 */
	public void write(OutputStream out) throws IOException {
		workbook.write(out);
	}

	/**
	 * 写入到文件
	 *
	 * @param file 写入的文件
	 * @throws IOException IO错误
	 */
	public void write(File file) throws IOException {
		workbook.write(file);
	}
}
