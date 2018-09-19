package org.ffpy.easyexcel;

import com.sun.istack.internal.Nullable;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * Excel表格辅助者
 */
public class ExcelHelper {
	/** 工作簿 */
	private final Workbooks workbook;
	/** Sheet */
	private final Sheets sheet;
	/** 标题 */
	private String title;
	/** 标题样式 */
	private CellStyleBuilder titleStyle;
	/** 自适应列宽，默认开启 */
	private boolean autoColumnSize = true;

	ExcelHelper() {
		workbook = Excels.createWorkbook();
		sheet = workbook.createSheet();
	}

	/**
	 * @param sheetname Sheet名称
	 */
	ExcelHelper(String sheetname) {
		workbook = Excels.createWorkbook();
		sheet = workbook.createSheet(sheetname);
	}

	/**
	 * 获取工作簿
	 *
	 * @return 工作簿
	 */
	public Workbooks getWorkbook() {
		return workbook;
	}

	/**
	 * 获取Sheet
	 *
	 * @return Sheet
	 */
	public Sheets getSheet() {
		return sheet;
	}

	/**
	 * 是否自适应列宽
	 *
	 * @param autoColumnSize 是否自适应列宽
	 * @return this
	 */
	public ExcelHelper autoColumnSize(boolean autoColumnSize) {
		this.autoColumnSize = autoColumnSize;
		return this;
	}

	/**
	 * 设置标题
	 *
	 * @param title 标题
	 * @return this
	 */
	public ExcelHelper title(String title) {
		return title(null, title);
	}

	/**
	 * 设置标题
	 *
	 * @param style 样式
	 * @param title 标题
	 * @return this
	 */
	public ExcelHelper title(@Nullable CellStyleBuilder style, String title) {
		this.title = title;
		this.titleStyle = style;
		sheet.nextRow();
		return this;
	}

	/**
	 * 设置全局日期格式
	 *
	 * @param format 日期格式
	 * @return this
	 */
	public ExcelHelper globalDateFormat(String format) {
		sheet.globalDateFormat(format);
		return this;
	}

	/**
	 * 设置表头
	 *
	 * @param headers 表头数组
	 * @return this
	 */
	public ExcelHelper header(String... headers) {
		return header(null, headers);
	}

	/**
	 * 设置表头
	 *
	 * @param style 样式
	 * @param headers 表头数组
	 * @return this
	 */
	public ExcelHelper header(@Nullable CellStyleBuilder style, String... headers) {
		sheet.values(style, headers).nextRow();
		return this;
	}

	/**
	 * 设置表身
	 *
	 * @param body 表身
	 * @return this
	 */
	public ExcelHelper body(String[][] body) {
		return body(null, body);
	}

	/**
	 * 设置表身
	 *
	 * @param style 样式
	 * @param body 表身
	 * @return this
	 */
	public ExcelHelper body(@Nullable CellStyleBuilder style, String[][] body) {
		sheet.values(style, body).nextRow();
		return this;
	}

	/**
	 * 设置表身
	 *
	 * @param style 样式
	 * @param body 表身
	 * @return this
	 */
	public ExcelHelper body(@Nullable CellStyleBuilder style, List<?> body) {
		sheet.values(style, body).nextRow();
		return this;
	}

	/**
	 * 输出到文件
	 *
	 * @param file 输出的文件
	 * @throws IOException IO错误
	 */
	public void write(File file) throws IOException {
		beforeWrite();
		workbook.write(file);
	}

	/**
	 * 输出到输出流
	 *
	 * @param out 输出流
	 * @throws IOException IO错误
	 */
	public void write(OutputStream out) throws IOException {
		beforeWrite();
		workbook.write(out);
	}

	/**
	 * 输出前的处理
	 */
	private void beforeWrite() {
		// 写入标题
		if (title != null) {
			sheet.mergedRegion(0, 0, 0, sheet.getMaxColNum() - 1)
				.to(0, 0);
			if (titleStyle != null)
				sheet.style(titleStyle);
			sheet.value(title);
		}
		// 自适应列宽
		if (autoColumnSize) {
			sheet.autoColumnSize();
		}
	}
}
