package org.easyexcel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 单元格样式建造者
 */
public class CellStyleBuilder {
	private HSSFWorkbook workbook;
	private HorizontalAlignment horizontalAlignment;
	private VerticalAlignment verticalAlignment;
	private Boolean bold;
	private Short color;
	private Boolean italic;
	private Byte underline;

	/**
	 * 创建一个自身实例
	 *
	 * @param workbooks Excel工作簿辅助类
	 * @return 实例
	 */
	public static CellStyleBuilder of(Workbooks workbooks) {
		return new CellStyleBuilder(workbooks.getWorkbook());
	}

	/**
	 * 创建一个自身实例
	 *
	 * @param workbook Excel工作簿
	 * @return 实例
	 */
	private CellStyleBuilder(HSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * 设置水平对齐方式
	 *
	 * @param alignment 水平对齐方式
	 * @return 自身实例
	 */
	public CellStyleBuilder alignment(HorizontalAlignment alignment) {
		this.horizontalAlignment = alignment;
		return this;
	}

	/**
	 * 设置垂直对齐方式
	 *
	 * @param alignment 垂直对齐方式
	 * @return 自身实例
	 */
	public CellStyleBuilder verticalAlignment(VerticalAlignment alignment) {
		this.verticalAlignment = alignment;
		return this;
	}

	/**
	 * 设置粗体
	 *
	 * @param bold true为粗体，false为非粗体
	 * @return 自身实例
	 */
	public CellStyleBuilder bold(boolean bold) {
		this.bold = bold;
		return this;
	}

	/**
	 * 设置字体颜色
	 *
	 * @param color 字体颜色
	 * @return 自身实例
	 */
	public CellStyleBuilder color(short color) {
		this.color = color;
		return this;
	}

	/**
	 * 设置斜体
	 *
	 * @param italic true为斜体，false为非斜体
	 * @return 自身实例
	 */
	public CellStyleBuilder italic(boolean italic) {
		this.italic = italic;
		return this;
	}

	/**
	 * 设置下划线
	 *
	 * @param underline 下划线
	 * @return 自身实例
	 */
	public CellStyleBuilder underline(byte underline) {
		this.underline = underline;
		return this;
	}

	/**
	 * 基于自身设置创建一个样式实例
	 *
	 * @return 样式实习
	 */
	public HSSFCellStyle newStyle() {
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();

		if (horizontalAlignment != null)
			cellStyle.setAlignment(horizontalAlignment);
		if (verticalAlignment != null)
			cellStyle.setVerticalAlignment(verticalAlignment);
		if (bold != null)
			font.setBold(bold);
		if (color != null)
			font.setColor(color);
		if (italic != null)
			font.setItalic(italic);
		if (underline != null)
			font.setUnderline(underline);

		cellStyle.setFont(font);
		return cellStyle;
	}
}
