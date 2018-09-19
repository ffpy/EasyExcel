package org.easyexcel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 单元格样式建造者
 */
public class CellStyleBuilder implements Cloneable {
	/** 水平对齐 */
	private HorizontalAlignment horizontalAlignment;
	/** 垂直对齐 */
	private VerticalAlignment verticalAlignment;
	/** 字体颜色 */
	private Short color;
	/** 是否加粗 */
	private Boolean bold;
	/** 是否斜体 */
	private Boolean italic;
	/** 下划线样式 */
	private Byte underline;
	/** 缓存build后的样式 */
	private HSSFCellStyle cellStyle;

	/**
	 * 创建一个CellStyle实例
	 *
	 * @return CellStyle实例
	 */
	public static CellStyleBuilder of() {
		return new CellStyleBuilder();
	}

	/**
	 * 复制一个CellStyle实例
	 *
	 * @return 复制的CellStyle实例
	 */
	public static CellStyleBuilder of(CellStyleBuilder source) {
		return source.clone();
	}

	/**
	 * 设置水平对齐方式
	 *
	 * @param alignment 水平对齐方式
	 * @return this
	 */
	public CellStyleBuilder alignment(HorizontalAlignment alignment) {
		this.horizontalAlignment = alignment;
		return this;
	}

	/**
	 * 设置垂直对齐方式
	 *
	 * @param alignment 垂直对齐方式
	 * @return this
	 */
	public CellStyleBuilder verticalAlignment(VerticalAlignment alignment) {
		this.verticalAlignment = alignment;
		return this;
	}

	/**
	 * 设置粗体
	 *
	 * @param bold true为粗体，false为非粗体
	 * @return this
	 */
	public CellStyleBuilder bold(boolean bold) {
		this.bold = bold;
		return this;
	}

	/**
	 * 设置字体颜色
	 *
	 * @param color 字体颜色
	 * @return this
	 */
	public CellStyleBuilder color(short color) {
		this.color = color;
		return this;
	}

	/**
	 * 设置斜体
	 *
	 * @param italic true为斜体，false为非斜体
	 * @return this
	 */
	public CellStyleBuilder italic(boolean italic) {
		this.italic = italic;
		return this;
	}

	/**
	 * 设置下划线
	 *
	 * @param underline 下划线
	 * @return this
	 */
	public CellStyleBuilder underline(byte underline) {
		this.underline = underline;
		return this;
	}

	/**
	 * 基于自身设置创建一个样式实例
	 *
	 * @return 样式实例
	 */
	public HSSFCellStyle build(HSSFWorkbook workbook) {
		if (cellStyle == null) {
			cellStyle = workbook.createCellStyle();
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
		}
		return cellStyle;
	}

	@Override
	protected CellStyleBuilder clone() {
		try {
			return (CellStyleBuilder) super.clone();
		} catch (CloneNotSupportedException e) {
			e.printStackTrace();
		}
		return null;
	}
}
