package org.ffpy.easyexcel;

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
    /** 日期格式 */
    private String dateFormat;
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
     * @param source 源样式
     * @return 复制的CellStyle实例
     */
    public static CellStyleBuilder of(CellStyleBuilder source) {
        return source.clone().clearCache();
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
     * 设置日期格式
     *
     * @param dateFormat 日期格式，如"yyyy-MM-dd"
     * @return this
     */
    public CellStyleBuilder dateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
        return this;
    }

    /**
     * 基于自身设置创建一个样式实例
     *
     * @param workbook 工作簿
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
            if (dateFormat != null) {
                short format = workbook.createDataFormat().getFormat(dateFormat);
                cellStyle.setDataFormat(format);
            }

            cellStyle.setFont(font);
        }
        return cellStyle;
    }

    /**
     * 清空缓存
     *
     * @return this
     */
    public CellStyleBuilder clearCache() {
        this.cellStyle = null;
        return this;
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

    @Override
    public String toString() {
        return "CellStyleBuilder{" +
                "horizontalAlignment=" + horizontalAlignment +
                ", verticalAlignment=" + verticalAlignment +
                ", color=" + color +
                ", bold=" + bold +
                ", italic=" + italic +
                ", underline=" + underline +
                ", dateFormat='" + dateFormat + '\'' +
                ", cellStyle=" + cellStyle +
                '}';
    }
}
