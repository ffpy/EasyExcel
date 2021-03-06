package org.ffpy.easyexcel;

import com.sun.istack.internal.Nullable;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * Excel Sheet的辅助类
 */
public class Sheets {
    /** 工作簿 */
    private Workbooks workbooks;
    /** Sheet实例 */
    private HSSFSheet sheet;
    /** 当前行 */
    private HSSFRow curRow;
    /** 当前单元格 */
    private HSSFCell curCell;
    /** 当前行号 */
    private int curRowIndex = -1;
    /** 当前列号 */
    private int curColIndex = -1;
    /** 最大列号 */
    private int maxColNum = 0;

    /**
     * @param workbooks 工作簿
     * @param sheet     Sheet
     */
    Sheets(Workbooks workbooks, HSSFSheet sheet) {
        this.workbooks = workbooks;
        this.sheet = sheet;
        nextRow();
    }

    /**
     * 获取工作簿
     *
     * @return 工作簿
     */
    public Workbooks getWorkbooks() {
        return workbooks;
    }

    /**
     * 获取Sheet
     *
     * @return Sheet
     */
    public HSSFSheet getSheet() {
        return sheet;
    }

    /**
     * 获取当前行
     *
     * @return 当前行
     */
    public HSSFRow getCurRow() {
        return curRow;
    }

    /**
     * 获取当前单元格
     *
     * @return 当前单元格
     */
    public HSSFCell getCurCell() {
        return curCell;
    }

    /**
     * 获取最大列数
     *
     * @return 最大列数
     */
    public int getMaxColNum() {
        return maxColNum;
    }

    /**
     * 合并单元格
     *
     * @param firstRow 起始行号
     * @param lastRow  结束行号（包括）
     * @param firstCol 起始列号
     * @param lastCol  结束列号（包括）
     * @return this
     */
    public Sheets mergedRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        return this;
    }

    /**
     * 跳转到指定行和列
     *
     * @param row    行号
     * @param column 列号
     * @return this
     */
    public Sheets to(int row, int column) {
        curRow = sheet.getRow(row);
        if (curRow == null)
            curRow = sheet.createRow(row);
        curCell = curRow.getCell(column);
        if (curCell == null)
            curCell = curRow.createCell(column);
        curRowIndex = row;
        setCurColIndex(column);
        return this;
    }

    /**
     * 获取指定行
     *
     * @param rowIndex 行号
     * @return 行
     */
    public HSSFRow getRow(int rowIndex) {
        return sheet.getRow(rowIndex);
    }

    /**
     * 获取指定单元格
     *
     * @param rowIndex 行号
     * @param colIndex 列号
     * @return 单元格
     */
    public HSSFCell getCell(int rowIndex, int colIndex) {
        HSSFRow row = sheet.getRow(rowIndex);
        if (row == null) return null;
        return row.getCell(colIndex);
    }

    /**
     * 跳到下一行，同时指向第一个单元格
     *
     * @return this
     */
    public Sheets nextRow() {
        curRow = sheet.createRow(++curRowIndex);
        setCurColIndex(-1);
        nextCell();
        return this;
    }

    /**
     * 跳到下一个单元格，自动跳过合并单元格
     *
     * @return this
     */
    public Sheets nextCell() {
        setCurColIndex(curColIndex + 1);
        skipMergedRegion();
        curCell = curRow.getCell(curColIndex);
        if (curCell == null)
            curCell = curRow.createCell(curColIndex);
        return this;
    }

    /**
     * 跳过合并单元格
     * <p>如果当前位置在合并单元格的内部，并且不在合并单元格的左上角，
     * 则跳到合并单元格下一个单元格
     */
    private void skipMergedRegion() {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (curRowIndex <= mergedRegion.getLastRow() &&
                    curColIndex <= mergedRegion.getLastColumn()) {
                if ((curRowIndex == mergedRegion.getFirstRow() && curColIndex > mergedRegion.getFirstColumn()) ||
                        (curRowIndex > mergedRegion.getFirstRow() && curColIndex >= mergedRegion.getFirstColumn())) {
                    setCurColIndex(mergedRegion.getLastColumn() + 1);
                    break;
                }
            }
        }
    }

    /**
     * 跳过指定数目的单元格
     *
     * @param num 跳过的单元格数目
     * @return this
     */
    public Sheets skipCell(int num) {
        setCurColIndex(curColIndex + num);
        nextCell();
        return this;
    }

    /**
     * 设置当前单元格的样式
     *
     * @param style 样式
     * @return this
     */
    public Sheets style(CellStyleBuilder style) {
        curCell.setCellStyle(style.build(sheet.getWorkbook()));
        return this;
    }

    /**
     * 设置当前单元格的样式
     *
     * @param style 样式
     * @return this
     */
    public Sheets style(HSSFCellStyle style) {
        curCell.setCellStyle(style);
        return this;
    }

    /**
     * 设置指定区域的单元格的样式
     *
     * @param style    样式
     * @param firstRow 起始行号
     * @param lastRow  结束行号（包括）
     * @param firstCol 起始列号
     * @param lastCol  结束列号（包括）
     * @return this
     */
    public Sheets style(CellStyleBuilder style, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int r = firstRow; r <= lastRow; r++) {
            HSSFRow row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = firstCol; c <= lastCol; c++) {
                HSSFCell cell = row.getCell(c);
                if (cell == null) continue;
                cell.setCellStyle(style.build(sheet.getWorkbook()));
            }
        }
        return this;
    }

    /**
     * 设置当前单元格的值
     *
     * @param value 值
     * @return this
     */
    public Sheets value(@Nullable String value) {
        curCell.setCellValue(value);
        return this;
    }

    /**
     * 设置当前单元格的值
     *
     * @param value 值
     * @return this
     */
    public Sheets value(@Nullable RichTextString value) {
        curCell.setCellValue(value);
        return this;
    }

    /**
     * 设置当前单元格的值
     *
     * @param value 值
     * @return this
     */
    public Sheets value(double value) {
        curCell.setCellValue(value);
        return this;
    }

    /**
     * 设置当前单元格的值
     *
     * @param value 值
     * @return this
     */
    public Sheets value(Date value) {
        curCell.setCellValue(value);
        return this;
    }

    /**
     * 设置当前单元格的值
     *
     * @param value 值
     * @return this
     */
    public Sheets value(Calendar value) {
        curCell.setCellValue(value);
        return this;
    }

    /**
     * 设置当前单元格的值
     *
     * @param value 值
     * @return this
     */
    public Sheets value(boolean value) {
        curCell.setCellValue(value);
        return this;
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(String... values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(RichTextString... values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(double... values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(Date... values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(Calendar... values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(boolean... values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(String[][] values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(RichTextString[][] values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(double[][] values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(Date[][] values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(Calendar[][] values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值
     *
     * @param values 值的数组
     * @return this
     */
    public Sheets values(boolean[][] values) {
        return values(null, values);
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, String... values) {
        for (String value : values) {
            if (style != null)
                style(style);
            value(value);
            nextCell();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, RichTextString... values) {
        for (RichTextString value : values) {
            if (style != null)
                style(style);
            value(value);
            nextCell();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, double... values) {
        for (double value : values) {
            if (style != null)
                style(style);
            value(value);
            nextCell();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, Date... values) {
        for (Date value : values) {
            if (style != null)
                style(style);
            value(value);
            nextCell();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, Calendar... values) {
        for (Calendar value : values) {
            if (style != null)
                style(style);
            value(value);
            nextCell();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, boolean... values) {
        for (boolean value : values) {
            if (style != null)
                style(style);
            value(value);
            nextCell();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, String[][] values) {
        for (String[] a : values) {
            values(style, a);
            nextRow();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, RichTextString[][] values) {
        for (RichTextString[] a : values) {
            values(style, a);
            nextRow();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, double[][] values) {
        for (double[] a : values) {
            values(style, a);
            nextRow();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, Date[][] values) {
        for (Date[] a : values) {
            values(style, a);
            nextRow();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, Calendar[][] values) {
        for (Calendar[] a : values) {
            values(style, a);
            nextRow();
        }
        return this;
    }

    /**
     * 顺序设置单元格的值和样式
     *
     * @param style  单元格的样式
     * @param values 值的数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, boolean[][] values) {
        for (boolean[] a : values) {
            values(style, a);
            nextRow();
        }
        return this;
    }

    /**
     * 按照Bean的字段的顺序设置单元格的值
     *
     * @param values bean数组
     * @return this
     */
    public Sheets values(List<?> values) {
        return values(null, values);
    }

    /**
     * 按照Bean的字段的顺序设置单元格的值和样式
     *
     * @param style      样式
     * @param values     bean数组
     * @return this
     */
    public Sheets values(@Nullable CellStyleBuilder style, List<?> values) {
        if (values == null || values.isEmpty()) return this;

        CellStyleBuilder dateStyle = style;
        if (style != null) {
            style = CellStyleBuilder.of(style).dateFormat(null);
        }


        Class<?> beanClass = values.get(0).getClass();
        List<PropertyHelper> propertyHelpers = BeanHelper.of(beanClass).getOrderedProperties();

        for (Object o : values) {
            for (PropertyHelper propertyHelper : propertyHelpers) {
                if (style != null) {
                    style(style);
                }

                if (propertyHelper.getPropertyType() == String.class) {
                    String value = propertyHelper.getProperty(o);
                    value(value);
                } else if (propertyHelper.getPropertyType() == RichTextString.class) {
                    RichTextString value = propertyHelper.getProperty(o);
                    value(value);
                } else if (propertyHelper.getPropertyType() == double.class) {
                    double value = propertyHelper.getProperty(o);
                    value(value);
                } else if (propertyHelper.getPropertyType() == Date.class) {
                    Date value = propertyHelper.getProperty(o);
                    if (dateStyle != null) {
                        style(dateStyle);
                    }
                    value(value);
                } else if (propertyHelper.getPropertyType() == Calendar.class) {
                    Calendar value = propertyHelper.getProperty(o);
                    if (dateStyle != null) {
                        style(dateStyle);
                    }
                    value(value);
                } else if (propertyHelper.getPropertyType() == boolean.class) {
                    boolean value = propertyHelper.getProperty(o);
                    value(value);
                } else {
                    throw new RuntimeException("不支持的字段类型：" + propertyHelper.getPropertyType().getName());
                }

                nextCell();
            }

            nextRow();
        }

        return this;
    }

    /**
     * 自动调整列宽（支持中文）
     *
     * @return this
     */
    public Sheets autoColumnSize() {
        autoColumnSize(0, maxColNum);
        return this;
    }

    /**
     * 自动调整列宽（支持中文）
     * <p>参考：https://blog.csdn.net/jeikerxiao/article/details/80702543
     *
     * @param firstColumn 起始列
     * @param lastColumn  结束列（包含）
     * @return this
     */
    public Sheets autoColumnSize(int firstColumn, int lastColumn) {
        for (int columnNum = firstColumn; columnNum <= lastColumn; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                HSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }

                if (currentRow.getCell(columnNum) != null) {
                    HSSFCell currentCell = currentRow.getCell(columnNum);
                    int length = -1;
                    switch (currentCell.getCellTypeEnum()) {
                        // 字符串
                        case STRING:
                            length = currentCell.getStringCellValue().getBytes().length;
                            break;
                        // 日期
                        case NUMERIC:
                            if (DateUtil.isValidExcelDate(currentCell.getNumericCellValue())) {
                                Date value = currentCell.getDateCellValue();
                                if (value != null) {
                                    String pattern = currentCell.getCellStyle().getDataFormatString();
                                    if (pattern != null && !pattern.isEmpty() && !"General".equals(pattern)) {
                                        SimpleDateFormat dateFormat = new SimpleDateFormat(pattern);
                                        length = dateFormat.format(value).getBytes().length;
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                    if (columnWidth < length) {
                        columnWidth = length;
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
        return this;
    }

    /**
     * 返回到工作簿
     *
     * @return 工作簿
     */
    public Workbooks end() {
        return workbooks;
    }

    /**
     * 设置当前列号，并记录最大列号
     *
     * @param col 列号
     */
    private void setCurColIndex(int col) {
        this.curColIndex = col;
        if (col > maxColNum) {
            maxColNum = col;
        }
    }
}
