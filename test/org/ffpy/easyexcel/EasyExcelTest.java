package org.ffpy.easyexcel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class EasyExcelTest {

	private List<Item> getData() {
		List<Item> data = new ArrayList<>();
		data.add(new Item("0001", "小明", "数学", 60, new Date()));
		data.add(new Item("0002", "小花", "数学", 59, new Date()));
		data.add(new Item("0003", "小黄", "数学", 90, new Date()));
		data.add(new Item("0004", "小红", "数学", 85, new Date()));
		return data;
	}

	@Test
	public void example1() throws IOException {
		// 创建居中样式（表身）
		CellStyleBuilder centerStyle = CellStyleBuilder.of().alignment(HorizontalAlignment.CENTER)
			.verticalAlignment(VerticalAlignment.CENTER);
		// 创建居中加粗样式（表头）
		CellStyleBuilder centerBoldStyle = CellStyleBuilder.of(centerStyle).bold(true);
		Excels.createWorkbook().createSheet()
			// 合并单元格
			.mergedRegion(0, 0, 0, 4)
			// 设置标题
			.style(centerBoldStyle)
			.value("成绩表")
			// 设置表头
			.nextRow()
			.values(centerBoldStyle, "学号", "姓名", "课程", "成绩", "日期")
			// 设置表身
			.nextRow()
			.values(centerStyle, getData(), "yyyy-MM-dd")
			// 自适应列宽
			.autoColumnSize()
			// 返回工作簿
			.end()
			// 写入到文件
			.write(new File("example/example1.xls"));
	}

	@Test
	public void example2() throws IOException {
		Excels.createWorkbook().createSheet()
			.mergedRegion(0, 0, 1, 2)
			.mergedRegion(1, 2, 1, 2)
			.value("a")
			.nextCell().value("b")
			.nextCell().value("c")
			.nextRow().value("1")
			.nextCell().value("2")
			.nextCell().value("3")
			.nextRow().value("aa")
			.nextCell().value("bb")
			.nextCell().value("cc")
			.nextRow().value(new Date(), "yyyy-MM-dd")
			.nextCell().value(Calendar.getInstance(), "yyyy-MM-dd HH:mm:ss")
			.autoColumnSize()
			.end()
			.write(new File("example/example2.xls"));
	}

	@Test
	public void example3() throws IOException {
		// 创建居中样式（表身）
		CellStyleBuilder centerStyle = CellStyleBuilder.of().alignment(HorizontalAlignment.CENTER)
			.verticalAlignment(VerticalAlignment.CENTER);
		// 创建居中加粗样式（表头）
		CellStyleBuilder centerBoldStyle = CellStyleBuilder.of(centerStyle).bold(true);
		// 创建表身数据
		String[][] body = new String[2][];
		for (int i = 0; i < 2; i++) {
			body[i] = new String[3];
			for (int j = 0; j < body[i].length; j++) {
				body[i][j] = "第" + (i * body[i].length + j) + "条测试数据";
			}
		}
		// 创建表格
		Excels.helper().title(centerBoldStyle, "测试标题")
			.header(centerBoldStyle, "标题1", "标题2", "标题3")
			.body(centerStyle, body)
			.write(new File("example/example3.xls"));
	}

	@Test
	public void example4() throws IOException {
		// 创建居中样式（表身）
		CellStyleBuilder centerStyle = CellStyleBuilder.of().alignment(HorizontalAlignment.CENTER)
			.verticalAlignment(VerticalAlignment.CENTER);
		// 创建居中加粗样式（表头）
		CellStyleBuilder centerBoldStyle = CellStyleBuilder.of(centerStyle).bold(true);
		// 创建表格
		Excels.helper().globalDateFormat("yyyy-MM-dd")
			// 标题
			.title(centerBoldStyle, "成绩表")
			// 表头
			.header(centerBoldStyle, "学号", "姓名", "课程", "成绩", "日期")
			// 表身
			.body(centerStyle, getData())
			// 写入文件
			.write(new File("example/example4.xls"));
	}

	/**
	 * 测试数据项
	 */
	private static class Item {
		/** 学号 */
		private String no;
		/** 姓名 */
		private String name;
		/** 课程 */
		private String course;
		/** 成绩 */
		private double score;
		/** 考试日期 */
		private Date examTime;

		public Item(String no, String name, String course, double score, Date examTime) {
			this.no = no;
			this.name = name;
			this.course = course;
			this.score = score;
			this.examTime = examTime;
		}

		public String getNo() {
			return no;
		}

		public void setNo(String no) {
			this.no = no;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getCourse() {
			return course;
		}

		public void setCourse(String course) {
			this.course = course;
		}

		public double getScore() {
			return score;
		}

		public void setScore(double score) {
			this.score = score;
		}

		public Date getExamTime() {
			return examTime;
		}

		public void setExamTime(Date examTime) {
			this.examTime = examTime;
		}
	}
}
