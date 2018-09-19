package org.easyexcel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
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
		// 创建工作簿
		Workbooks workbook = Excels.createWorkbook();
		// 创建样式建造者
		CellStyleBuilder styleBuilder = CellStyleBuilder.of(workbook);
		// 创建居中样式（内容）
		HSSFCellStyle centerStyle = styleBuilder.alignment(HorizontalAlignment.CENTER)
			.verticalAlignment(VerticalAlignment.CENTER).newStyle();
		// 创建居中加粗样式（表头）
		HSSFCellStyle centerBoldStyle = styleBuilder.bold(true).newStyle();
		workbook.createSheet()
			.mergedRegion(0, 0, 0, 3)
			.style(centerBoldStyle)
			.value("成绩表")
			.nextRow()
			.values(centerBoldStyle, "学号", "姓名", "课程", "成绩", "日期")
			.nextRow()
			.beanValues(centerStyle, getData(), "yyyy-MM-dd")
			.autoSizeColumn();
		// 写入到文件
		workbook.write(new File("example/example1.xls"));
	}

	@Test
	public void example2() throws IOException {
		Workbooks workbook = Excels.createWorkbook();
		workbook.createSheet()
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
			.autoSizeColumn();
		workbook.write(new File("example/example2.xls"));
	}

	private static class Item {
		private String no;
		private String name;
		private String course;
		private double score;
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
