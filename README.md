## 基于Apache POI封装的Excel操作API
## 添加依赖
- [easyexcel-binary-0.1.jar](https://raw.githubusercontent.com/ffpy/EasyExcel/master/downloads/easyexcel-binary-0.1.jar)
- [easyexcel-source-0.1.jar](https://raw.githubusercontent.com/ffpy/EasyExcel/master/downloads/easyexcel-source-0.1.jar)
- [poi-3.17.jar](https://raw.githubusercontent.com/ffpy/EasyExcel/master/downloads/poi-3.17.jar)

## 快速开始
### 测试数据项
```
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

    getter/setter...
}
```

### 生成数据
```
private List<Item> getData() {
    List<Item> data = new ArrayList<>();
    data.add(new Item("0001", "小明", "数学", 60, new Date()));
    data.add(new Item("0002", "小花", "数学", 59, new Date()));
    data.add(new Item("0003", "小黄", "数学", 90, new Date()));
    data.add(new Item("0004", "小红", "数学", 85, new Date()));
    return data;
}
```

### 方式一：通过ExcelHelper快速创建表格
```
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
```
通过ExcelHelper创建的表格默认会开启自适应列宽，
也可以通过autoColumnSize(false)来关闭它。

### 方式二：通过游标设置数据
```
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
```

### 生成的表格
![example](https://raw.githubusercontent.com/ffpy/EasyExcel/master/image/example.png)

## License
EasyExcel is licensed under the Apache License, Version 2.0 
