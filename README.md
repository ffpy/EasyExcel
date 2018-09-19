## 基于Apache POI封装的Excel操作API
## 使用示例
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
```

### 方式二：通过游标设置数据
```
// 创建工作簿
Workbooks workbook = Excels.createWorkbook();
// 创建居中样式（表身）
CellStyleBuilder centerStyle = CellStyleBuilder.of().alignment(HorizontalAlignment.CENTER)
	.verticalAlignment(VerticalAlignment.CENTER);
// 创建居中加粗样式（表头）
CellStyleBuilder centerBoldStyle = CellStyleBuilder.of(centerStyle).bold(true);
workbook.createSheet()
	.mergedRegion(0, 0, 0, 4)
	.style(centerBoldStyle)
	.value("成绩表")
	.nextRow()
	.values(centerBoldStyle, "学号", "姓名", "课程", "成绩", "日期")
	.nextRow()
	.values(centerStyle, getData(), "yyyy-MM-dd")
	.autoColumnSize();
// 写入到文件
workbook.write(new File("example/example1.xls"));
```