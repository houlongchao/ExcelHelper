# Excel 帮助类

- 通过对象模型进行数据的导入导出，简单易用
- 支持多种Excel驱动
- 不同Excel驱动使用相同代码，可无法切换
- 支持导入多个Sheet
- 支持导出多个Sheet



## Nuget 引用

``` sh
# 基于NPOI
dotnet add package ExcelHelper.NPOI

# 基于 Aspose
dotnet add package ExcelHelper.Aspose
```



## 使用

### 读数据

``` C#
// 通过ExcelHelperBuilder从指定文件（指定流或指定内存字节）构建读取器
_excelHelper = new ExcelHelperBuilder().BuildRead("Excel.xlsx");
_excelHelper = new ExcelHelperBuilder().BuildRead(stream);
_excelHelper = new ExcelHelperBuilder().BuildRead(bytes);
// 导入，如果没有指定Sheet则从第一个sheet读取
var demos = _excelHelper.ImportSheet<DemoIO>();
// 指定了Sheet则从指定Sheet读取
var demos = _excelHelper.ImportSheet<DemoIO>("Sheet1");
// 指定多个Sheet名称时，依次读取，只要找到对应的Sheet则读取返回，可以适用于Sheet名称修改后的兼容
var demos = _excelHelper.ImportSheet<DemoIO>("Sheet1", "S1", "S");
```

### 写数据

``` C#
// 通过ExcelHelperBuilder构建写入器
_excelHelper = new ExcelHelperBuilder().BuildWrite();
// 写入树datas到test页
_excelHelper.ExportSheet("test", datas);
// 写入树datas到test2页
_excelHelper.ExportSheet("test2", datas);
// 导出为bytes数据
var bytes = _excelHelper.ToBytes();
// 写入到文件
File.WriteAllBytes("test.xlsx", bytes);
```

### 导入导出模型

``` C#
/// <summary>
/// 导入导出测试模型
/// </summary>
public class DemoIO
{
  [ImportHeader("A")]
  [ImportHeader("AA")]
  [ExportHeader("A2")]
  public string A { get; set; }

  [ImportHeader("B")]
  [ImportHeader("BB")]
  [ImportMapper("True", "true")]
  [ExportHeader("B2")]
  public string B { get; set; }

  [ImportHeader("C")]
  [ImportHeader("CC")]
  [ImportMapper("A3", "b")]
  [ImportMapper("False", "false")]
  [ExportHeader("C2", Comment = "备注")]
  [ExportMapper("a", "Aa")]
  [ExportMapper("b", "Ab")]
  [ExportMapper("c", "Ac")]
  public string C { get; set; }

  [ExportHeader("日期", ColumnWidth = 30)]
  public DateTime DateTime { get; set; }

  [ExportIgnore]
  public DateTime Date { get; set; }

  [ExportMapper(0, "011")]
  public double Number { get; set; }

  public bool Boolean { get; set; }

  public string Formula { get; set; }

  [ExportMapper(Status.A, "AA")]
  public Status Status { get; set; }
}

public enum Status
{
  A = 0,
  B = 1,
}
```

## 属性说明

### 导入

#### ImportHeaderAttribute

导入头设置，可以指定多个，方便兼容导入模板的改动

``` C#
[ImportHeader("A")]   // 读取列A的数据
[ImportHeader("AA")]  // 读取列AA的数据
public string A { get; set; }
```

#### ImportMapperAttribute

导入映射转换器，可以将导入数据进行转换，可指定多个

``` C#
[ImportMapper("A3", "b")]           // 当Excel中数据为A3时读取后数据为b
[ImportMapper("False", "false")]    // 当Excel中数据为False时读取后为小写false
public string C { get; set; }
```

### 导出

#### ExportHeaderAttribute

导出头设置，可以设置列名称，列备注信息，列宽度等

``` C#
[ExportHeader("C2", Comment = "备注", IsAutoSizeColumn = true)]
public string C { get; set; }

[ExportHeader("日期", ColumnWidth = 30)]
public DateTime DateTime { get; set; }
```

#### ExportMapperAttribute

导出映射器，可以对数据进行转换后导出，可以指定多个

``` C#
[ExportMapper("a", "Aa")]
[ExportMapper("b", "Ab")]
[ExportMapper("c", "Ac")]
public string C { get; set; }
```

#### ExportIgnoreAttribute

忽略导出该字段

``` C#
[ExportIgnore]
public DateTime Date { get; set; }
```

