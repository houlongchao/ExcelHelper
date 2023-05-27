# Excel 导入导出工具库

[![NuGet Badge](https://buildstats.info/nuget/ExcelHelper.Core)](https://www.nuget.org/packages/ExcelHelper.Core)
[![NuGet Badge](https://buildstats.info/nuget/ExcelHelper.NPOI)](https://www.nuget.org/packages/ExcelHelper.NPOI)
[![NuGet Badge](https://buildstats.info/nuget/ExcelHelper.Aspose)](https://www.nuget.org/packages/ExcelHelper.Aspose)
![GitHub](https://img.shields.io/github/license/houlongchao/ExcelHelper?style=social)

简单，易用，灵活的Excel导入导出工具库。支持不同Excel驱动（`NPOI`, `Aspose`），只需切换驱动包，无需修改代码。

## 功能说明

### 数据导入


- [x] 支持导入多个Sheet页 `.ImportSheet<DemoIO>()`
- [x] 支持数据列标题设置 `[ImportHeader("姓名")]`
- [x] 支持导入配置数据限制 `[ImportLimit("A1", "A2", "A3")]`
- [x] 支持导入验证必填 `[ImportRequired]`
- [x] 支持设置导入必填验证消息 `[ImportRequired(Message = "数据A必填")]`
- [x] 支持导入移除前后空格 `[ImportTrim(Trim.Start)]`
- [x] 支持导入数据映射 `[ImportMapper("A3", "b")]`
- [x] 支持导入数据唯一性校验 `[ImportUnique]`
- [x] 支持导入组合数据唯一性校验 `[ImportUniques(nameof(A), nameof(B))]`
- [x] 支持导入时动态设置 `new ImportSetting()`

### 数据导出


- [x] 支持导出多个Sheet页 `.ExportSheet("sheet", data)`
- [x] 支持数据列标题设置 `[ExportHeader("日期"]`
- [x] 支持导出格式化字符串 `[ExportFormat("yyyy/MM/dd")]`
- [x] 支持导出设置列宽 `[ExportHeader("日期", ColumnWidth = 30)]`
- [x] 支持导出头设置字体颜色 `[ExportHeader("A2", ColorName = "Red")]`
- [x] 支持导出数据映射 `[ExportMapper("a", "Aa")]`
- [x] 支持导出表头设置备注信息 `[ExportHeader("C2", Comment = "备注")]`
- [x] 支持导出忽略指定字段导出 `[ExportIgnore]`
- [x] 支持导出时动态设置 `new ExportSetting()`
- [x] 支持导出时设置Sheet位置 `.SetSheetIndex("sheet", 1)`

### 模板操作

- [x] 标识数据属性对应Excel位置 `[Temp("A1")]`
- [x] 标识列表数据行/列范围`[TempList(TempListType.Row, 5, 8)]`
- [x] 标识列表数据位置`[TempListItem(1)]`
- [x] 模板导入导出时支持Import和Export操作的限制型属性

### 公共

- [x] 支持导入导出图片 `[Image]`



## Nuget 引用

``` sh
# 基于NPOI
dotnet add package ExcelHelper.NPOI

# 基于 Aspose
dotnet add package ExcelHelper.Aspose
```

## 基本使用

### 读数据

``` C#
// 构建读取器
_excelHelper = new ExcelReadHelper("Excel.xlsx");
_excelHelper = new ExcelReadHelper(stream);
_excelHelper = new ExcelReadHelper(bytes);
// 导入，如果没有指定Sheet则从第一个sheet读取
var demos = _excelHelper.ImportSheet<DemoIO>();
// 指定了Sheet则从指定Sheet读取
var demos = _excelHelper.ImportSheet<DemoIO>("Sheet1");
// 指定多个Sheet名称时，依次读取，只要找到对应的Sheet则读取返回，可以适用于Sheet名称修改后的兼容
var demos = _excelHelper.ImportSheet<DemoIO>("Sheet1", "S1", "S");
```

### 写数据

``` C#
// 构建写入器
_excelHelper = new ExcelWriteHelper();
// 写入树datas到test页
_excelHelper.ExportSheet("test", datas);
// 写入树datas到test2页
_excelHelper.ExportSheet("test2", datas);
// 创建一个sheet页aaa，然后sheet页中依次写入数据data1，写入一个空行，写入数据data2，写入数据data3（不写入标题）
_excelHelper.CreateExcelSheet("aaa").AppendData(data1).AppendEmptyRow().AppendData(data2).AppendData(data3, false);
// 导出为bytes数据
var bytes = _excelHelper.ToBytes();
// 写入到文件
File.WriteAllBytes("test.xlsx", bytes);
```

### 模板操作

``` C#
// 构建模板操作类
_excelHelper = new ExcelTempHelper();
// 将数据 tempIO 写到模板 Excel.xlsx 中
var bytes = _excelHelper.SetData("Excel.xlsx", tempIo);
// 将数据写入到文件
File.WriteAllBytes("test.xlsx", bytes);

// 从包含数据的模板 test.xlsx 中读取数据
var data = _excelHelper.GetData<DemoTempIO>("test.xlsx");
```

### 模型

#### 导入导出

``` C#
/// <summary>
/// 导入导出测试模型
/// </summary>
[ImportUniques(nameof(A), nameof(B))]
//[ImportUniques(nameof(A), nameof(B), Message = "数据必须唯一提示")]
public class DemoIO
{
  [ImportHeader("A")]
  [ImportHeader("AA")]
  [ImportRequired]
  //[ImportRequired(Message = "数据必填提示")]
  [ImportUnique]
  //[ImportUnique(Message = "数据唯一提示")]
  [ImportTrim(Trim.Start)]
  [ImportLimit("A1", "A2", "A3")]
  [ExportHeader("A2", ColorName = "Red")]
  public string A { get; set; }

  [ImportHeader("B")]
  [ImportHeader("BB")]
  [ImportRequired(Message = "数据B必填")]
  [ExportHeader("B2")]
  public string B { get; set; }

  [ImportHeader("C")]
  [ImportHeader("CC")]
  [ImportMapper("A3", "b")]
  [ImportLimit("A3", true, 123)]
  [ExportHeader("C2", Comment = "备注")]
  [ExportMapper("a", "Aa")]
  [ExportMapper("b", "Ab")]
  [ExportMapper("c", "Ac")]
  public string C { get; set; }

  [ExportHeader("日期", ColumnWidth = 30)]
  public DateTime DateTime { get; set; }

  [ExportHeader("日期2", ColumnWidth = 30)]
  [ExportFormat("yyyy/MM/dd")]
  public DateTime? DateTime2 { get; set; }

  [ExportIgnore]
  public DateTime Date { get; set; }

  [ExportHeader("数字")]
  [ExportFormat("0.0")]
  public double Number { get; set; }

  public bool Boolean { get; set; }

  public string Formula { get; set; }

  [ExportMapper(ExcelHelperTest.Status.A, "AA")]
  [ExportMapper(null, "")]
  [ExportMapperElse("else")]
  public Status? Status { get; set; }

  public string ImageName { get; set; }

  [ExportHeader("图片")]
  [ImportHeader("图片")]
  [Image]
  public byte[] Image { get; set; }
}

public enum Status
{
  A = 0,
  B = 1,
}
```

#### 模板

``` C#
public class DemoTempIO
{
  [Temp("A1")]
  public string A { get; set; }

  [Temp("B2")]
  public int B { get; set; }

  [Temp("C3")]
  public DateTime C { get; set; }

  public string D { get; set; }

  [TempList(TempListType.Row, 5, 8)]
  public List<DemoTempChild> Children { get; set; }
}

public class DemoTempChild
{
  [TempListItem(1)]
  public string Name { get; set; }

  [TempListItem(2)]
  public int Age { get; set; }

  public string Other { get; set; }
}
```



## 模型配置说明

### 公共Attribute

#### ImageAttribute

设置在`byte[]`数组属性上，标识该属性为图片的二进制数据。

``` C#
[Image]
public byte[] Image { get; set; }  // 图片数据必须用 byte[] 接收
```



### 导入Attribute

#### ImportHeaderAttribute

导入头设置，可以指定多个，方便兼容导入模板的改动。未配置时以属性名称作为列名称。

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

#### ImportMapperElseAttribute

与导入映射转换器`ImportMapperAttribute`配合使用，当`ImportMapperAttribute`没有匹配的数据时全部数据值设置为该属性配置的值。

```C#
[ImportMapper("A3", "b")]        // 当Excel中数据为A3时读取后数据为b
[ImportMapperElse("其它数据")]    // 否则其它数据都读取为"其它数据"
public string C { get; set; }
```

#### ImportTrimAttribute

设置导入数据时对数据的Trim操作方式

``` C#
[ImportTrim(Trim.Start)]
public string A { get; set; }   //移除数据前面的空白字符
```

#### ImportRequiredAttribute

导入数据必填不能为空

``` C#
[ImportRequired]
public string A { get; set; }

[ImportRequired(Message = "数据必填提示")]
public string A { get; set; }
```

#### ImportLimitAttribute

导入限制，只能导入设置的数据

``` C#
[ImportLimit("A1", "A2", "A3")]   // 导入限制
public string A { get; set; }

[ImportLimit("A1", "A2", "A3", Message = "数据限制提示")] // 导入限制
public string A { get; set; }
```

#### ImportUniqueAttribute

导入唯一性数据限制

```C#
[ImportRequired]
public string A { get; set; }

[ImportRequired(Message = "数据必填提示")]
public string A { get; set; }
```

#### ImportUniquesAttribute

导入唯一性数据限制，**在class上设置**

``` C#
[ImportUnique(nameof(A), nameof(B))] // A和B的组合数据都唯一
public class DemoIO
{
   // ...
}

[ImportUniques(nameof(A), nameof(B), Message = "数据必须唯一提示")] // A和B的组合数据都唯一
public class DemoIO
{
   // ...
}
```



### 导出Attribute

#### ExportHeaderAttribute

导出头设置，可以设置列名称，列备注信息，列宽度等。未配置时以属性名称作为列名称。

``` C#
[ExportHeader("C2", Comment = "备注", IsAutoSizeColumn = true)]
public string C { get; set; }

[ExportHeader("日期", ColumnWidth = 30)]
public DateTime DateTime { get; set; }

[ExportHeader("A2", ColorName = "Red", IsBold = true, FontSize = 12)] // 指定导出标题字体
public string A { get; set; }
```

#### ExportMapperAttribute

导出映射器，可以对数据进行转换后导出，可以指定多个

``` C#
[ExportMapper("a", "Aa")]
[ExportMapper("b", "Ab")]
[ExportMapper("c", "Ac")]
public string C { get; set; }
```

#### ExportMapperElseAttribute

与导出映射转换器`ExportMapperAttribute`配合使用，当`ExportMapperAttribute`没有匹配的数据时全部数据值设置为该属性配置的值。

```C#
[ExportMapper("A3", "b")]        // 当数据为A3时Excel中写入数据b
[ExportMapperElse("其它数据")]    // 否则其它数据都写入为"其它数据"
public string C { get; set; }
```

#### ExportFormatAttribute

``` C#
[ExportFormat("yyyy/MM/dd")]
public DateTime? DateTime2 { get; set; }

[ExportFormat("0.0")]
public double Number { get; set; }
```

#### ExportIgnoreAttribute

忽略导出该字段

``` C#
[ExportIgnore]
public DateTime Date { get; set; }
```

### 模板Attribute

####  TempAttribute

设置字段属性与Excel单元格的关系

``` C#
[Temp("A1")]
public string A { get; set; }   // 设置A数据和Excel中A1单元格数据绑定
```

#### TempListAttribute

设置列表属性对应Excel中表格的位置关系。这是表格的行/列模式，行/列的开始和结束索引。

``` C#
[TempList(TempListType.Row, 5, 8)]         // 数据为行模式，数据从索引5(第6行，包含)到缩影8(第9行，包含)
public List<DemoTempChild> Children { get; set; }
```

#### TempListItemAttribute

设置列表数据中数据属性的行/列索引。如果列表为行模式，则此处为列索引。

```C#
[TempListItem(1)]
public string Name { get; set; }
```


## 动态设置

### ImportSetting

数据导入时设置的动态配置。

``` c#
var importSetting = new ImportSetting();
importSetting.AddTitleMapping(nameof(DemoIO.A), "AA");
importSetting.AddRequiredProperties(nameof(DemoIO.A));
importSetting.AddRequiredMessage(nameof(DemoIO.A), "AA是必须的");
importSetting.AddUniqueProperties(nameof(DemoIO.A));
importSetting.AddUniqueMessage(nameof(DemoIO.A), "AA必须唯一");
importSetting.AddLimitValues(nameof(DemoIO.A), "A1", "A2", "A3");
importSetting.AddLimitMessage(nameof(DemoIO.A), "AA数据非法");
importSetting.AddValueTrim(nameof(DemoIO.A), Trim.All);

var sheets2 = _excelHelper.ImportSheet<DemoIO>(importSetting);
```

### ExportSetting

数据导出时设置的动态配置

``` C#
var exportSetting = new ExportSetting()
setting.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));
setting.AddIncludeProperties(nameof(DemoIO.Date), nameof(DemoIO.B));
setting.AddTitleMapping(nameof(DemoIO.Date), "日期");
setting.AddTitleComment(nameof(DemoIO.Date), "日期备注");

_excelHelper.ExportSheet("test3", data3, exportSetting);
```

### TempSetting

数据模板导入导出动态设置

``` C#
var tempSetting = new TempSetting();
tempSetting.AddCellAddress(nameof(DemoTempIO.A), "A8");
tempSetting.AddCellAddress(nameof(DemoTempIO.B), "B8");
tempSetting.AddCellAddress(nameof(DemoTempIO.C), "C8");
tempSetting.AddCellAddress(nameof(DemoTempIO.D), "D8");
//tempSetting.AddRequiredProperties(nameof(DemoTempIO.A));
//tempSetting.AddRequiredMessage(nameof(DemoTempIO.A), "AA是必须的");
//tempSetting.AddUniqueProperties(nameof(DemoTempIO.A));
//tempSetting.AddUniqueMessage(nameof(DemoTempIO.A), "AA必须唯一");
//tempSetting.AddLimitValues(nameof(DemoTempIO.A), "A1", "A2", "A3");
//tempSetting.AddLimitMessage(nameof(DemoTempIO.A), "AA数据非法");
//tempSetting.AddValueTrim(nameof(DemoTempIO.A), Trim.All);
var childrenSetting = tempSetting.AddTempListSetting(nameof(DemoTempIO.Children), TempListType.Row, 10, 15);
childrenSetting.AddItemIndex(nameof(DemoTempChild.Name), 0);
childrenSetting.AddItemIndex(nameof(DemoTempChild.Age), 5);
childrenSetting.AddItemIndex(nameof(DemoTempChild.Other), 3);

var bytes = _excelHelper.SetData("Excel.xlsx", tempIo, tempSetting: tempSetting);
File.WriteAllBytes("test.xlsx", bytes);

var data = _excelHelper.GetData<DemoTempIO>("test.xlsx", tempSetting: tempSetting);
```

