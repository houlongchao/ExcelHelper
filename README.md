# Excel 导入导出工具库

[![NuGet Badge](https://buildstats.info/nuget/ExcelHelper.Core)](https://www.nuget.org/packages/ExcelHelper.Core)
[![NuGet Badge](https://buildstats.info/nuget/ExcelHelper.NPOI)](https://www.nuget.org/packages/ExcelHelper.NPOI)
[![NuGet Badge](https://buildstats.info/nuget/ExcelHelper.Aspose)](https://www.nuget.org/packages/ExcelHelper.Aspose)
![GitHub](https://img.shields.io/github/license/houlongchao/ExcelHelper?style=social)

简单，易用，灵活的Excel导入导出工具库。支持不同Excel驱动（`NPOI`, `Aspose`），只需切换驱动包，无需修改代码。

## 功能说明

### 导入


- [x] 支持导入多个Sheet页 `.ImportSheet<DemoIO>()`
- [x] 支持导入图片 `[ImportHeader("图片", IsImage = true)]`
- [x] 支持导入配置数据限制 `[ImportLimit("A1", "A2", "A3")]`
- [x] 支持导入验证必填 `[ImportHeader("A", IsRequired = true)]`
- [x] 支持设置导入必填验证消息 `[ImportHeader("A", RequiredMessage = "数据A必填")]`
- [x] 支持导入移除前后空格 `[ImportHeader("AA", Trim = Trim.Start)]`
- [x] 支持导入数据映射 `[ImportMapper("A3", "b")]`
- [x] 支持导入数据唯一性校验 `[ImportHeader("A", IsUnique = true)]`
- [x] 支持导入组合数据唯一性校验 `[ImportUnique(nameof(A), nameof(B))]`
- [x] 支持导入时动态设置 `new ImportSetting()`

### 导出


- [x] 支持导出多个Sheet页 `.ExportSheet("sheet", data)`
- [x] 支持导出图片 `[ExportHeader("图片", IsImage = true)]`
- [x] 支持导出格式化字符串 `[ExportHeader("日期", Format = "yyyy/MM/dd")]`
- [x] 支持导出设置列宽 `[ExportHeader("日期", ColumnWidth = 30)]`
- [x] 支持导出头设置字体颜色 `[ExportHeader("A2", ColorName = "Red")]`
- [x] 支持导出数据映射 `[ExportMapper("a", "Aa")]`
- [x] 支持导出表头设置备注信息 `[ExportHeader("C2", Comment = "备注")]`
- [x] 支持导出忽略指定字段导出 `[ExportIgnore]`
- [x] 支持导出时动态设置 `new ExportSetting()`
- [x] 支持导出时设置Sheet位置 `.SetSheetIndex("sheet", 1)`

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
// 创建一个sheet页aaa，然后sheet页中依次写入数据data1，写入一个空行，写入数据data2，写入数据data3（不写入标题）
_excelHelper.CreateExcelSheet("aaa").AppendData(data1).AppendEmptyRow().AppendData(data2).AppendData(data3, false);
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
[ImportUnique(nameof(A), nameof(B))]
public class DemoIO
{
  [ImportHeader("A", IsRequired = true, IsUnique = false)]
  [ImportHeader("AA", Trim = Trim.Start)]
  [ExportHeader("A2", ColorName = "Red")]
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

  [ExportHeader("日期", ColumnWidth = 30, Format = "yyyy/MM/dd")]
  public DateTime DateTime { get; set; }

  [ExportIgnore]
  public DateTime Date { get; set; }

  [ExportMapper(0, "011")]
  [ExportHeader("数字", Format = "0.0")]
  public double Number { get; set; }

  public bool Boolean { get; set; }

  public string Formula { get; set; }

  [ExportMapper(Status.A, "AA")]
  [ExportMapper(null, "")]
  [ExportMapperElse("其它数据")]
  public Status? Status { get; set; }
  
  
  [ExportHeader("图片", IsImage = true)]
  [ImportHeader("图片", IsImage = true)]
  public byte[] Image { get; set; }
}

public enum Status
{
  A = 0,
  B = 1,
}
```

## 模型配置说明

### 导入

#### ImportHeaderAttribute

导入头设置，可以指定多个，方便兼容导入模板的改动。未配置时以属性名称作为列名称。

``` C#
[ImportHeader("A")]   // 读取列A的数据
[ImportHeader("AA")]  // 读取列AA的数据
public string A { get; set; }

[ImportHeader("图片", IsImage = true)]  // 读取图片
public byte[] Image { get; set; }       // 图片数据必须用 byte[] 接收

[ImportHeader("A", IsRequired = true)] // 数据必须不能为空
public string A { get; set; }

[ImportHeader("A", RequiredMessage = "数据A必填")] // 数据必填消息自定义
public string A { get; set; }

[ImportHeader("A", IsUnique = false)] // 数据必须不能重复
public string A { get; set; }

[ImportHeader("AA", Trim = Trim.Start)]  // 依次数据前面的空白字符
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

#### ImportLimitAttribute

导入限制，只能导入设置的数据

``` C#
[ImportLimit("A1", "A2", "A3")]   // 导入限制
public string C { get; set; }
```

#### ImportUniqueAttribute

导入唯一性数据限制，**在class上设置**

```C#
[ImportUnique(nameof(A), nameof(B))] // A和B的组合数据都唯一
public class DemoIO
{
   // ...
}
```

### 导出

#### ExportHeaderAttribute

导出头设置，可以设置列名称，列备注信息，列宽度等。未配置时以属性名称作为列名称。

``` C#
[ExportHeader("C2", Comment = "备注", IsAutoSizeColumn = true)]
public string C { get; set; }

[ExportHeader("日期", ColumnWidth = 30)]
public DateTime DateTime { get; set; }

[ExportHeader("图片", IsImage = true)]
public byte[] Image { get; set; }

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

#### ExportMapperAttribute

与导出映射转换器`ExportMapperAttribute`配合使用，当`ExportMapperAttribute`没有匹配的数据时全部数据值设置为该属性配置的值。

```C#
[ExportMapper("A3", "b")]        // 当数据为A3时Excel中写入数据b
[ExportMapperrElse("其它数据")]    // 否则其它数据都写入为"其它数据"
public string C { get; set; }
```

#### ExportIgnoreAttribute

忽略导出该字段

``` C#
[ExportIgnore]
public DateTime Date { get; set; }
```

## 动态设置

### ImportSetting

数据导入时设置的动态配置。

``` c#
var importSetting = new ImportSetting();
importSetting.AddTitleMapping(nameof(DemoIO.A), "AA");
importSetting.AddRequiredProperties(nameof(DemoIO.Image));
importSetting.AddUniqueProperties(nameof(DemoIO.A));
importSetting.AddLimitValues(nameof(DemoIO.A), "A1", "A2", "A3");
importSetting.AddValueTrim(nameof(DemoIO.A), Trim.All);

var sheets2 = _excelHelper.ImportSheet<DemoIO>(importSetting);
```

> `TitleMapping` : 导入头映射，参数为一个字典，`key`为接收模型属性名，`value`为excel表格导入列名。
>
> `RequiredProperties` : 对导入数据进行必须性验证，参数为属性名字符串列表。如果模型属性名称在列表中，则对该列数据进行必须性验证。
>
> `UniqueProperties` ：对导入数据进行唯一性验证，参数为属性名字符串列表。如果模型属性名称在列表中，则对该列数据进行唯一性验证。
>
> `LimitValues` ：对导入数据进行验证，参数为一个字典。`key`为属性名，`value`为要检查的值列表。如果导入的值在指定列表中，则该导入值有效。
>
> `ValueTrim` ：对导入数据进行前后空白字符移除，参数为一个字典。`key`为属性名，`value`为要移除的模式。

## 

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

> `AddTitle` : 导出时是否添加列标题，默认为`true`。
>
> `TitleMapping` : 导出头映射，参数为一个字典。`key`为接收模型属性名，`value`为excel表格导出列名。
>
> `IgnoreProperties` : 要忽略导出的属性，参数为属性名字符串列表。如果模型属性名称在列表中，则导出时不导出该列数据。
>
> `IncludeProperties` : 要导出的属性，参数为属性名字符串列表。如果模型属性名称在列表中，则导出时导出该列数据。优先级：`IgnoreProperties` > `IncludeProperties` > `ExportIgnoreAttribute`
>
> `TitleComment` : 导出列标题备注信息，参数为一个字典。`key`为接收模型属性名，`value`为excel表格导出列名的备注信息。