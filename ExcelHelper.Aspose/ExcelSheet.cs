using Aspose.Cells;
using ExcelHelper.Settings;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public class ExcelSheet : IExcelSheet
    {
        private readonly Worksheet _sheet;

        /// <summary>
        /// Aspose Worksheet
        /// </summary>
        public Worksheet Sheet => _sheet;

        /// <summary>
        /// Excel Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public ExcelSheet(Worksheet sheet)
        {
            _sheet = sheet;
        }

        /// <inheritdoc/>
        public IExcelSheet AppendData<T>(IEnumerable<T> datas, bool addTitle = true) where T : new()
        {
            var exportSetting = new ExportSetting();
            exportSetting.AddTitle = addTitle;

            return AppendData(datas, exportSetting);
        }

        /// <inheritdoc/>
        public IExcelSheet AppendData<T>(IEnumerable<T> datas, ExportSetting exportSetting) where T : new()
        {
            if (exportSetting == null)
            {
                exportSetting = new ExportSetting();
            }

            // 获取导出模型属性信息列表
            var excelPropertyInfoList = typeof(T).GetExportExcelPropertyInfoList(exportSetting);

            var rowIndex = _sheet.GetRowCount();

            // 设置表头信息
            if (exportSetting.AddTitle)
            {
                int colIndex = 0;
                foreach (var property in excelPropertyInfoList)
                {
                    var cell = _sheet.CreateCell(rowIndex, colIndex++);
                    cell.SetValue(property.ExportHeaderTitle);

                    var exportHeader = property.ExportHeader;
                    if (exportHeader == null)
                    {
                        exportHeader = new ExportHeaderAttribute(null);
                    }

                    cell.SetFont(font =>
                    {
                        font.Size = exportHeader.FontSize;
                        font.IsBold = exportHeader.IsBold;
                        font.Color = Color.FromName(exportHeader.ColorName);
                    });

                    if (!string.IsNullOrEmpty(exportHeader.Comment))
                    {
                        cell.SetComment(exportHeader.Comment);
                    }
                }
                rowIndex++;
            }

            // 写数据
            foreach (var data in datas)
            {
                var colIndex = 0;
                foreach (var property in excelPropertyInfoList)
                {
                    var value = property.PropertyInfo.GetValue(data);

                    // 如果导出的是图片二进制数据
                    if (property.ExportHeader != null && property.ExportHeader.IsImage)
                    {
                        if (value is byte[] imageBytes)
                        {
                            _sheet.CreateCell(rowIndex, colIndex).SetImage(imageBytes);
                        }
                        continue;
                    }

                    var displayValue = property.ExportMappedToDisplay(value);
                    var cell = _sheet.CreateCell(rowIndex, colIndex);
                    if (displayValue is DateTime dt)
                    {
                        if (DateTime.MinValue != dt)
                        {
                            cell.SetValue(dt).SetDataFormat();
                        }
                    }
                    else if (displayValue is bool b)
                    {
                        cell.SetValue(b);
                    }
                    else if (displayValue is double d)
                    {
                        cell.SetValue(d);
                    }
                    else if (displayValue is int di)
                    {
                        cell.SetValue(di);
                    }
                    else if (displayValue is decimal dc)
                    {
                        cell.SetValue((double)dc);
                    }
                    else
                    {
                        cell.SetValue(displayValue?.ToString());
                    }

                    if (!string.IsNullOrEmpty(property.ExportHeader?.Format))
                    {
                        cell.SetDataFormat(property.ExportHeader?.Format);
                    }

                    colIndex++;
                }
                rowIndex++;
            }

            // 设置列宽度
            {
                var colIndex = 0;
                foreach (var property in excelPropertyInfoList)
                {
                    var exportHeader = property.ExportHeader;
                    if (exportHeader == null)
                    {
                        exportHeader = new ExportHeaderAttribute(null);
                    }

                    if (exportHeader.IsAutoSizeColumn)
                    {
                        _sheet.AutoFitColumn(colIndex);
                    }
                    else if (exportHeader.ColumnWidth > 0)
                    {
                        _sheet.Cells.SetColumnWidth(colIndex, exportHeader.ColumnWidth);
                    }

                    colIndex++;
                }

            }

            return this;
        }

        /// <inheritdoc/>
        public IExcelSheet AppendEmptyRow()
        {
            int rowIndex = _sheet.GetRowCount();
            _sheet.CreateCell(rowIndex, 0).SetValue(null);

            return this;
        }

        /// <inheritdoc/>
        public List<T> GetData<T>() where T : new()
        {
            var result = new List<T>();

            // 获取导入模型属性信息字典
            var excelPropertyInfoNameDict = typeof(T).GetImportNamePropertyInfoDict();

            // 获取导入数据列对应的模型属性
            var excelPropertyInfoIndexDict = new Dictionary<int, ExcelPropertyInfo>();

            var columnCount = _sheet.Cells.MaxColumn + 1;
            for (int i = 0; i < columnCount; i++)
            {
                var titleCell = _sheet.GetCell(0, i);
                var title = titleCell.GetData()?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    continue;
                }
                if (!excelPropertyInfoNameDict.ContainsKey(title))
                {
                    continue;
                }
                excelPropertyInfoIndexDict[i] = excelPropertyInfoNameDict[title];
                excelPropertyInfoIndexDict[i].ImportHeaderTitle = title;
            }

            var checkDict = new Dictionary<string, HashSet<string>>();

            var rowCount = _sheet.GetRowCount();
            // 读取数据
            for (int i = 1; i < rowCount; i++)
            {
                var row = _sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                var t = new T();
                foreach (var excelPropertyInfo in excelPropertyInfoIndexDict)
                {
                    // 导入图片
                    if (excelPropertyInfo.Value.ImportIsImage())
                    {
                        var bytes = row[excelPropertyInfo.Key].GetImage();
                        excelPropertyInfo.Value.ImportCheckRequired(bytes);
                        excelPropertyInfo.Value.PropertyInfo.SetValue(t, bytes);
                        continue;
                    }

                    // 导入其它数据
                    var value = row.GetCell(excelPropertyInfo.Key).GetData();
                    excelPropertyInfo.Value.ImportCheckRequired(value);
                    excelPropertyInfo.Value.ImportTrim(ref value);
                    excelPropertyInfo.Value.ImportLimitCheckValue(value);

                    var actualValue = excelPropertyInfo.Value.ImportMappedToActual(value);

                    excelPropertyInfo.Value.PropertyInfo.SetValueAuto(t, actualValue);

                    // 唯一检查
                    if (actualValue is string strValue && excelPropertyInfo.Value.IsUnqiue())
                    {
                        if (!checkDict.ContainsKey(excelPropertyInfo.Value.PropertyInfo.Name))
                        {
                            checkDict[excelPropertyInfo.Value.PropertyInfo.Name] = new HashSet<string>();
                        }
                        if (checkDict[excelPropertyInfo.Value.PropertyInfo.Name].Contains(strValue))
                        {
                            throw new ImportException($"{excelPropertyInfo.Value.ImportHeaderTitle} 列存在重复数据");
                        }
                        checkDict[excelPropertyInfo.Value.PropertyInfo.Name].Add(strValue);
                    }
                }

                result.Add(t);
            }

            return result;
        }

        /// <inheritdoc/>
        public int GetRowCount()
        {
            return _sheet.GetRowCount();
        }
    }
}
