using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public class ExcelSheet : IExcelSheet
    {
        private readonly ISheet _sheet;

        /// <summary>
        /// NPOI ISheet
        /// </summary>
        public ISheet Sheet => _sheet;

        /// <summary>
        /// Excel Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public ExcelSheet(ISheet sheet)
        {
            _sheet = sheet;
        }

        /// <inheritdoc/>
        public IExcelSheet AppendData<T>(IEnumerable<T> datas, bool addTitle = true) where T : new()
        {
            // 获取导出模型属性信息字典
            var excelPropertyInfoNameDict = typeof(T).GetExportNamePropertyInfoDict();
            int rowIndex = _sheet.GetRowCount();

            // 表头
            if (addTitle)
            {
                // 设置表头
                var titleRow = _sheet.CreateRow(rowIndex++);
                int colIndex = 0;
                foreach (var property in excelPropertyInfoNameDict)
                {
                    var cell = titleRow.CreateCell(colIndex).SetValue(property.Key);

                    var exportHeader = property.Value.ExportHeader;
                    if (exportHeader == null)
                    {
                        exportHeader = new ExportHeaderAttribute(null);
                    }

                    var indexedColor = IndexedColors.ValueOf(exportHeader.ColorName);
                    cell.SetFont(font =>
                    {
                        font.FontHeight = exportHeader.FontSize * 20;
                        font.IsBold = exportHeader.IsBold;
                        font.Color = indexedColor?.Index ?? IndexedColors.Black.Index;
                    });

                    if (!string.IsNullOrEmpty(exportHeader.Comment))
                    {
                        cell.SetComment(exportHeader.Comment);
                    }

                    colIndex++;
                }
            }
            
            // 写入数据
            foreach (var data in datas)
            {
                var dataRow = _sheet.CreateRow(rowIndex++);
                var colIndex = 0;
                foreach (var property in excelPropertyInfoNameDict)
                {
                    var value = property.Value.PropertyInfo.GetValue(data);
                    
                    // 如果导出的是图片二进制数据
                    if (property.Value.ExportHeader != null && property.Value.ExportHeader.IsImage)
                    {
                        if (value is byte[] imageBytes)
                        {
                            dataRow.CreateCell(colIndex).SetImage(imageBytes);
                        }
                        continue;
                    }

                    var displayValue = property.Value.ExportMappedToDisplay(value);
                    var cell = dataRow.CreateCell(colIndex);
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

                    if (!string.IsNullOrEmpty(property.Value.ExportHeader?.Format))
                    {
                        cell.SetDataFormat(property.Value.ExportHeader?.Format);
                    }

                    colIndex++;
                }
            }

            // 设置列宽度
            {
                var colIndex = 0;
                foreach (var property in excelPropertyInfoNameDict)
                {
                    var exportHeader = property.Value.ExportHeader;
                    if (exportHeader == null)
                    {
                        exportHeader = new ExportHeaderAttribute(null);
                    }

                    if (exportHeader.IsAutoSizeColumn)
                    {
                        _sheet.AutoSizeColumn(colIndex);
                    }
                    else if (exportHeader.ColumnWidth > 0)
                    {
                        _sheet.SetColumnWidth(colIndex, exportHeader.ColumnWidth * 256);
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
            _sheet.CreateRow(rowIndex);

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
            var titleRow = _sheet.GetRow(0);
            foreach (var titleCell in titleRow)
            {
                var title = titleCell.GetData()?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    continue;
                }
                if (!excelPropertyInfoNameDict.ContainsKey(title))
                {
                    continue;
                }
                excelPropertyInfoIndexDict[titleCell.ColumnIndex] = excelPropertyInfoNameDict[title];
            }

            // 读取数据
            var rowCount = _sheet.GetRowCount();
            for (int i = 1; i < rowCount; i++)
            {
                var row = _sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }

                var t = new T();
                var hasValue = false;

                foreach (var excelPropertyInfo in excelPropertyInfoIndexDict)
                {
                    // 导入图片
                    if (excelPropertyInfo.Value.ImportIsImage())
                    {
                        var bytes = row.GetCellOrCreate(excelPropertyInfo.Key).GetImage();
                        excelPropertyInfo.Value.ImportCheckRequired(bytes);
                        excelPropertyInfo.Value.PropertyInfo.SetValue(t, bytes);
                        hasValue = true;
                        continue;
                    }

                    // 导入其它数据
                    var value = row.GetCell(excelPropertyInfo.Key).GetData();
                    excelPropertyInfo.Value.ImportCheckRequired(value);
                    excelPropertyInfo.Value.ImportTrim(ref value);
                    excelPropertyInfo.Value.ImportLimitCheckValue(value);

                    if (value != null)
                    {
                        var actualValue = excelPropertyInfo.Value.ImportMappedToActual(value);

                        excelPropertyInfo.Value.PropertyInfo.SetValueAuto(t, actualValue);

                        hasValue = true;
                    }
                }

                if (hasValue)
                {
                    result.Add(t);
                }
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
