using Aspose.Cells;
using System;
using System.Collections.Generic;
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
                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    var cell = _sheet.CreateCell(rowIndex, colIndex++);
                    cell.SetValue(excelPropertyInfo.ExportHeaderTitle);

                    var exportHeader = excelPropertyInfo.ExportHeader;
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

                    if (!string.IsNullOrEmpty(excelPropertyInfo.ExportHeaderComment))
                    {
                        cell.SetComment(excelPropertyInfo.ExportHeaderComment);
                    }
                }
                rowIndex++;
            }

            // 写数据
            foreach (var data in datas)
            {
                var colIndex = 0;
                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    var value = excelPropertyInfo.PropertyInfo.GetValue(data);

                    // 如果导出的是图片二进制数据
                    if (excelPropertyInfo.IsExportImage())
                    {
                        if (value is byte[] imageBytes)
                        {
                            _sheet.CreateCell(rowIndex, colIndex).SetImage(imageBytes);
                        }
                        continue;
                    }

                    var displayValue = excelPropertyInfo.ExportMappedToDisplay(value);
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

                    if (!string.IsNullOrEmpty(excelPropertyInfo.ExportHeader?.Format))
                    {
                        cell.SetDataFormat(excelPropertyInfo.ExportHeader?.Format);
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
        public List<T> GetData<T>(ImportSetting importSetting = null) where T : new()
        {
            var result = new List<T>();

            // 读标题
            var titleIndexDict = new Dictionary<string, int>();
            var columnCount = _sheet.Cells.MaxColumn + 1;
            for (int i = 0; i < columnCount; i++)
            {
                var titleCell = _sheet.GetCell(0, i);
                var title = titleCell.GetData()?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    continue;
                }
                titleIndexDict[title] = i;
            }

            // 获取导入模型信息
            var excelObjectInfo = typeof(T).GetExcelObjectInfo();
            // 获取导入模型属性信息列表
            var excelPropertyInfoList = typeof(T).GetImportExcelPropertyInfoList(titleIndexDict, importSetting);

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
                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    // 导入图片
                    if (excelPropertyInfo.IsImportImage())
                    {
                        var bytes = row[excelPropertyInfo.ImportHeaderColumnIndex].GetImage();
                        excelPropertyInfo.ImportCheckRequired(bytes);
                        excelPropertyInfo.PropertyInfo.SetValue(t, bytes);
                        continue;
                    }

                    // 导入其它数据
                    var value = row.GetCell(excelPropertyInfo.ImportHeaderColumnIndex).GetData();
                    excelPropertyInfo.ImportCheckRequired(value);
                    excelPropertyInfo.ImportTrim(ref value);
                    excelPropertyInfo.ImportLimitCheckValue(value);
                    excelPropertyInfo.ImportCheckUnqiue(value);

                    var actualValue = excelPropertyInfo.ImportMappedToActual(value);

                    excelPropertyInfo.PropertyInfo.SetValueAuto(t, actualValue);
                }

                excelObjectInfo.CheckImportUnique(t);

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
