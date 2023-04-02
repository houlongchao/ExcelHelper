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
            int rowIndex = _sheet.GetRowCount();

            // 写表头
            if (exportSetting.AddTitle)
            {
                // 设置表头
                var titleRow = _sheet.CreateRow(rowIndex++);
                int colIndex = 0;
                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    var cell = titleRow.CreateCell(colIndex).SetValue(excelPropertyInfo.ExportHeaderTitle);

                    var exportHeader = excelPropertyInfo.ExportHeaderAttribute;
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

                    if (!string.IsNullOrEmpty(excelPropertyInfo.ExportHeaderComment))
                    {
                        cell.SetComment(excelPropertyInfo.ExportHeaderComment);
                    }
                    colIndex++;
                }
            }
            
            // 写入数据
            foreach (var data in datas)
            {
                var dataRow = _sheet.CreateRow(rowIndex++);
                var colIndex = 0;
                foreach (var property in excelPropertyInfoList)
                {
                    var value = property.PropertyInfo.GetValue(data);
                    
                    // 如果导出的是图片二进制数据
                    if (property.IsExportImage())
                    {
                        if (value is byte[] imageBytes)
                        {
                            dataRow.CreateCell(colIndex).SetImage(imageBytes);
                        }
                        continue;
                    }

                    var displayValue = property.ExportMappedToDisplay(value);
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

                    if (!string.IsNullOrEmpty(property.ExportHeaderAttribute?.Format))
                    {
                        cell.SetDataFormat(property.ExportHeaderAttribute?.Format);
                    }

                    colIndex++;
                }
            }

            // 设置列宽度
            {
                var colIndex = 0;
                foreach (var property in excelPropertyInfoList)
                {
                    var exportHeader = property.ExportHeaderAttribute;
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
        public List<T> GetData<T>(ImportSetting importSetting = null) where T : new()
        {
            var result = new List<T>();

            // 读标题
            var titleIndexDict = new Dictionary<string, int>();
            var titleRow = _sheet.GetRow(0);
            foreach (var titleCell in titleRow)
            {
                var title = titleCell.GetData()?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    continue;
                }
                titleIndexDict[title] = titleCell.ColumnIndex;
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
                var hasValue = false;

                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    // 导入图片
                    if (excelPropertyInfo.IsImportImage())
                    {
                        var bytes = row.GetCellOrCreate(excelPropertyInfo.ImportHeaderColumnIndex).GetImage();
                        excelPropertyInfo.ImportCheckRequired(bytes);
                        excelPropertyInfo.PropertyInfo.SetValue(t, bytes);
                        hasValue = true;
                        continue;
                    }

                    // 导入其它数据
                    var value = row.GetCell(excelPropertyInfo.ImportHeaderColumnIndex).GetData();
                    excelPropertyInfo.ImportCheckRequired(value);
                    excelPropertyInfo.ImportTrim(ref value);
                    excelPropertyInfo.ImportLimitCheckValue(value);
                    excelPropertyInfo.ImportCheckUnqiue(value);

                    if (value != null)
                    {
                        var actualValue = excelPropertyInfo.ImportMappedToActual(value);

                        excelPropertyInfo.PropertyInfo.SetValueAuto(t, actualValue);

                        hasValue = true;
                    }
                }

                if (hasValue)
                {
                    excelObjectInfo.CheckImportUnique(t);
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
