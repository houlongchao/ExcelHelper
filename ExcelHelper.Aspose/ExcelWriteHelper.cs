using Aspose.Cells;
using System;
using System.Collections.Generic;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// Excel 写入帮助类
    /// </summary>
    public class ExcelWriteHelper : BaseExcelWriteHelper
    {
        private readonly Workbook _excel;

        /// <summary>
        /// Excel 写入帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        public ExcelWriteHelper(ExcelHelperBuilder excelHelperBuilder) : base(excelHelperBuilder)
        {
            _excel = AsposeCellHelper.CreateExcel();
            _excel.Worksheets.Clear();
        }

        /// <inheritdoc/>
        public override IExcelWriteHelper ExportSheet<T>(string sheetName, IEnumerable<T> datas)
        {
            var sheet = _excel.CreateSheet(sheetName);

            // 获取导出模型属性信息字典
            var excelPropertyInfoNameDict = typeof(T).GetExportNamePropertyInfoDict();

            // 设置表头信息
            int colIndex = 0;
            foreach (var property in excelPropertyInfoNameDict)
            {
                var cell = sheet.CreateCell(0, colIndex++);
                cell.SetValue(property.Key);

                var exportHeader = property.Value.ExportHeader;
                if (exportHeader == null)
                {
                    exportHeader = new ExportHeaderAttribute(null);
                }

                cell.SetFont(font =>
                {
                    font.Size = exportHeader.FontSize;
                    font.IsBold = exportHeader.IsBold;
                });

                if (!string.IsNullOrEmpty(exportHeader.Comment))
                {
                    cell.SetComment(exportHeader.Comment);
                }
            }

            // 写数据
            int rowIndex = 0;
            foreach (var data in datas)
            {
                rowIndex++;
                colIndex = 0;
                foreach (var property in excelPropertyInfoNameDict)
                {
                    var value = property.Value.PropertyInfo.GetValue(data);
                    var displayValue = property.Value.ExportMappers.MappedToDisplay(value);
                    if (displayValue is DateTime dt)
                    {
                        if (DateTime.MinValue != dt)
                        {
                            var cell = sheet.CreateCell(rowIndex, colIndex);
                            cell.SetValue(dt);
                        }
                    }
                    else if (displayValue is bool b)
                    {
                        var cell = sheet.CreateCell(rowIndex, colIndex);
                        cell.SetValue(b);
                    }
                    else if (displayValue is double d)
                    {
                        var cell = sheet.CreateCell(rowIndex, colIndex);
                        cell.SetValue(d);
                    }
                    else if (displayValue is int di)
                    {
                        var cell = sheet.CreateCell(rowIndex, colIndex);
                        cell.SetValue(di);
                    }
                    else if (displayValue is decimal dc)
                    {
                        var cell = sheet.CreateCell(rowIndex, colIndex);
                        cell.SetValue((double)dc);
                    }
                    else
                    {
                        var cell = sheet.CreateCell(rowIndex, colIndex);
                        cell.SetValue(displayValue?.ToString());
                    }

                    colIndex++;
                }
            }

            // 设置列宽度
            colIndex = 0;
            foreach (var property in excelPropertyInfoNameDict)
            {
                var exportHeader = property.Value.ExportHeader;
                if (exportHeader == null)
                {
                    exportHeader = new ExportHeaderAttribute(null);
                }

                if (exportHeader.IsAutoSizeColumn)
                {
                    sheet.AutoFitColumn(colIndex);
                }
                else
                {
                    sheet.Cells.SetColumnWidth(colIndex, exportHeader.ColumnWidth);
                }

                colIndex++;
            }

            return this;
        }

        /// <inheritdoc/>
        public override IExcelWriteHelper SetSheetIndex(string sheetName, int index)
        {
            var sheet = _excel.GetSheet(sheetName);
            if (sheet == null)
            {
                return this;
            }

            sheet.MoveTo(index);

            return this;
        }

        /// <inheritdoc/>
        public override byte[] ToBytes()
        {
            return _excel.ToBytes();
        }
    }
}
