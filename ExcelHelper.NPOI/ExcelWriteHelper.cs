using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// Excel 写入帮助类
    /// </summary>
    public class ExcelWriteHelper : BaseExcelWriteHelper
    {
        private readonly IWorkbook _excel;

        /// <summary>
        /// Excel 写入帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        public ExcelWriteHelper(ExcelHelperBuilder excelHelperBuilder) : base(excelHelperBuilder)
        {
            _excel = NpoiHelper.CreateExcel();
        }

        /// <inheritdoc/>
        public override IExcelWriteHelper ExportSheet<T>(string sheetName, IEnumerable<T> datas)
        {
            var sheet = _excel.CreateSheet(sheetName);

            // 获取导出模型属性信息字典
            var excelPropertyInfoNameDict = typeof(T).GetExportNamePropertyInfoDict();

            // 设置表头
            var titleRow = sheet.CreateRow(0);
            int colIndex = 0;
            foreach (var property in excelPropertyInfoNameDict)
            {
                var cell = titleRow.CreateCell(colIndex).SetValue(property.Key);

                var exportHeader = property.Value.ExportHeader;
                if (exportHeader == null)
                {
                    exportHeader = new ExportHeaderAttribute(null);
                }

                cell.SetFont(font =>
                {
                    font.FontHeight = exportHeader.FontSize * 20;
                    font.IsBold = exportHeader.IsBold;
                });

                if (!string.IsNullOrEmpty(exportHeader.Comment))
                {
                    cell.SetComment(exportHeader.Comment);
                }

                colIndex++;
            }

            // 写入数据
            int rowIndex = 1;
            foreach (var data in datas)
            {
                var dataRow = sheet.CreateRow(rowIndex++);
                colIndex = 0;
                foreach (var property in excelPropertyInfoNameDict)
                {
                    var value = property.Value.PropertyInfo.GetValue(data);
                    var displayValue = property.Value.ExportMappers.MappedToDisplay(value);
                    if (displayValue is DateTime dt)
                    {
                        if (DateTime.MinValue != dt)
                        {
                            dataRow.CreateCell(colIndex).SetValue(dt);
                        }
                    }
                    else if (displayValue is bool b)
                    {
                        dataRow.CreateCell(colIndex).SetValue(b);
                    }
                    else if (displayValue is double d)
                    {
                        dataRow.CreateCell(colIndex).SetValue(d);
                    }
                    else if (displayValue is int di)
                    {
                        dataRow.CreateCell(colIndex).SetValue(di);
                    }
                    else if (displayValue is decimal dc)
                    {
                        dataRow.CreateCell(colIndex).SetValue((double)dc);
                    }
                    else
                    {
                        dataRow.CreateCell(colIndex).SetValue(displayValue?.ToString());
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
                    sheet.AutoSizeColumn(colIndex);
                }
                else
                {
                    sheet.SetColumnWidth(colIndex, exportHeader.ColumnWidth * 256);
                }

                colIndex++;
            }

            return this;
        }

        /// <inheritdoc/>
        public override byte[] ToBytes()
        {
            return _excel.ToBytes();
        }
    }
}
