using Aspose.Cells;
using System.Drawing;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public class ExcelSheet : BaseExcelSheet
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
        public override int GetRowCount()
        {
            return _sheet.GetRowCount();
        }

        /// <inheritdoc/>
        public override int GetColumnCount(int rowIndex = 0)
        {
            return _sheet.GetRow(rowIndex).LastDataCell.Column + 1;
        }

        /// <inheritdoc/>
        public override byte[] ToBytes()
        {
            return _sheet.Workbook.ToBytes();
        }

        /// <inheritdoc/>
        public override byte[] GetImage(int rowIndex, int colIndex)
        {
            return _sheet.GetOrCreateCell(rowIndex, colIndex).GetImage();
        }

        /// <inheritdoc/>
        public override byte[] GetImage(string cellAddress)
        {
            return _sheet.GetOrCreateCell(cellAddress).GetImage();
        }

        /// <inheritdoc/>
        public override object GetValue(int rowIndex, int colIndex)
        {
            return _sheet.GetOrCreateCell(rowIndex, colIndex).GetData();
        }

        /// <inheritdoc/>
        public override object GetValue(string cellAddress)
        {
            return _sheet.GetOrCreateCell(cellAddress).GetData();
        }

        /// <inheritdoc/>
        public override void SetValue(int rowIndex, int colIndex, object value)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetValue(value);
        }

        /// <inheritdoc/>
        public override void SetValue(string cellAddress, object value)
        {
            _sheet.GetOrCreateCell(cellAddress).SetValue(value);
        }

        /// <inheritdoc/>
        public override void SetImage(int rowIndex, int colIndex, byte[] value)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetImage(value);
        }

        /// <inheritdoc/>
        public override void SetImage(string cellAddress, byte[] value)
        {
            _sheet.GetOrCreateCell(cellAddress).SetImage(value);
        }

        /// <inheritdoc/>
        public override void SetComment(int rowIndex, int colIndex, string comment)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetComment(comment);
        }

        /// <inheritdoc/>
        public override void SetFormat(int rowIndex, int colIndex, string format)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetDataFormat(format);
        }

        /// <inheritdoc/>
        public override void SetAutoSizeColumn(int colIndex)
        {
            _sheet.AutoFitColumn(colIndex);
        }

        /// <inheritdoc/>
        public override void SetColumnWidth(int colIndex, int width)
        {
            _sheet.Cells.SetColumnWidth(colIndex, width);
        }

        /// <inheritdoc/>
        public override void SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetFont(font =>
                    {
                        font.Size = fontSize;
                        font.IsBold = isBold;
                        font.Color = Color.FromName(colorName);
                    });
        }

        ///// <inheritdoc/>
        //public IExcelSheet AppendData<T>(IEnumerable<T> datas, bool addTitle = true) where T : new()
        //{
        //    var exportSetting = new ExportSetting();
        //    exportSetting.AddTitle = addTitle;

        //    return AppendData(datas, exportSetting);
        //}

        ///// <inheritdoc/>
        //public IExcelSheet AppendData<T>(IEnumerable<T> datas, ExportSetting exportSetting) where T : new()
        //{
        //    if (exportSetting == null)
        //    {
        //        exportSetting = new ExportSetting();
        //    }

        //    // 获取导出模型属性信息列表
        //    var excelPropertyInfoList = typeof(T).GetExportExcelPropertyInfoList(exportSetting);

        //    var rowIndex = _sheet.GetRowCount();

        //    // 设置表头信息
        //    if (exportSetting.AddTitle)
        //    {
        //        int colIndex = 0;
        //        foreach (var excelPropertyInfo in excelPropertyInfoList)
        //        {
        //            var cell = _sheet.CreateCell(rowIndex, colIndex++);
        //            cell.SetValue(excelPropertyInfo.ExportHeaderTitle);

        //            var exportHeader = excelPropertyInfo.ExportHeaderAttribute;
        //            if (exportHeader == null)
        //            {
        //                exportHeader = new ExportHeaderAttribute(null);
        //            }

        //            cell.SetFont(font =>
        //            {
        //                font.Size = exportHeader.FontSize;
        //                font.IsBold = exportHeader.IsBold;
        //                font.Color = Color.FromName(exportHeader.ColorName);
        //            });

        //            if (!string.IsNullOrEmpty(excelPropertyInfo.ExportHeaderComment))
        //            {
        //                cell.SetComment(excelPropertyInfo.ExportHeaderComment);
        //            }
        //        }
        //        rowIndex++;
        //    }

        //    // 写数据
        //    foreach (var data in datas)
        //    {
        //        var colIndex = 0;
        //        foreach (var excelPropertyInfo in excelPropertyInfoList)
        //        {
        //            var value = excelPropertyInfo.PropertyInfo.GetValue(data);

        //            // 如果导出的是图片二进制数据
        //            if (excelPropertyInfo.IsImage())
        //            {
        //                if (value is byte[] imageBytes)
        //                {
        //                    _sheet.CreateCell(rowIndex, colIndex).SetImage(imageBytes);
        //                }
        //                continue;
        //            }

        //            var displayValue = excelPropertyInfo.ExportMappedToDisplay(value);
        //            var cell = _sheet.CreateCell(rowIndex, colIndex);
        //            cell.SetValue(displayValue);

        //            if (!string.IsNullOrEmpty(excelPropertyInfo.ExportFormatAttribute?.Format))
        //            {
        //                cell.SetDataFormat(excelPropertyInfo.ExportFormatAttribute?.Format);
        //            }

        //            colIndex++;
        //        }
        //        rowIndex++;
        //    }

        //    // 设置列宽度
        //    {
        //        var colIndex = 0;
        //        foreach (var property in excelPropertyInfoList)
        //        {
        //            var exportHeader = property.ExportHeaderAttribute;
        //            if (exportHeader == null)
        //            {
        //                exportHeader = new ExportHeaderAttribute(null);
        //            }

        //            if (exportHeader.IsAutoSizeColumn)
        //            {
        //                _sheet.AutoFitColumn(colIndex);
        //            }
        //            else if (exportHeader.ColumnWidth > 0)
        //            {
        //                _sheet.Cells.SetColumnWidth(colIndex, exportHeader.ColumnWidth);
        //            }

        //            colIndex++;
        //        }

        //    }

        //    return this;
        //}

        ///// <inheritdoc/>
        //public IExcelSheet AppendEmptyRow()
        //{
        //    int rowIndex = _sheet.GetRowCount();
        //    _sheet.CreateCell(rowIndex, 0).SetValue(null);

        //    return this;
        //}

        ///// <inheritdoc/>
        //public List<T> GetData<T>(ImportSetting importSetting = null) where T : new()
        //{
        //    var result = new List<T>();

        //    // 读标题
        //    var titleIndexDict = new Dictionary<string, int>();
        //    var columnCount = _sheet.Cells.MaxColumn + 1;
        //    for (int i = 0; i < columnCount; i++)
        //    {
        //        var titleCell = _sheet.GetCell(0, i);
        //        var title = titleCell.GetData()?.ToString();
        //        if (string.IsNullOrEmpty(title))
        //        {
        //            continue;
        //        }
        //        titleIndexDict[title] = i;
        //    }

        //    // 获取导入模型信息
        //    var excelObjectInfo = typeof(T).GetExcelObjectInfo();
        //    // 获取导入模型属性信息列表
        //    var excelPropertyInfoList = typeof(T).GetImportExcelPropertyInfoList(titleIndexDict, importSetting);

        //    // 读取数据
        //    var rowCount = _sheet.GetRowCount();
        //    for (int i = 1; i < rowCount; i++)
        //    {
        //        var rowIndex = _sheet.GetRow(i);
        //        if (rowIndex == null)
        //        {
        //            continue;
        //        }
        //        var t = new T();
        //        var hasValue = false;

        //        foreach (var excelPropertyInfo in excelPropertyInfoList)
        //        {
        //            // 导入图片
        //            if (excelPropertyInfo.IsImage())
        //            {
        //                var bytes = rowIndex[excelPropertyInfo.ImportHeaderColumnIndex].GetImage();
        //                excelPropertyInfo.PropertyInfo.SetValue(t, bytes);
        //                hasValue = true;
        //                continue;
        //            }
        //            else
        //            {   
        //                // 导入其它数据
        //                var value = rowIndex.GetCell(excelPropertyInfo.ImportHeaderColumnIndex).GetData();
        //                if (value != null)
        //                {
        //                    excelPropertyInfo.ImportTrim(ref value);
        //                    var actualValue = excelPropertyInfo.ImportMappedToActual(value);
        //                    excelPropertyInfo.ImportCheckLimitValue(value);
        //                    excelPropertyInfo.ImportCheckUnqiue(value);
        //                    excelPropertyInfo.PropertyInfo.SetValueAuto(t, actualValue);

        //                    hasValue = true;
        //                }
        //            }


        //        }

        //        if (hasValue)
        //        {
        //            foreach (var excelPropertyInfo in excelPropertyInfoList)
        //            {
        //                excelPropertyInfo.ImportCheckRequired(excelPropertyInfo.PropertyInfo.GetValue(t));
        //            }

        //            excelObjectInfo.CheckImportUnique(t, importSetting);
        //            result.Add(t);
        //        }
        //    }

        //    return result;
        //}

        ///// <inheritdoc/>
        //public int GetRowCount()
        //{
        //    return _sheet.GetRowCount();
        //}


        ///// <inheritdoc/>
        //public byte[] ToBytes()
        //{
        //    return _sheet.Workbook.ToBytes();
        //}

        ///// <inheritdoc/>
        //public T GetTempData<T>(TempSetting tempSetting = null) where T : new()
        //{
        //    // 获取导入模型属性信息列表
        //    var excelPropertyInfoList = typeof(T).GetTempExcelPropertyInfoList(tempSetting);

        //    var result = new T();

        //    foreach (var excelPropertyInfo in excelPropertyInfoList)
        //    {
        //        if (excelPropertyInfo.IsArray)
        //        {
        //            var newList = Activator.CreateInstance(typeof(List<>).MakeGenericType(excelPropertyInfo.PropertyInfo.PropertyType.GenericTypeArguments));
        //            excelPropertyInfo.PropertyInfo.SetValueAuto(result, newList);
        //            var addMethod = newList.GetType().GetMethod("Add");
        //            for (int i = excelPropertyInfo.TempListStartIndex; i <= excelPropertyInfo.TempListEndIndex; i++)
        //            {
        //                var ct = Activator.CreateInstance(excelPropertyInfo.PropertyInfo.PropertyType.GenericTypeArguments[0]);
        //                bool hasValue = false;
        //                foreach (var child in excelPropertyInfo.Children)
        //                {
        //                    Cell cell = null;
        //                    switch (excelPropertyInfo.TempListType)
        //                    {
        //                        case TempListType.Row:
        //                            {
        //                                cell = _sheet.GetCell(i, child.TempListItemIndex);
        //                                break;
        //                            }
        //                        case TempListType.Column:
        //                            {
        //                                cell = _sheet.GetCell(child.TempListItemIndex, i);
        //                                break;
        //                            }
        //                        default:
        //                            break;
        //                    }
        //                    if (child.IsImage())
        //                    {
        //                        var bytes = cell.GetImage();
        //                        excelPropertyInfo.PropertyInfo.SetValue(result, bytes);
        //                    }
        //                    else
        //                    {
        //                        // 导入其它数据
        //                        var value = cell.GetData();
        //                        if (value != null)
        //                        {
        //                            child.ImportTrim(ref value);
        //                            var actualValue = child.ImportMappedToActual(value);
        //                            child.ImportCheckLimitValue(value);
        //                            child.ImportCheckUnqiue(value);
        //                            child.PropertyInfo.SetValueAuto(ct, actualValue);
        //                            hasValue = true;
        //                        }
        //                    }
        //                }
        //                if (hasValue)
        //                {
        //                    foreach (var child in excelPropertyInfo.Children)
        //                    {
        //                        child.ImportCheckRequired(child.PropertyInfo.GetValue(ct));
        //                    }
        //                    addMethod.Invoke(newList, new object[] { ct });
        //                }
        //            }
        //        }
        //        else
        //        {
        //            var cell = _sheet.GetOrCreateCell(excelPropertyInfo.TempCellAddress);
        //            // 导入图片
        //            if (excelPropertyInfo.IsImage())
        //            {
        //                var bytes = cell.GetImage();
        //                excelPropertyInfo.ImportCheckRequired(bytes);
        //                excelPropertyInfo.PropertyInfo.SetValue(result, bytes);
        //            }
        //            else
        //            {
        //                // 导入其它数据
        //                var value = cell.GetData();
        //                excelPropertyInfo.ImportCheckRequired(value);
        //                excelPropertyInfo.ImportTrim(ref value);
        //                excelPropertyInfo.ImportCheckLimitValue(value);
        //                if (value != null)
        //                {
        //                    var actualValue = excelPropertyInfo.ImportMappedToActual(value);
        //                    excelPropertyInfo.PropertyInfo.SetValueAuto(result, actualValue);
        //                }
        //            }
        //        }

        //    }

        //    return result;
        //}

        ///// <inheritdoc/>
        //public IExcelSheet SetTempData<T>(T data, TempSetting tempSetting = null) where T : new()
        //{
        //    // 获取导出模型属性信息列表
        //    var excelPropertyInfoList = typeof(T).GetTempExcelPropertyInfoList(tempSetting);
        //    foreach (var excelPropertyInfo in excelPropertyInfoList)
        //    {
        //        if (excelPropertyInfo.IsArray)
        //        {
        //            var arrayData = excelPropertyInfo.PropertyInfo.GetValue(data) as IEnumerable;
        //            int i = excelPropertyInfo.TempListStartIndex;
        //            foreach (var itemData in arrayData)
        //            {
        //                if (i > excelPropertyInfo.TempListEndIndex)
        //                {
        //                    break;
        //                }

        //                foreach (var childProperty in excelPropertyInfo.Children)
        //                {
        //                    Cell cell = null;
        //                    switch (excelPropertyInfo.TempListType)
        //                    {
        //                        case TempListType.Row:
        //                            {
        //                                cell = _sheet.GetOrCreateCell(i, childProperty.TempListItemIndex);
        //                                break;
        //                            }
        //                        case TempListType.Column:
        //                            {
        //                                cell = _sheet.GetOrCreateCell(childProperty.TempListItemIndex, i);
        //                                break;
        //                            }
        //                        default:
        //                            break;
        //                    }
        //                    var value = childProperty.PropertyInfo.GetValue(itemData);
        //                    // 如果导出的是图片二进制数据
        //                    if (childProperty.IsImage())
        //                    {
        //                        if (value is byte[] imageBytes)
        //                        {
        //                            cell.SetImage(imageBytes);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        var displayValue = childProperty.ExportMappedToDisplay(value);
        //                        cell.SetValue(displayValue);
        //                    }
        //                }
        //                ++i;
        //            }
        //        }
        //        else
        //        {
        //            var cell = _sheet.GetCell(excelPropertyInfo.TempCellAddress);
        //            var value = excelPropertyInfo.PropertyInfo.GetValue(data);

        //            // 如果导出的是图片二进制数据
        //            if (excelPropertyInfo.IsImage())
        //            {
        //                if (value is byte[] imageBytes)
        //                {
        //                    cell.SetImage(imageBytes);
        //                }
        //            }
        //            else
        //            {
        //                var displayValue = excelPropertyInfo.ExportMappedToDisplay(value);
        //                cell.SetValue(displayValue);
        //            }
        //        }
        //    }
        //    return this;
        //}
    }
}
