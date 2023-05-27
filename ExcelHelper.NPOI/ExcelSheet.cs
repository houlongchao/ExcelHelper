using NPOI.SS.UserModel;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public class ExcelSheet : BaseExcelSheet
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
        public override int GetRowCount()
        {
            return _sheet.GetRowCount();
        }

        /// <inheritdoc/>
        public override int GetColumnCount(int rowIndex = 0)
        {
            return _sheet.GetRow(rowIndex).LastCellNum;
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
            _sheet.AutoSizeColumn(colIndex);
        }

        /// <inheritdoc/>
        public override void SetColumnWidth(int colIndex, int width)
        {
            _sheet.SetColumnWidth(colIndex, width * 256);
        }

        /// <inheritdoc/>
        public override void SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true)
        {
            var indexedColor = IndexedColors.ValueOf(colorName);
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetFont(font =>
            {
                font.FontHeight = fontSize * 20;
                font.IsBold = isBold;
                font.Color = indexedColor?.Index ?? IndexedColors.Black.Index;
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
        //    int rowIndex = _sheet.GetRowCount();

        //    // 写表头
        //    if (exportSetting.AddTitle)
        //    {
        //        // 设置表头
        //        var titleRow = _sheet.CreateRow(rowIndex++);
        //        int colIndex = 0;
        //        foreach (var excelPropertyInfo in excelPropertyInfoList)
        //        {
        //            var cell = titleRow.CreateCell(colIndex).SetValue(excelPropertyInfo.ExportHeaderTitle);

        //            var exportHeader = excelPropertyInfo.ExportHeaderAttribute;
        //            if (exportHeader == null)
        //            {
        //                exportHeader = new ExportHeaderAttribute(null);
        //            }

        //            var indexedColor = IndexedColors.ValueOf(exportHeader.ColorName);
        //            cell.SetFont(font =>
        //            {
        //                font.FontHeight = exportHeader.FontSize * 20;
        //                font.IsBold = exportHeader.IsBold;
        //                font.Color = indexedColor?.Index ?? IndexedColors.Black.Index;
        //            });

        //            if (!string.IsNullOrEmpty(excelPropertyInfo.ExportHeaderComment))
        //            {
        //                cell.SetComment(excelPropertyInfo.ExportHeaderComment);
        //            }
        //            colIndex++;
        //        }
        //    }

        //    // 写入数据
        //    foreach (var data in datas)
        //    {
        //        var dataRow = _sheet.CreateRow(rowIndex++);
        //        var colIndex = 0;
        //        foreach (var property in excelPropertyInfoList)
        //        {
        //            var value = property.PropertyInfo.GetValue(data);

        //            // 如果导出的是图片二进制数据
        //            if (property.IsImage())
        //            {
        //                if (value is byte[] imageBytes)
        //                {
        //                    dataRow.CreateCell(colIndex).SetImage(imageBytes);
        //                }
        //                continue;
        //            }

        //            var displayValue = property.ExportMappedToDisplay(value);
        //            var cell = dataRow.CreateCell(colIndex);
        //            cell.SetValue(displayValue);

        //            if (!string.IsNullOrEmpty(property.ExportFormatAttribute?.Format))
        //            {
        //                cell.SetDataFormat(property.ExportFormatAttribute?.Format);
        //            }

        //            colIndex++;
        //        }
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
        //                _sheet.AutoSizeColumn(colIndex);
        //            }
        //            else if (exportHeader.ColumnWidth > 0)
        //            {
        //                _sheet.SetColumnWidth(colIndex, exportHeader.ColumnWidth * 256);
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
        //    _sheet.CreateRow(rowIndex);

        //    return this;
        //}

        ///// <inheritdoc/>
        //public List<T> GetData<T>(ImportSetting importSetting = null) where T : new()
        //{
        //    var result = new List<T>();

        //    // 读标题
        //    var titleIndexDict = new Dictionary<string, int>();
        //    var titleRow = _sheet.GetRow(0);
        //    foreach (var titleCell in titleRow)
        //    {
        //        var title = titleCell.GetData()?.ToString();
        //        if (string.IsNullOrEmpty(title))
        //        {
        //            continue;
        //        }
        //        titleIndexDict[title] = titleCell.ColumnIndex;
        //    }

        //    // 获取导入模型信息
        //    var excelObjectInfo = typeof(T).GetExcelObjectInfo();
        //    // 获取导入模型属性信息列表
        //    var excelPropertyInfoList = typeof(T).GetImportExcelPropertyInfoList(titleIndexDict, importSetting);

        //    // 读取数据
        //    var rowCount = _sheet.GetRowCount();
        //    for (int i = 1; i < rowCount; i++)
        //    {
        //        var row = _sheet.GetRow(i);
        //        if (row == null)
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
        //                var bytes = row.GetCellOrCreate(excelPropertyInfo.ImportHeaderColumnIndex).GetImage();
        //                excelPropertyInfo.PropertyInfo.SetValue(t, bytes);
        //                hasValue = true;
        //                continue;
        //            }
        //            else
        //            {
        //                // 导入其它数据
        //                var value = row.GetCell(excelPropertyInfo.ImportHeaderColumnIndex).GetData();

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
        //                    ICell cell = null;
        //                    switch (excelPropertyInfo.TempListType)
        //                    {
        //                        case TempListType.Row:
        //                            {
        //                                cell = _sheet.GetOrCreateCell(i, child.TempListItemIndex);
        //                                break;
        //                            }
        //                        case TempListType.Column:
        //                            {
        //                                cell = _sheet.GetOrCreateCell(child.TempListItemIndex, i);
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
        //                    ICell cell = null;
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
        //            var cell = _sheet.GetOrCreateCell(excelPropertyInfo.TempCellAddress);
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
