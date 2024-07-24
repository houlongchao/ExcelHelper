using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// Excel Sheet 基础操作实现
    /// </summary>
    public abstract class BaseExcelSheet : IExcelSheet
    {
        #region abstract

        /// <summary>
        /// 获取总行数
        /// </summary>
        /// <returns></returns>
        public abstract int GetRowCount();

        /// <summary>
        /// 获取指定行的列数
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public abstract int GetColumnCount(int rowIndex = 0);

        /// <summary>
        /// 生成字节数组
        /// </summary>
        /// <returns></returns>
        public abstract byte[] ToBytes();

        /// <summary>
        /// 获取指定位置的图片数据
        /// </summary>
        /// <returns></returns>
        public abstract byte[] GetImage(int rowIndex, int colIndex);

        /// <summary>
        /// 获取指定位置的图片数据
        /// </summary>
        /// <returns></returns>
        public abstract byte[] GetImage(string cellAddress);

        /// <summary>
        /// 获取指定位置的数据
        /// </summary>
        /// <returns></returns>
        public abstract object GetValue(int rowIndex, int colIndex);

        /// <summary>
        /// 获取指定位置的数据
        /// </summary>
        /// <returns></returns>
        public abstract object GetValue(string cellAddress);

        /// <summary>
        /// 设置指定位置的数据
        /// </summary>
        /// <returns></returns>
        public abstract void SetValue(int rowIndex, int colIndex, object value);

        /// <summary>
        /// 设置指定位置的数据
        /// </summary>
        /// <returns></returns>
        public abstract void SetValue(string cellAddress, object value);

        /// <summary>
        /// 设置指定位置的图片数据
        /// </summary>
        /// <returns></returns>
        public abstract void SetImage(int rowIndex, int colIndex, byte[] value);

        /// <summary>
        /// 设置指定位置的图片数据
        /// </summary>
        /// <returns></returns>
        public abstract void SetImage(string cellAddress, byte[] value);

        /// <summary>
        /// 设置指定位置的备注信息
        /// </summary>
        public abstract void SetComment(int rowIndex, int colIndex, string comment);

        /// <summary>
        /// 设置指定位置的格式化字符串
        /// </summary>
        public abstract void SetFormat(int rowIndex, int colIndex, string format);

        /// <summary>
        /// 设置指定列自动调整宽度
        /// </summary>
        public abstract void SetAutoSizeColumn(int colIndex);

        /// <summary>
        /// 设置指定列的宽度
        /// </summary>
        public abstract void SetColumnWidth(int colIndex, int width);

        /// <summary>
        /// 设置指定位置的字体
        /// </summary>
        public abstract void SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true);

        /// <summary>
        /// 设置验证数据
        /// </summary>
        /// <param name="firstRowIndex"></param>
        /// <param name="lastRowIndex"></param>
        /// <param name="firstColIndex"></param>
        /// <param name="lastColIndex"></param>
        /// <param name="explicitListValues"></param>
        public abstract void SetValidationData(int firstRowIndex, int lastRowIndex, int firstColIndex, int lastColIndex, string[] explicitListValues);

        #endregion

        #region Implement

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
            int rowIndex = GetRowCount();

            // 写表头
            if (exportSetting.AddTitle)
            {
                // 设置表头
                int colIndex = 0;
                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    SetValue(rowIndex, colIndex, excelPropertyInfo.ExportHeaderTitle);

                    var exportHeader = excelPropertyInfo.ExportHeaderAttribute;
                    if (exportHeader == null)
                    {
                        exportHeader = new ExportHeaderAttribute(null);
                    }

                    SetFont(rowIndex, colIndex, exportHeader.ColorName, exportHeader.FontSize, exportHeader.IsBold);

                    if (!string.IsNullOrEmpty(excelPropertyInfo.ExportHeaderComment))
                    {
                        SetComment(rowIndex, colIndex, excelPropertyInfo.ExportHeaderComment);
                    }
                    colIndex++;
                }
                rowIndex++;
            }

            int firstRowIndex = rowIndex;
            // 写入数据
            foreach (var data in datas)
            {
                var colIndex = 0;
                foreach (var property in excelPropertyInfoList)
                {
                    var value = property.GetValue(data);

                    // 如果导出的是图片二进制数据
                    if (property.IsImage())
                    {
                        if (value is byte[] imageBytes)
                        {
                            SetImage(rowIndex, colIndex, imageBytes);
                        }
                        continue;
                    }

                    var displayValue = property.ExportMappedToDisplay(value);
                    SetValue(rowIndex, colIndex, displayValue);

                    if (!string.IsNullOrEmpty(property.ExportFormatAttribute?.Format))
                    {
                        SetFormat(rowIndex, colIndex, property.ExportFormatAttribute?.Format);
                    }

                    colIndex++;
                }
                rowIndex++;
            }

           int lastRowIndex = rowIndex - 1;

            // 设置列宽度及数据样式属性
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
                        SetAutoSizeColumn(colIndex);
                    }
                    else if (exportHeader.ColumnWidth > 0)
                    {
                        SetColumnWidth(colIndex, exportHeader.ColumnWidth);
                    }

                    if (property.ExportValidationsAttribute != null)
                    {
                        SetValidationData(firstRowIndex, lastRowIndex, colIndex, colIndex, property.ExportValidationsAttribute.Validations);
                    }

                    colIndex++;
                }
            }

            return this;
        }

        /// <inheritdoc/>
        public IExcelSheet AppendEmptyRow()
        {
            int rowIndex = GetRowCount();
            SetValue(rowIndex, 0, null);

            return this;
        }

        /// <inheritdoc/>
        public List<T> GetData<T>(ImportSetting importSetting = null) where T : new()
        {
            var result = new List<T>();

            // 读标题
            var titleIndexDict = new Dictionary<string, int>();
            var columnCount = GetColumnCount(0);
            for (int i = 0; i < columnCount; i++)
            {
                var title = GetValue(0, i)?.ToString();
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
            var rowCount = GetRowCount();
            for (int i = 1; i < rowCount; i++)
            {
                var t = new T();
                var hasValue = false;

                foreach (var excelPropertyInfo in excelPropertyInfoList)
                {
                    // 导入图片
                    if (excelPropertyInfo.IsImage())
                    {
                        var bytes = GetImage(i, excelPropertyInfo.ImportHeaderColumnIndex);
                        if (bytes != null)
                        {
                            excelPropertyInfo.SetValue(t, bytes);
                            hasValue = true;
                        }
                    }
                    else
                    {
                        // 导入其它数据
                        var value = GetValue(i, excelPropertyInfo.ImportHeaderColumnIndex);
                        if (value != null)
                        {
                            excelPropertyInfo.ImportTrim(ref value);
                            var actualValue = excelPropertyInfo.ImportMappedToActual(value);
                            excelPropertyInfo.ImportCheckLimitValue(value);
                            excelPropertyInfo.ImportCheckUnqiue(value);
                            excelPropertyInfo.SetValue(t, actualValue);

                            hasValue = true;
                        }
                    }
                }

                if (hasValue)
                {
                    foreach (var excelPropertyInfo in excelPropertyInfoList)
                    {
                        excelPropertyInfo.ImportCheckRequired(excelPropertyInfo.GetValue(t));
                    }

                    excelObjectInfo.CheckImportUnique(t, importSetting);
                    result.Add(t);
                }
            }

            return result;
        }

        /// <inheritdoc/>
        public T GetTempData<T>(TempSetting tempSetting = null) where T : new()
        {
            // 获取导入模型属性信息列表
            var excelPropertyInfoList = typeof(T).GetTempExcelPropertyInfoList(tempSetting);

            var result = new T();

            foreach (var excelPropertyInfo in excelPropertyInfoList)
            {
                if (excelPropertyInfo.IsArray)
                {
                    var newList = excelPropertyInfo.CreateListObject();
                    excelPropertyInfo.SetValue(result, newList);
                    var addMethod = newList.GetType().GetMethod("Add");
                    for (int i = excelPropertyInfo.TempListStartIndex; i <= excelPropertyInfo.TempListEndIndex; i++)
                    {
                        var ct = excelPropertyInfo.CreateGenericTypeObject();
                        bool hasValue = false;
                        foreach (var child in excelPropertyInfo.Children)
                        {
                            if (child.IsImage())
                            {
                                byte[] bytes = null;
                                switch (excelPropertyInfo.TempListType)
                                {
                                    case TempListType.Row:
                                        {
                                            bytes = GetImage(i, child.TempListItemIndex);
                                            break;
                                        }
                                    case TempListType.Column:
                                        {
                                            bytes = GetImage(child.TempListItemIndex, i);
                                            break;
                                        }
                                    default:
                                        break;
                                }
                                if (bytes != null)
                                {
                                    excelPropertyInfo.SetValue(result, bytes);
                                    hasValue = true;
                                }
                            }
                            else
                            {
                                // 导入其它数据
                                object value = null;
                                switch (excelPropertyInfo.TempListType)
                                {
                                    case TempListType.Row:
                                        {
                                            value = GetValue(i, child.TempListItemIndex);
                                            break;
                                        }
                                    case TempListType.Column:
                                        {
                                            value = GetValue(child.TempListItemIndex, i);
                                            break;
                                        }
                                    default:
                                        break;
                                }
                                if (value != null)
                                {
                                    child.ImportTrim(ref value);
                                    var actualValue = child.ImportMappedToActual(value);
                                    child.ImportCheckLimitValue(value);
                                    child.ImportCheckUnqiue(value);
                                    child.SetValue(ct, actualValue);
                                    hasValue = true;
                                }
                            }
                        }
                        if (hasValue)
                        {
                            foreach (var child in excelPropertyInfo.Children)
                            {
                                child.ImportCheckRequired(child.GetValue(ct));
                            }
                            addMethod.Invoke(newList, new object[] { ct });
                        }
                    }
                }
                else
                {
                    // 导入图片
                    if (excelPropertyInfo.IsImage())
                    {
                        var bytes = GetImage(excelPropertyInfo.TempCellAddress);
                        excelPropertyInfo.ImportCheckRequired(bytes);
                        excelPropertyInfo.SetValue(result, bytes);
                    }
                    else
                    {
                        // 导入其它数据
                        var value = GetValue(excelPropertyInfo.TempCellAddress);
                        excelPropertyInfo.ImportCheckRequired(value);
                        excelPropertyInfo.ImportTrim(ref value);
                        excelPropertyInfo.ImportCheckLimitValue(value);
                        if (value != null)
                        {
                            var actualValue = excelPropertyInfo.ImportMappedToActual(value);
                            excelPropertyInfo.SetValue(result, actualValue);
                        }
                    }
                }

            }

            return result;
        }

        /// <inheritdoc/>
        public IExcelSheet SetTempData<T>(T data, TempSetting tempSetting = null) where T : new()
        {
            // 获取导出模型属性信息列表
            var excelPropertyInfoList = typeof(T).GetTempExcelPropertyInfoList(tempSetting);
            foreach (var excelPropertyInfo in excelPropertyInfoList)
            {
                if (excelPropertyInfo.IsArray)
                {
                    var arrayData = excelPropertyInfo.GetValue(data) as IEnumerable;
                    int i = excelPropertyInfo.TempListStartIndex;
                    foreach (var itemData in arrayData)
                    {
                        if (i > excelPropertyInfo.TempListEndIndex)
                        {
                            break;
                        }

                        foreach (var childProperty in excelPropertyInfo.Children)
                        {
                            var value = childProperty.GetValue(itemData);

                            // 如果导出的是图片二进制数据
                            if (childProperty.IsImage())
                            {
                                if (value is byte[] imageBytes)
                                {
                                    switch (excelPropertyInfo.TempListType)
                                    {
                                        case TempListType.Row:
                                            {
                                                SetImage(i, childProperty.TempListItemIndex, imageBytes);
                                                break;
                                            }
                                        case TempListType.Column:
                                            {
                                                SetImage(childProperty.TempListItemIndex, i, imageBytes);
                                                break;
                                            }
                                        default:
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                var displayValue = childProperty.ExportMappedToDisplay(value); 
                                switch (excelPropertyInfo.TempListType)
                                {
                                    case TempListType.Row:
                                        {
                                            SetValue(i, childProperty.TempListItemIndex, displayValue);
                                            break;
                                        }
                                    case TempListType.Column:
                                        {
                                            SetValue(childProperty.TempListItemIndex, i, displayValue);
                                            break;
                                        }
                                    default:
                                        break;
                                }
                            }
                        }
                        ++i;
                    }
                }
                else
                {
                    var value = excelPropertyInfo.GetValue(data);

                    // 如果导出的是图片二进制数据
                    if (excelPropertyInfo.IsImage())
                    {
                        if (value is byte[] imageBytes)
                        {
                            SetImage(excelPropertyInfo.TempCellAddress, imageBytes);
                        }
                    }
                    else
                    {
                        var displayValue = excelPropertyInfo.ExportMappedToDisplay(value);
                        SetValue(excelPropertyInfo.TempCellAddress, displayValue);
                    }
                }
            }
            return this;
        }

        #endregion
    }
}
