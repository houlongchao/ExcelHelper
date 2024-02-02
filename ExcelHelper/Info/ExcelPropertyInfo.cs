using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 属性信息
    /// </summary>
    public class ExcelPropertyInfo
    {
        #region Property

        /// <summary>
        /// 字段属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; }

        /// <summary>
        /// 是否是数组
        /// </summary>
        public bool IsArray { get; }

        #region 导入配置

        /// <summary>
        /// 导入头标题
        /// </summary>
        public string ImportHeaderTitle { get; private set; }

        /// <summary>
        /// 导入标题所在列索引,执行<see cref="SetImportHeaderColumnIndex(Dictionary{string, int})"/>之后有效
        /// </summary>
        public int ImportHeaderColumnIndex { get; private set; }

        /// <summary>
        /// 导入唯一性限制
        /// </summary>
        public bool ImportUnique { get; private set; } = false;

        /// <summary>
        /// 导入唯一性限制提示信息
        /// </summary>
        public string ImportUniqueMessage { get; private set; }

        /// <summary>
        /// 导入数据必须
        /// </summary>
        public bool ImportRequired { get; private set; } = false;

        /// <summary>
        /// 导入数据必须提示信息
        /// </summary>
        public string ImportRequiredMessage { get; private set; }

        /// <summary>
        /// 导入限制
        /// </summary>
        public List<object> ImportLimitValues { get; private set; } = new List<object>();

        /// <summary>
        /// 导入限制提示信息
        /// </summary>
        public string ImportLimitMessage { get; private set; }

        /// <summary>
        /// 导入值Trim
        /// </summary>
        public Trim ImportTrimValue { get; private set; } = Trim.None;

        #endregion

        #region 导出配置

        /// <summary>
        /// 导出头标题
        /// </summary>
        public string ExportHeaderTitle { get; private set; }

        /// <summary>
        /// 导出头备注
        /// </summary>
        public string ExportHeaderComment { get; private set; }

        /// <summary>
        /// 导出忽略
        /// </summary>
        public bool ExportIgnore { get; private set; } = false;

        #endregion

        #region 模板配置

        /// <summary>
        /// 是否设置了模板配置信息
        /// </summary>
        public bool HasTempInfo { get; private set; } = false;

        /// <summary>
        /// 是否设置了列表模板项配置信息
        /// </summary>
        public bool HasTempItemInfo { get; private set; } = false;

        /// <summary>
        /// 模板字段位置
        /// </summary>
        public string TempCellAddress { get; private set; }

        /// <summary>
        /// 模板列表类型
        /// </summary>
        public TempListType TempListType { get; private set; }

        /// <summary>
        /// 模板数据开始
        /// </summary>
        public int TempListStartIndex { get; private set; }

        /// <summary>
        /// 模板数据结束
        /// </summary>
        public int TempListEndIndex { get; private set; }

        /// <summary>
        /// 模板列表数据位置
        /// </summary>
        public int TempListItemIndex { get; private set; }

        /// <summary>
        /// 子属性字段信息
        /// </summary>
        public List<ExcelPropertyInfo> Children { get; set; } = new List<ExcelPropertyInfo>();

        #endregion

        #endregion

        #region Attribute

        #region 公共Attribute

        /// <summary>
        /// 图片属性
        /// </summary>
        public ImageAttribute ImageAttribute { get; set; }

        #endregion

        #region 导入Attribute

        /// <summary>
        /// 导入头
        /// </summary>
        public IEnumerable<ImportHeaderAttribute> ImportHeaderAttributes { get; }

        /// <summary>
        /// 导入映射
        /// </summary>
        public IEnumerable<ImportMapperAttribute> ImportMapperAttributes { get; }

        /// <summary>
        /// 导入映射else
        /// </summary>
        public ImportMapperElseAttribute ImportMapperElseAttribute { get; }

        /// <summary>
        /// 导入限制
        /// </summary>
        public ImportLimitAttribute ImportLimitAttribute { get; }

        /// <summary>
        /// 导入必须
        /// </summary>
        public ImportRequiredAttribute ImportRequiredAttribute { get; }

        /// <summary>
        /// 导入数据Trim
        /// </summary>
        public ImportTrimAttribute ImportTrimAttribute { get; }

        /// <summary>
        /// 导入唯一限制
        /// </summary>
        public ImportUniqueAttribute ImportUniqueAttribute { get; }

        #endregion

        #region 导出Attribute

        /// <summary>
        /// 导出头
        /// </summary>
        public ExportHeaderAttribute ExportHeaderAttribute { get; }

        /// <summary>
        /// 导出单元格验证
        /// </summary>
        public ExportValidationsAttribute ExportValidationsAttribute { get; }

        /// <summary>
        /// 导出映射
        /// </summary>
        public IEnumerable<ExportMapperAttribute> ExportMapperAttributes { get; }
        /// <summary>
        /// 导出映射else
        /// </summary>
        public ExportMapperElseAttribute ExportMapperElseAttribute { get; }

        /// <summary>
        /// 忽略导出，如果为null则导出，不为null则不导出
        /// </summary>
        public ExportIgnoreAttribute ExportIgnoreAttribute { get; }

        /// <summary>
        /// 导出格式化
        /// </summary>
        public ExportFormatAttribute ExportFormatAttribute { get; set; }

        #endregion

        #region 模板Attribute

        /// <summary>
        /// 模板头
        /// </summary>
        public TempAttribute TempAttribute { get; }

        /// <summary>
        /// 模板头-列表数据
        /// </summary>
        public TempListAttribute TempListAttribute { get; set; }

        /// <summary>
        /// 模板头-列表数据项
        /// </summary>
        public TempListItemAttribute TempListItemAttribute { get; set; }

        #endregion

        #endregion

        #region 构造函数

        /// <summary>
        /// Excel 属性信息
        /// </summary>
        /// <param name="propertyInfo"></param>
        public ExcelPropertyInfo(PropertyInfo propertyInfo)
        {
            PropertyInfo = propertyInfo;
            IsArray = propertyInfo.PropertyType.IsGenericType;
            ImportHeaderTitle = propertyInfo.Name;

            ImageAttribute = propertyInfo.GetCustomAttribute<ImageAttribute>();

            ImportHeaderAttributes = propertyInfo.GetCustomAttributes<ImportHeaderAttribute>();
            ImportMapperAttributes = propertyInfo.GetCustomAttributes<ImportMapperAttribute>();
            ImportMapperElseAttribute = propertyInfo.GetCustomAttribute<ImportMapperElseAttribute>();
            ImportLimitAttribute = propertyInfo.GetCustomAttribute<ImportLimitAttribute>();
            ImportRequiredAttribute = propertyInfo.GetCustomAttribute<ImportRequiredAttribute>();
            ImportTrimAttribute = propertyInfo.GetCustomAttribute<ImportTrimAttribute>();
            ImportUniqueAttribute = propertyInfo.GetCustomAttribute<ImportUniqueAttribute>();
            SetImportInfo();

            ExportHeaderAttribute = propertyInfo.GetCustomAttribute<ExportHeaderAttribute>() ?? new ExportHeaderAttribute(null);
            ExportValidationsAttribute = propertyInfo.GetCustomAttribute<ExportValidationsAttribute>();
            ExportMapperAttributes = propertyInfo.GetCustomAttributes<ExportMapperAttribute>();
            ExportMapperElseAttribute = propertyInfo.GetCustomAttribute<ExportMapperElseAttribute>();
            ExportIgnoreAttribute = propertyInfo.GetCustomAttribute<ExportIgnoreAttribute>();
            ExportFormatAttribute = propertyInfo.GetCustomAttribute<ExportFormatAttribute>();
            SetExportInfo();

            if (IsArray)
            {
                TempListAttribute = propertyInfo.GetCustomAttribute<TempListAttribute>();
                if (TempListAttribute != null)
                {
                    foreach (var genericType in propertyInfo.PropertyType.GenericTypeArguments)
                    {
                        var properties = genericType.GetProperties();
                        foreach (var property in properties)
                        {
                            var excelPropertyInfo = property.GetExcelPropertyInfo();
                            Children.Add(excelPropertyInfo);
                        }
                    }
                }
            }
            else
            {
                TempAttribute = propertyInfo.GetCustomAttribute<TempAttribute>();
                TempListItemAttribute = propertyInfo.GetCustomAttribute<TempListItemAttribute>();
            }
            SetTempInfo();
        }

        /// <summary>
        /// 设置导入信息
        /// </summary>
        private void SetImportInfo()
        {
            // 导入唯一限制
            if (ImportUniqueAttribute != null)
            {
                ImportUnique = true;
                ImportUniqueMessage = ImportUniqueAttribute.Message;
            }
            // 导入必须限制
            if (ImportRequiredAttribute != null)
            {
                ImportRequired = true;
                ImportRequiredMessage = ImportRequiredAttribute.Message;
            }
            // 导入值限制
            if (ImportLimitAttribute?.Limits != null)
            {
                foreach (var limit in ImportLimitAttribute.Limits)
                {
                    ImportLimitValues.Add(limit);
                }
            }
            // 导入头Trim
            if (ImportTrimAttribute != null)
            {
                ImportTrimValue = ImportTrimAttribute.Trim;
            }
        }

        /// <summary>
        /// 获取导出头标题
        /// </summary>
        /// <returns></returns>
        private void SetExportInfo()
        {
            // 导出头标题
            if (!string.IsNullOrEmpty(ExportHeaderAttribute?.Name))
            {
                ExportHeaderTitle = ExportHeaderAttribute.Name;
            }
            else
            {
                ExportHeaderTitle = PropertyInfo.Name;
            }

            // 导出头备注
            ExportHeaderComment = ExportHeaderAttribute?.Comment;

            // 导出忽略
            ExportIgnore = ExportIgnoreAttribute != null;
        }

        /// <summary>
        /// 设置模板信息
        /// </summary>
        private void SetTempInfo()
        {
            if (TempAttribute != null)
            {
                TempCellAddress = TempAttribute.CellAddress;
                HasTempInfo = true;
            }
            if (TempListAttribute != null)
            {
                TempListType = TempListAttribute.Type;
                TempListStartIndex = TempListAttribute.StartIndex;
                TempListEndIndex = TempListAttribute.EndIndex;
                HasTempInfo = true;
            }
            if (TempListItemAttribute != null)
            {
                TempListItemIndex = TempListItemAttribute.ItemIndex;
                HasTempItemInfo = true;
            }
        }

        #endregion

        #region Method

        #region 公共

        /// <summary>
        /// 是否是图片
        /// </summary>
        public bool IsImage()
        {
            return ImageAttribute != null;
        }

        #endregion

        #region Export

        #region Mapper

        /// <summary>
        /// 将导出实际值映射为显示数据
        /// </summary>
        /// <param name="actual">实际值</param>
        /// <returns></returns>
        public object ExportMappedToDisplay(object actual)
        {
            if (ExportMapperAttributes != null)
            {
                foreach (var mapper in ExportMapperAttributes)
                {
                    if (CheckExportMapperActual(actual, mapper.Actual))
                    {
                        return mapper.Display;
                    }
                }
            }

            return ExportMappedToElseDisplay(actual);
        }

        private bool CheckExportMapperActual(object actual, object mapperActual)
        {
            if (actual is DateTime dt && dt.Equals(mapperActual))
            {
                return true;
            }
            else if (actual is Boolean b && b.Equals(mapperActual))
            {
                return true;
            }
            else if (actual is double d && d.Equals(Convert.ToDouble(mapperActual)))
            {
                return true;
            }
            else if (actual is float df && df.Equals(Convert.ToDouble(mapperActual)))
            {
                return true;
            }
            else if (actual is decimal dc && dc.Equals(Convert.ToDecimal(mapperActual)))
            {
                return true;
            }
            else if (actual is int di && di.Equals(Convert.ToInt32(mapperActual)))
            {
                return true;
            }
            else if (actual == null)
            {
                if (actual == mapperActual)
                {
                    return true;
                }
            }
            else if (actual.Equals(mapperActual))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 导出映射else
        /// </summary>
        /// <param name="actual"></param>
        /// <returns></returns>
        private object ExportMappedToElseDisplay(object actual)
        {
            if (ExportMapperElseAttribute == null)
            {
                return actual;
            }
            else
            {
                return ExportMapperElseAttribute.Display;
            }
        }

        #endregion

        #region Set

        /// <summary>
        /// 更新导出信息，
        /// 动态设置会覆盖模型静态配置
        /// </summary>
        public void UpdateByExportSetting(ExportSetting exportSetting)
        {
            if (exportSetting == null)
            {
                return;
            }

            // 导出头标题映射
            if (exportSetting.TitleMapping.TryGetValue(PropertyInfo.Name, out var title))
            {
                ExportHeaderTitle = title;
            }

            // 导出头备注
            if (exportSetting.TitleComment.TryGetValue(PropertyInfo.Name, out var comment))
            {
                ExportHeaderComment = comment;
            }

            // 导出时是否被忽略
            // 如果动态配置的导出忽略，则必定忽略
            if (exportSetting.IgnoreProperties.Contains(PropertyInfo.Name))
            {
                ExportIgnore = true;
            }
            // 否则，如果动态配置的导出包含，则必定包含
            else if (exportSetting.IncludeProperties.Contains(PropertyInfo.Name))
            {
                ExportIgnore = false;
            }
        }

        #endregion

        #region Check

        /// <summary>
        /// 是否导出时忽略当前属性
        /// </summary>
        /// <returns></returns>
        public bool IsExportIgnore()
        {
            return ExportIgnore;
        }

        #endregion

        #endregion

        #region Import

        #region Mapper

        /// <summary>
        /// 将显示值映射为实际值
        /// </summary>
        /// <param name="display">显示值</param>
        /// <returns></returns>
        public object ImportMappedToActual(object display)
        {
            if (ImportMapperAttributes != null)
            {
                foreach (var mapper in ImportMapperAttributes)
                {
                    if (CheckImportMapperDisplay(display, mapper.Display))
                    {
                        return mapper.Actual;
                    }
                }
            }

            return ImportMappedToElseActual(display);
        }

        private bool CheckImportMapperDisplay(object display, string mapperDisplay)
        {
            if (display is DateTime dt && mapperDisplay == dt.ToString("yyyy-MM-dd HH:mm:ss"))
            {
                return true;
            }
            else if (display is Boolean b && mapperDisplay == b.ToString().ToUpper())
            {
                return true;
            }
            else if (mapperDisplay == display?.ToString())
            {
                return true;
            }
            return false;
        }

        private object ImportMappedToElseActual(object display)
        {
            if (ImportMapperElseAttribute == null)
            {
                return display;
            }
            else
            {
                return ImportMapperElseAttribute.Actual;
            }
        }

        #endregion

        #region Check

        /// <summary>
        /// 是否必须
        /// </summary>
        /// <returns></returns>
        public bool IsImportRequired()
        {
            return ImportRequired;
        }

        /// <summary>
        /// 是否唯一
        /// </summary>
        /// <returns></returns>
        public bool IsImportUnqiue()
        {
            return ImportUnique;
        }

        /// <summary>
        /// 检查导入限制
        /// </summary>
        /// <param name="value"></param>
        public void ImportCheckLimitValue(object value)
        {
            if (ImportLimitValues.Count <= 0)
            {
                return;
            }

            foreach (var limit in ImportLimitValues)
            {
                if (limit?.ToString() == value?.ToString())
                {
                    return;
                }
            }

            if (!string.IsNullOrEmpty(ImportLimitMessage))
            {
                throw ImportException.New(ImportLimitMessage);
            }
            else
            {
                throw ImportException.New($"【{ImportHeaderTitle}】设置了不被支持的值【{value}】");
            }
        }

        /// <summary>
        /// 检查必须,如果设置了必须且没有数据则报错
        /// </summary>
        /// <returns></returns>
        public void ImportCheckRequired(object data)
        {
            if (IsImportRequired())
            {
                if (string.IsNullOrEmpty(data?.ToString()))
                {
                    if (!string.IsNullOrEmpty(ImportRequiredMessage))
                    {
                        throw ImportException.New(ImportRequiredMessage);
                    }
                    else
                    {
                        throw ImportException.New($"【{ImportHeaderTitle}】是必须的!");
                    }
                }
            }
        }

        /// <summary>
        /// 导入检查字典
        /// </summary>
        private HashSet<string> importUnqiueCheckSet = new HashSet<string>();

        /// <summary>
        /// 导入检查唯一性
        /// </summary>
        /// <param name="actualValue">导入的数据</param>
        /// <exception cref="ImportException"></exception>
        public void ImportCheckUnqiue(object actualValue)
        {
            // 唯一检查
            if (IsImportUnqiue())
            {
                if (importUnqiueCheckSet.Contains(actualValue?.ToString()))
                {
                    if (!string.IsNullOrEmpty(ImportUniqueMessage))
                    {
                        throw ImportException.New(ImportUniqueMessage);
                    }
                    else
                    {
                        throw ImportException.New($"【{ImportHeaderTitle}】存在重复数据：{actualValue}");
                    }
                }
                importUnqiueCheckSet.Add(actualValue?.ToString());
            }
        }

        #endregion

        #region Trim

        /// <summary>
        /// 移除前后空字符串
        /// </summary>
        /// <param name="data"></param>
        public void ImportTrim(ref object data)
        {
            switch (ImportTrimValue)
            {
                case Trim.None:
                    break;
                case Trim.All:
                    data = data?.ToString()?.Trim();
                    break;
                case Trim.Start:
                    data = data?.ToString()?.TrimStart();
                    break;
                case Trim.End:
                    data = data?.ToString()?.TrimEnd();
                    break;
                default:
                    break;
            }
        }

        #endregion

        #region Set

        /// <summary>
        /// 更新导入信息
        /// </summary>
        /// <param name="baseImportSetting"></param>
        private void UpdateByBaseImportSetting(BaseImportSetting baseImportSetting)
        {
            if (baseImportSetting == null)
            {
                return;
            }

            // 导入限制
            if (baseImportSetting.LimitValues.TryGetValue(PropertyInfo.Name, out var values))
            {
                foreach (var value in values)
                {
                    ImportLimitValues.Add(value);
                }
            }
            if (baseImportSetting.LimitMessage.TryGetValue(PropertyInfo.Name, out var limitMessage))
            {
                ImportLimitMessage = limitMessage;
            }
            // 导入值Trim
            if (baseImportSetting.ValueTrim.TryGetValue(PropertyInfo.Name, out var trim))
            {
                ImportTrimValue = trim;
            }

            // 导入唯一性限制
            ImportUnique = baseImportSetting.UniqueProperties.Contains(PropertyInfo.Name);
            if (baseImportSetting.UniqueMessage.TryGetValue(PropertyInfo.Name, out var uniqueMessage))
            {
                ImportUniqueMessage = uniqueMessage;
            }
            // 导入必须限制
            ImportRequired = baseImportSetting.RequiredProperties.Contains(PropertyInfo.Name);
            if (baseImportSetting.RequiredMessage.TryGetValue(PropertyInfo.Name, out var requiredMessage))
            {
                ImportRequiredMessage = requiredMessage;
            }
        }

        /// <summary>
        /// 更新导入信息
        /// </summary>
        /// <param name="importSetting"></param>
        public void UpdateByImportSetting(ImportSetting importSetting)
        {
            if (importSetting == null)
            {
                return;
            }

            // 导入头标题映射
            if (importSetting.TitleMapping.TryGetValue(PropertyInfo.Name, out var title))
            {
                ImportHeaderTitle = title;
            }

            UpdateByBaseImportSetting(importSetting);
        }

        /// <summary>
        /// 设置导入头列索引
        /// </summary>
        /// <param name="titleIndexDict"></param>
        public bool SetImportHeaderColumnIndex(Dictionary<string, int> titleIndexDict)
        {
            // 从excel标题列表中获取到了导入标题，直接设置对应的列索引
            // 此时的导入标题是从动态导入配置中获取
            if (!string.IsNullOrEmpty(ImportHeaderTitle) && titleIndexDict.TryGetValue(ImportHeaderTitle, out var index))
            {
                ImportHeaderColumnIndex = index;
                return true;
            }

            // 识别模型上的导入头设置
            foreach (var importHeader in ImportHeaderAttributes)
            {
                if (titleIndexDict.ContainsKey(importHeader.Name))
                {
                    ImportHeaderTitle = importHeader.Name;
                    ImportHeaderColumnIndex = titleIndexDict[importHeader.Name];
                    return true;
                }
            }

            // 从属性自身识别
            if (titleIndexDict.ContainsKey(PropertyInfo.Name))
            {
                ImportHeaderTitle = PropertyInfo.Name;
                ImportHeaderColumnIndex = titleIndexDict[PropertyInfo.Name];
                return true;
            }

            return false;
        }

        #endregion

        #endregion

        #region Temp

        #region Set

        /// <summary>
        /// 更新模板信息，
        /// 动态设置会覆盖模型静态配置
        /// </summary>
        public void UpdateByTempSetting(TempSetting tempSetting)
        {
            if (tempSetting == null)
            {
                return;
            }
            if (IsArray)
            {
                if (tempSetting.ListSettings.TryGetValue(PropertyInfo.Name, out var tempListSetting))
                {
                    TempListType = tempListSetting.Type;
                    TempListStartIndex = tempListSetting.StartIndex;
                    TempListEndIndex = tempListSetting.EndIndex;
                    HasTempInfo = true;
                    foreach (var child in Children)
                    {
                        child.UpdateByTempListSetting(tempListSetting);
                        child.UpdateImportHeaderTitle($"{ImportHeaderTitle}.{child.ImportHeaderTitle}");
                    }
                }
            }
            else
            {
                if (tempSetting.CellAddress.TryGetValue(PropertyInfo.Name, out var cellAddress))
                {
                    TempCellAddress = cellAddress;
                    HasTempInfo = true;
                }

            }

            UpdateByBaseImportSetting(tempSetting);
        }

        private void UpdateByTempListSetting(TempListSetting tempListSetting)
        {
            if (tempListSetting == null)
            {
                return;
            }
            if (tempListSetting.ItemIndexs.TryGetValue(PropertyInfo.Name, out var itemIndex))
            {
                TempListItemIndex = itemIndex;
                HasTempItemInfo = true;
            }

            UpdateByBaseImportSetting(tempListSetting);
        }

        private void UpdateImportHeaderTitle(string title)
        {
            ImportHeaderTitle = title;
        }

        /// <summary>
        /// 裁剪掉没有设置模板配置的子项
        /// </summary>
        public void TrimTempChildren()
        {
            var needRemoves = new List<ExcelPropertyInfo>();
            foreach (var child in Children)
            {
                if (!child.HasTempItemInfo)
                {
                    needRemoves.Add(child);
                }
            }
            foreach (var item in needRemoves)
            {
                Children.Remove(item);
            }
        }

        #endregion

        #endregion

        #endregion
    }
}
