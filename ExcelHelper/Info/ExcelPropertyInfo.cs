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
        /// <summary>
        /// 字段属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; }

        #region 导入

        /// <summary>
        /// 导入头标题
        /// </summary>
        public string ImportHeaderTitle { get; private set; }

        /// <summary>
        /// 导入标题所在列索引
        /// </summary>
        public int ImportHeaderColumnIndex { get; private set; }

        /// <summary>
        /// 导入唯一性限制
        /// </summary>
        public bool ImportUnique { get; private set; } = false;

        /// <summary>
        /// 导入数据必须
        /// </summary>
        public bool ImportRequired { get; private set; } = false;

        /// <summary>
        /// 导入数据必须提示信息
        /// </summary>
        public string ImportRequiredMessage { get; private set; }

        /// <summary>
        /// 导入头
        /// </summary>
        public IEnumerable<ImportHeaderAttribute> ImportHeaders { get; }

        /// <summary>
        /// 导入映射
        /// </summary>
        public IEnumerable<ImportMapperAttribute> ImportMappers { get; }

        /// <summary>
        /// 导入映射else
        /// </summary>
        public ImportMapperElseAttribute ImportMapperElse { get; }

        /// <summary>
        /// 导入限制
        /// </summary>
        public ImportLimitAttribute ImportLimit { get; }

        #endregion

        #region 导出

        /// <summary>
        /// 导出头标题
        /// </summary>
        public string ExportHeaderTitle { get; private set; }

        /// <summary>
        /// 导出头备注
        /// </summary>
        public string ExportHeaderComment { get; private set; }

        /// <summary>
        /// 导出头
        /// </summary>
        public ExportHeaderAttribute ExportHeader { get; }

        /// <summary>
        /// 导出映射
        /// </summary>
        public IEnumerable<ExportMapperAttribute> ExportMappers { get; }

        /// <summary>
        /// 导出映射else
        /// </summary>
        public ExportMapperElseAttribute ExportMapperElse { get; }

        /// <summary>
        /// 忽略导出，如果为null则导出，不为null则不导出
        /// </summary>
        public ExportIgnoreAttribute ExportIgnore { get; }

        /// <summary>
        /// 是否忽略导出
        /// </summary>
        public bool IsIgnoreExport => ExportIgnore != null;

        #endregion

        /// <summary>
        /// Excel 属性信息
        /// </summary>
        /// <param name="propertyInfo"></param>
        public ExcelPropertyInfo(PropertyInfo propertyInfo)
        {
            PropertyInfo = propertyInfo;

            ImportHeaders = propertyInfo.GetCustomAttributes<ImportHeaderAttribute>();
            ImportMappers = propertyInfo.GetCustomAttributes<ImportMapperAttribute>();
            ImportMapperElse = propertyInfo.GetCustomAttribute<ImportMapperElseAttribute>();
            ImportLimit = propertyInfo.GetCustomAttribute<ImportLimitAttribute>();

            ExportHeader = propertyInfo.GetCustomAttribute<ExportHeaderAttribute>() ?? new ExportHeaderAttribute(null);
            ExportMappers = propertyInfo.GetCustomAttributes<ExportMapperAttribute>();
            ExportMapperElse = propertyInfo.GetCustomAttribute<ExportMapperElseAttribute>();
            ExportIgnore = propertyInfo.GetCustomAttribute<ExportIgnoreAttribute>();
            SetExportHeaderInfo();
        }

        /// <summary>
        /// 获取导出头标题
        /// </summary>
        /// <returns></returns>
        private void SetExportHeaderInfo()
        {
            // 导出头标题
            if (!string.IsNullOrEmpty(ExportHeader?.Name))
            {
                ExportHeaderTitle = ExportHeader.Name;
            }
            else
            {
                ExportHeaderTitle = PropertyInfo.Name;
            }

            // 导出头备注
            ExportHeaderComment = ExportHeader?.Comment;
        }

        #region Export

        /// <summary>
        /// 将导出实际值映射为显示数据
        /// </summary>
        /// <param name="actual">实际值</param>
        /// <returns></returns>
        public object ExportMappedToDisplay(object actual)
        {
            if (ExportMappers == null)
            {
                return ExportMappedToElseDisplay(actual);
            }

            foreach (var mapper in ExportMappers)
            {
                if (actual is DateTime dt && dt.Equals(mapper.Actual))
                {
                    return mapper.Display;
                }
                else if (actual is Boolean b && b.Equals(mapper.Actual))
                {
                    return mapper.Display;
                }
                else if (actual is double d && d.Equals(Convert.ToDouble(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual is float df && df.Equals(Convert.ToDouble(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual is decimal dc && dc.Equals(Convert.ToDecimal(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual is int di && di.Equals(Convert.ToInt32(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual == null)
                {
                    if (actual == mapper.Actual)
                    {
                        return mapper.Display;
                    }
                }
                else if (actual.Equals(mapper.Actual))
                {
                    return mapper.Display;
                }
            }

            return ExportMappedToElseDisplay(actual);
        }

        /// <summary>
        /// 导出映射else
        /// </summary>
        /// <param name="actual"></param>
        /// <returns></returns>
        private object ExportMappedToElseDisplay(object actual)
        {
            if (ExportMapperElse == null)
            {
                return actual;
            }
            else
            {
                return ExportMapperElse.Display;
            }
        }

        /// <summary>
        /// 更新导出信息
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
        }

        /// <summary>
        /// 是否是图片
        /// </summary>
        public bool IsExportImage()
        {
            return ExportHeader != null && ExportHeader.IsImage;
        }


        #endregion

        #region Import

        /// <summary>
        /// 将显示值映射为实际值
        /// </summary>
        /// <param name="display">显示值</param>
        /// <returns></returns>
        public object ImportMappedToActual(object display)
        {
            if (ImportMappers == null)
            {
                return ImportMappedToElseActual(display);
            }

            foreach (var mapper in ImportMappers)
            {
                if (display is DateTime dt && mapper.Display == dt.ToString("yyyy-MM-dd HH:mm:ss"))
                {
                    return mapper.Actual;
                }
                else if (display is Boolean b && mapper.Display == b.ToString().ToUpper())
                {
                    return mapper.Actual;
                }
                else if (mapper.Display == display?.ToString())
                {
                    return mapper.Actual;
                }
            }

            return ImportMappedToElseActual(display);
        }

        private object ImportMappedToElseActual(object display)
        {
            if (ImportMapperElse == null)
            {
                return display;
            }
            else
            {
                return ImportMapperElse.Actual;
            }
        }

        /// <summary>
        /// 检查导入限制
        /// </summary>
        /// <param name="value"></param>
        public void ImportLimitCheckValue(object value)
        {
            if (ImportLimit == null || ImportLimit.Limits == null || ImportLimit.Limits.Length <= 0)
            {
                return;
            }

            foreach (var limit in ImportLimit.Limits)
            {
                if (limit?.ToString() == value?.ToString())
                {
                    return;
                }
            }

            throw ImportException.New($"列【{ImportHeaderTitle}】值【{value}】不被支持");
        }

        /// <summary>
        /// 是否是图片
        /// </summary>
        public bool IsImportImage()
        {
            if (ImportHeaders == null)
            {
                return false;
            }

            foreach (var importHeader in ImportHeaders)
            {
                if (importHeader.IsImage)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 是否必须
        /// </summary>
        /// <returns></returns>
        public bool IsImportRequired()
        {
            if (ImportRequired)
            {
                return true;
            }

            if (ImportHeaders == null)
            {
                return false;
            }

            foreach (var importHeader in ImportHeaders)
            {
                if (!string.IsNullOrEmpty(importHeader.RequiredMessage))
                {
                    ImportRequired = true;
                    ImportRequiredMessage = importHeader.RequiredMessage;
                    return true;
                }
                else if (importHeader.IsRequired)
                {
                    ImportRequired = true;
                    return true;
                }
            }

            return false;
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
                    if (string.IsNullOrEmpty(ImportRequiredMessage))
                    {
                        throw new ImportException($"【{ImportHeaderTitle}】是必须的!");
                    }
                    else
                    {
                        throw new ImportException(ImportRequiredMessage);
                    }
                }
            }
        }

        /// <summary>
        /// 移除前后空字符串
        /// </summary>
        /// <param name="data"></param>
        public void ImportTrim(ref object data)
        {
            if (ImportHeaders == null || data == null)
            {
                return;
            }

            foreach (var importHeader in ImportHeaders)
            {
                switch (importHeader.Trim)
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
        }
        
        /// <summary>
        /// 导入检查字典
        /// </summary>
        private HashSet<string> importCheckSet = new HashSet<string>();

        /// <summary>
        /// 是否唯一
        /// </summary>
        /// <returns></returns>
        public bool IsImportUnqiue()
        {
            if (ImportUnique)
            {
                return true;
            }

            if (ImportHeaders == null)
            {
                return false;
            }

            foreach (var header in ImportHeaders)
            {
                if (header.IsUnique)
                {
                    ImportUnique = true;
                    return true;
                }
            }

            return false;
        }

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
                if (importCheckSet.Contains(actualValue?.ToString()))
                {
                    throw new ImportException($"【{ImportHeaderTitle}】列存在重复数据：{actualValue}");
                }
                importCheckSet.Add(actualValue?.ToString());
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

            // 导入唯一性限制
            ImportUnique = importSetting.UniqueProperties.Contains(PropertyInfo.Name);
            // 导入必须限制
            ImportRequired = importSetting.RequiredProperties.Contains(PropertyInfo.Name);
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
            foreach (var importHeader in ImportHeaders)
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
    }
}
