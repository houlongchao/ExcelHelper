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

            ExportHeader = propertyInfo.GetCustomAttribute<ExportHeaderAttribute>();
            ExportMappers = propertyInfo.GetCustomAttributes<ExportMapperAttribute>();
            ExportMapperElse = propertyInfo.GetCustomAttribute<ExportMapperElseAttribute>();
            ExportHeader = propertyInfo.GetCustomAttribute<ExportHeaderAttribute>();
            ExportIgnore = propertyInfo.GetCustomAttribute<ExportIgnoreAttribute>();
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
                if (limit.Equals(value))
                {
                    return;
                }
            }

            throw ImportException.New($"【{value}】is limit");
        }

        /// <summary>
        /// 是否是图片
        /// </summary>
        public bool ImportIsImage()
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
        /// 检查必须,如果设置了必须且没有数据则报错
        /// </summary>
        /// <returns></returns>
        public void ImportCheckRequired(object data)
        {
            if (ImportHeaders == null)
            {
                return;
            }

            foreach (var importHeader in ImportHeaders)
            {
                if (string.IsNullOrEmpty(data?.ToString()))
                {
                    if (!string.IsNullOrEmpty(importHeader.RequiredMessage))
                    {
                        throw new ImportException(importHeader.RequiredMessage);
                    }
                    else if (importHeader.IsRequired)
                    {
                        throw new ImportException($"Column【{importHeader.Name}】is Required!");
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
        
        #endregion
    }
}
