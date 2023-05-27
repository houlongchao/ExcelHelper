using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelHelper
{
    /// <summary>
    /// 属性信息扩展方法
    /// </summary>
    public static class PropertyInfoExtensions
    {
        /// <summary>
        /// 获取Excel对象
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static ExcelObjectInfo GetExcelObjectInfo(this Type type)
        {
            return new ExcelObjectInfo(type);
        }

        /// <summary>
        /// 获取Excel属性对象
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        public static ExcelPropertyInfo GetExcelPropertyInfo(this PropertyInfo propertyInfo)
        {
            return new ExcelPropertyInfo(propertyInfo);
        }

        /// <summary>
        /// 获取导入模型属性信息列表
        /// </summary>
        /// <param name="type"></param>
        /// <param name="titleIndexDict"></param>
        /// <param name="importSetting"></param>
        /// <returns></returns>
        public static List<ExcelPropertyInfo> GetImportExcelPropertyInfoList(this Type type, Dictionary<string, int> titleIndexDict, ImportSetting importSetting = null)
        {
            var result = new List<ExcelPropertyInfo>();
            // 获取导入模型属性信息
            var properties = type.GetProperties();
            foreach (var property in properties)
            {
                var excelPropertyInfo = property.GetExcelPropertyInfo();
                excelPropertyInfo.UpdateByImportSetting(importSetting);

                // 如果表头能被识别则加入要读取的列表
                if (excelPropertyInfo.SetImportHeaderColumnIndex(titleIndexDict))
                {
                    result.Add(excelPropertyInfo);
                }
            }

            return result;
        }

        /// <summary>
        /// 设置值，自动转换类型
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <param name="obj"></param>
        /// <param name="value"></param>

        public static void SetValueAuto(this PropertyInfo propertyInfo, object obj, object value)
        {
            if (propertyInfo.PropertyType == typeof(double))
            {
                propertyInfo.SetValue(obj, Convert.ToDouble(value));
            }
            else if (propertyInfo.PropertyType == typeof(int))
            {
                propertyInfo.SetValue(obj, Convert.ToInt32(value));
            }
            else if (propertyInfo.PropertyType == typeof(float))
            {
                propertyInfo.SetValue(obj, Convert.ToDouble(value));
            }
            else if (propertyInfo.PropertyType == typeof(decimal))
            {
                propertyInfo.SetValue(obj, Convert.ToDecimal(value));
            }
            else if (propertyInfo.PropertyType == typeof(DateTime))
            {
                propertyInfo.SetValue(obj, Convert.ToDateTime(value));
            }
            else if (propertyInfo.PropertyType == typeof(string))
            {
                propertyInfo.SetValue(obj, Convert.ToString(value));
            }
            else
            {
                propertyInfo.SetValue(obj, value);
            }
        }

        /// <summary>
        /// 获取导出模型属性信息列表
        /// </summary>
        /// <param name="type"></param>
        /// <param name="exportSetting"></param>
        /// <returns></returns>
        public static List<ExcelPropertyInfo> GetExportExcelPropertyInfoList(this Type type, ExportSetting exportSetting)
        {
            var result = new List<ExcelPropertyInfo>();
            var properties = type.GetProperties();
            foreach (var property in properties)
            {
                var excelPropertyInfo = property.GetExcelPropertyInfo();
                excelPropertyInfo.UpdateByExportSetting(exportSetting);
                if (!excelPropertyInfo.IsExportIgnore())
                {
                    result.Add(excelPropertyInfo);
                }
            }

            return result;
        }

        /// <summary>
        /// 获取模板模型属性信息列表
        /// </summary>
        /// <param name="type"></param>
        /// <param name="tempSetting"></param>
        /// <returns></returns>
        public static List<ExcelPropertyInfo> GetTempExcelPropertyInfoList(this Type type, TempSetting tempSetting = null)
        {
            var result = new List<ExcelPropertyInfo>();
            // 获取导入模型属性信息
            var properties = type.GetProperties();
            foreach (var property in properties)
            {
                var excelPropertyInfo = property.GetExcelPropertyInfo();
                if (excelPropertyInfo != null)
                {
                    excelPropertyInfo.UpdateByTempSetting(tempSetting);
                    if (excelPropertyInfo.HasTempInfo)
                    {
                        excelPropertyInfo.TrimTempChildren();
                        result.Add(excelPropertyInfo);
                    }
                }
            }

            return result;
        }

    }
}
