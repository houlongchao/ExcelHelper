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
        /// 获取导入Excel属性对象
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        public static ExcelPropertyInfo GetImportExcelPropertyInfo(this PropertyInfo propertyInfo)
        {
            return new ExcelPropertyInfo(propertyInfo);
        }

        /// <summary>
        /// 获取导入模型属性信息字典
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Dictionary<string, ExcelPropertyInfo> GetImportNamePropertyInfoDict(this Type type)
        {
            // 获取导入模型属性信息
            var properties = type.GetProperties();
            var excelPropertyInfoNameDict = new Dictionary<string, ExcelPropertyInfo>();
            foreach (var property in properties)
            {
                var excelPropertyInfo = property.GetImportExcelPropertyInfo();

                foreach (var importHeader in excelPropertyInfo.ImportHeaders)
                {
                    excelPropertyInfoNameDict.Add(importHeader.Name, excelPropertyInfo);
                }

                if (!excelPropertyInfoNameDict.ContainsKey(property.Name))
                {
                    excelPropertyInfoNameDict.Add(property.Name, excelPropertyInfo);
                }
            }

            return excelPropertyInfoNameDict;
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
        /// 获取导出Excel属性对象
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        public static ExcelPropertyInfo GetExportExcelPropertyInfo(this PropertyInfo propertyInfo)
        {
            return new ExcelPropertyInfo(propertyInfo);
        }

        /// <summary>
        /// 获取导出模型属性信息字典
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Dictionary<string, ExcelPropertyInfo> GetExportNamePropertyInfoDict(this Type type)
        {
            var properties = type.GetProperties();
            var excelPropertyInfoNameDict = new Dictionary<string, ExcelPropertyInfo>();
            foreach (var property in properties)
            {
                var excelPropertyInfo = property.GetExportExcelPropertyInfo();

                // 忽略当前属性导出
                if (excelPropertyInfo.ExportIgnore != null)
                {
                    continue;
                }

                if (string.IsNullOrEmpty(excelPropertyInfo.ExportHeader?.Name))
                {
                    excelPropertyInfoNameDict.Add(property.Name, excelPropertyInfo);
                }
                else
                {
                    excelPropertyInfoNameDict.Add(excelPropertyInfo.ExportHeader.Name, excelPropertyInfo);
                }
            }

            return excelPropertyInfoNameDict;
        }
    }
}
