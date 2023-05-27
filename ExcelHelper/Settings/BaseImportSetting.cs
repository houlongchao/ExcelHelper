using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 基础导入设置
    /// </summary>
    public abstract class BaseImportSetting
    {
        /// <summary>
        /// 导入值限制 (<c>nameof(A)</c>, <c>value list</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>value list</c> : 值列表</para>
        /// </summary>
        public Dictionary<string, List<object>> LimitValues { get; private set; } = new Dictionary<string, List<object>>();

        /// <summary>
        /// 导入限制提示
        /// </summary>
        public Dictionary<string, string> LimitMessage { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 值Trim
        /// </summary>
        public Dictionary<string, Trim> ValueTrim { get; private set; } = new Dictionary<string, Trim>();

        /// <summary>
        /// 必须有数据的属性
        /// </summary>
        public List<string> RequiredProperties { get; private set; } = new List<string>();

        /// <summary>
        /// 导入必须提示
        /// </summary>
        public Dictionary<string, string> RequiredMessage { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 唯一性验证的属性
        /// </summary>
        public List<string> UniqueProperties { get; private set; } = new List<string>();

        /// <summary>
        /// 导入唯一性验证提示
        /// </summary>
        public Dictionary<string, string> UniqueMessage { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 导入唯一限制
        /// </summary>
        public List<ImportUniquesAttribute> ImportUniquesAttributes { get; private set; } = new List<ImportUniquesAttribute>();

        #region Add


        /// <summary>
        /// 添加限制值清单
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="values">导入数据限制列表</param>
        public void AddLimitValues(string propertyName, params object[] values)
        {
            if (!LimitValues.ContainsKey(propertyName))
            {
                LimitValues[propertyName] = new List<object>();
            }

            LimitValues[propertyName].AddRange(values);
        }

        /// <summary>
        /// 添加限制值提示信息
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="message">提示信息</param>
        public void AddLimitMessage(string propertyName, string message)
        {
            LimitMessage[propertyName] = message;
        }

        /// <summary>
        /// 添加数据值Trim
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="trim">数据Trim方式</param>
        public void AddValueTrim(string propertyName, Trim trim)
        {
            ValueTrim[propertyName] = trim;
        }

        /// <summary>
        /// 添加必填属性
        /// </summary>
        /// <param name="propertyNames">对象属性名称</param>
        public void AddRequiredProperties(params string[] propertyNames)
        {
            RequiredProperties.AddRange(propertyNames);
        }

        /// <summary>
        /// 添加必填提示信息
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="message">提示信息</param>
        public void AddRequiredMessage(string propertyName, string message)
        {
            RequiredMessage[propertyName] = message;
            RequiredProperties.Add(propertyName);
        }

        /// <summary>
        /// 添加唯一属性
        /// </summary>
        /// <param name="propertyNames">对象属性名称</param>
        public void AddUniqueProperties(params string[] propertyNames)
        {
            UniqueProperties.AddRange(propertyNames);
        }

        /// <summary>
        /// 添加唯一性提示信息
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="message">提示信息</param>
        public void AddUniqueMessage(string propertyName, string message)
        {
            UniqueMessage[propertyName] = message;
            UniqueProperties.Add(propertyName);
        }

        /// <summary>
        /// 添加导入数据唯一限制（针对列表对象，多字段唯一）
        /// </summary>
        /// <param name="importUniquesAttribute"></param>
        public void AddUniquesAttribute(ImportUniquesAttribute importUniquesAttribute)
        {
            ImportUniquesAttributes.Add(importUniquesAttribute);
        }

        #endregion
    }
}
