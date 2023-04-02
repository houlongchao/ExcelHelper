using System.Collections.Generic;
using System.Linq;

namespace ExcelHelper
{
    /// <summary>
    /// 导入配置
    /// </summary>
    public class ImportSetting
    {
        /// <summary>
        /// 导入头映射 (<c>nameof(A)</c>, <c>title</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>title</c> : Excel列标题</para>
        /// </summary>
        public Dictionary<string, string> TitleMapping { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 导入值限制 (<c>nameof(A)</c>, <c>value list</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>value list</c> : 值列表</para>
        /// </summary>
        public Dictionary<string, List<object>> LimitValues { get; private set; } = new Dictionary<string, List<object>>();

        /// <summary>
        /// 值Trim
        /// </summary>
        public Dictionary<string, Trim> ValueTrim { get; private set; } = new Dictionary<string, Trim>();

        /// <summary>
        /// 必须有数据的属性
        /// </summary>
        public List<string> RequiredProperties { get; private set; } = new List<string>();

        /// <summary>
        /// 唯一性验证的属性
        /// </summary>
        public List<string> UniqueProperties { get; private set; } = new List<string>();

        #region Set

        /// <summary>
        /// 添加属性与excel头映射
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="title">excel中列标题</param>
        public void AddTitleMapping(string propertyName, string title)
        {
            TitleMapping[propertyName] = title;
        }

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
        /// 添加唯一属性
        /// </summary>
        /// <param name="propertyNames">对象属性名称</param>
        public void AddUniqueProperties(params string[] propertyNames)
        {
            UniqueProperties.AddRange(propertyNames);
        }

        #endregion
    }
}
