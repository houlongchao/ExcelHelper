using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入映射设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ImportMapperAttribute : Attribute
    {
        /// <summary>
        /// 显示值
        /// </summary>
        public string Display { get; set; }

        /// <summary>
        /// 真实值
        /// </summary>
        public object Actual { get; set; }

        /// <summary>
        /// 导入映射设置
        /// </summary>
        /// <param name="display">显示值</param>
        /// <param name="actual">真实值</param>
        public ImportMapperAttribute(string display, object actual)
        {
            Display = display;
            Actual = actual;
        }
    }
}
