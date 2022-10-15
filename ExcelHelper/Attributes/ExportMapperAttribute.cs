using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导出映射设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ExportMapperAttribute : Attribute
    {
        /// <summary>
        /// 真实值
        /// </summary>
        public object Actual { get; set; }

        /// <summary>
        /// 显示值
        /// </summary>
        public string Display { get; set; }

        /// <summary>
        /// 导出映射设置
        /// </summary>
        /// <param name="actual">真实值</param>
        /// <param name="display">显示值</param>
        public ExportMapperAttribute(object actual, string display)
        {
            Actual = actual;
            Display = display;
        }
    }
}
