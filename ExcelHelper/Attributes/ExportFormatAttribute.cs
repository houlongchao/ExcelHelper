using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导出格式化设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExportFormatAttribute : Attribute
    {
        /// <summary>
        /// 格式化字符串
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 导出格式化设置
        /// </summary>
        public ExportFormatAttribute(string format)
        {
            Format = format;
        }
    }
}
