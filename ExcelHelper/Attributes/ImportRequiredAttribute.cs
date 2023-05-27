using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入必须
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ImportRequiredAttribute : Attribute
    {
        /// <summary>
        /// 提示信息
        /// </summary>
        public string Message { get; set; }
    }
}
