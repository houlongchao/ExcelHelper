using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入时对字符串进行Trim
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ImportTrimAttribute : Attribute
    {
        /// <summary>
        /// 字符串Trim
        /// </summary>
        public Trim Trim { get; set; } = Trim.None;

        /// <summary>
        /// 导入时对字符串进行Trim
        /// </summary>
        public ImportTrimAttribute(Trim trim = Trim.None)
        {
            Trim = trim;
        }
    }
}
