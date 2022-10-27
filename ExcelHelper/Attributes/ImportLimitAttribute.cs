using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入限制
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ImportLimitAttribute : Attribute
    {
        /// <summary>
        /// 导入限制
        /// </summary>
        /// <param name="limits"></param>
        public ImportLimitAttribute(params object[] limits)
        {
            Limits = limits;
        }

        /// <summary>
        /// 导入限制
        /// </summary>
        public object[] Limits { get; private set; }
    }
}
