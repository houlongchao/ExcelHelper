using System;

namespace ExcelHelper
{
    /// <summary>
    /// 单元格数据限制设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExportValidationsAttribute : Attribute
    {
        /// <summary>
        /// 单元格限制设置
        /// </summary>
        public string[] Validations { get; set; }

        /// <summary>
        /// 单元格数据限制设置
        /// </summary>
        public ExportValidationsAttribute(params string[] validations)
        {
            Validations = validations;
        }
    }
}
