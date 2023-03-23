using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导出映射else设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExportMapperElseAttribute : Attribute
    {
        /// <summary>
        /// 显示值
        /// </summary>
        public string Display { get; set; }

        /// <summary>
        /// 导出映射设置
        /// </summary>
        /// <param name="display">显示值</param>
        public ExportMapperElseAttribute(string display)
        {
            Display = display;
        }
    }
}
