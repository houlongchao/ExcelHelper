using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入映射else设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ImportMapperElseAttribute : Attribute
    {
        /// <summary>
        /// 真实值
        /// </summary>
        public object Actual { get; set; }

        /// <summary>
        /// 导入映射设置
        /// </summary>
        /// <param name="actual">真实值</param>
        public ImportMapperElseAttribute(object actual)
        {
            Actual = actual;
        }
    }
}
