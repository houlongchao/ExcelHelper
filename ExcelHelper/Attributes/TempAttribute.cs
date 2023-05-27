using System;

namespace ExcelHelper
{
    /// <summary>
    /// 模板设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class TempAttribute : Attribute
    {
        /// <summary>
        /// 模板头设置
        /// </summary>
        /// <param name="cellAddress">单元格位置，如：<c>A11</c></param>
        public TempAttribute(string cellAddress)
        {
            CellAddress = cellAddress;
        }

        /// <summary>
        /// 单元格位置
        /// </summary>
        public string CellAddress { get; set; }

        /// <summary>
        /// 是否是图片数据
        /// </summary>
        public bool IsImage { get; set; } = false;

        /// <summary>
        /// 是否必须
        /// </summary>
        public bool IsRequired { get; set; } = false;

        /// <summary>
        /// 必须提示消息
        /// </summary>
        public string RequiredMessage { get; set; }

        /// <summary>
        /// 字符串Trim
        /// </summary>
        public Trim Trim { get; set; } = Trim.None;
    }
}
