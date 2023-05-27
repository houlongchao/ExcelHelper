using System;

namespace ExcelHelper
{
    /// <summary>
    /// 模板列表数据位置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class TempListItemAttribute : Attribute
    {
        /// <summary>
        /// 模板列表数据位置
        /// </summary>
        /// <param name="itemIndex">行/列 索引</param>
        public TempListItemAttribute(int itemIndex)
        {
            ItemIndex = itemIndex;
        }

        /// <summary>
        /// 模板列表数据位置
        /// </summary>
        public int ItemIndex { get; set; } = -1;

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

        /// <summary>
        /// 是否唯一
        /// </summary>
        public bool IsUnique { get; set; } = false;
    }
}
