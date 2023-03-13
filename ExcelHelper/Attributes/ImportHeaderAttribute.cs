using System;
using System.ComponentModel;

namespace ExcelHelper
{
    /// <summary>
    /// 导入头设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ImportHeaderAttribute : Attribute
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 导入头设置
        /// </summary>
        /// <param name="name"></param>
        public ImportHeaderAttribute(string name)
        {
            Name = name;
        }

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

    /// <summary>
    /// Trim类型
    /// </summary>
    public enum Trim
    {
        /// <summary>
        /// 不做处理
        /// </summary>
        [Description("不做处理")] None = 0,

        /// <summary>
        /// 处理两面
        /// </summary>
        [Description("处理两面")] All = 1,

        /// <summary>
        /// 处理前面
        /// </summary>
        [Description("处理前面")] Start = 2,

        /// <summary>
        /// 处理后面
        /// </summary>
        [Description("处理后面")] End = 3,
    }
}
