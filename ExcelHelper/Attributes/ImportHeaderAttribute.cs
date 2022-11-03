using System;
using System.Collections.Generic;

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

    }
}
