﻿using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导出头设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExportHeaderAttribute : Attribute
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// 是否自动列宽度
        /// </summary>
        public bool IsAutoSizeColumn { get; set; } = false;

        /// <summary>
        /// 列宽度, <see cref="IsAutoSizeColumn"/>为<c>false</c>时生效
        /// </summary>
        public int ColumnWidth { get; set; } = 20;

        /// <summary>
        /// title 是否加粗
        /// </summary>
        public bool IsBold { get; set; } = true;

        /// <summary>
        /// title 字体大小
        /// </summary>
        public int FontSize { get; set; } = 12;

        /// <summary>
        /// title 颜色
        /// <br/>颜色受限于Excel驱动框架的颜色系统
        /// </summary>
        public string ColorName { get; set; } = "Black";

        /// <summary>
        /// 导出值为空是回退属性值，可以回退多个属性，用;隔开
        /// </summary>
        public string EmptyFallbackPropertyName { get; set; }

        /// <summary>
        /// 导出头设置
        /// </summary>
        /// <param name="name"></param>
        public ExportHeaderAttribute(string name)
        {
            Name = name;
        }
    }
}
