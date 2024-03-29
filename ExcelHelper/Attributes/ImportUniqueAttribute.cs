﻿using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入唯一限制
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ImportUniqueAttribute : Attribute
    {
        /// <summary>
        /// 提示信息
        /// </summary>
        public string Message { get; set; }
    }
}
