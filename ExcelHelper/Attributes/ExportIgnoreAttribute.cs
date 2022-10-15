using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导出忽略
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExportIgnoreAttribute : Attribute
    {
    }
}
