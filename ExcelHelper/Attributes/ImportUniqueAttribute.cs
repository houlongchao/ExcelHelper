using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入唯一限制
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ImportUniqueAttribute : Attribute
    {
        /// <summary>
        /// 导入唯一限制联合属性
        /// </summary>
        public string[] UniquePropertites { get; }

        /// <summary>
        /// 导入唯一限制
        /// </summary>
        public ImportUniqueAttribute(params string[] uniquePropertites)
        {
            UniquePropertites = uniquePropertites;
        }
    }
}
