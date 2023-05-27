using System;

namespace ExcelHelper
{
    /// <summary>
    /// 导入唯一限制
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public class ImportUniquesAttribute : Attribute
    {
        /// <summary>
        /// 提示信息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 导入唯一限制联合属性
        /// </summary>
        public string[] UniquePropertites { get; }

        /// <summary>
        /// 导入唯一限制
        /// </summary>
        public ImportUniquesAttribute(params string[] uniquePropertites)
        {
            UniquePropertites = uniquePropertites;
        }
    }
}
