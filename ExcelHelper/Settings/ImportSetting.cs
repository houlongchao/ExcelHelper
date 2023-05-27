using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 导入配置
    /// </summary>
    public class ImportSetting : BaseImportSetting
    {
        /// <summary>
        /// 导入头映射 (<c>nameof(A)</c>, <c>title</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>title</c> : Excel列标题</para>
        /// </summary>
        public Dictionary<string, string> TitleMapping { get; private set; } = new Dictionary<string, string>();

        
        #region Add

        /// <summary>
        /// 添加属性与excel头映射
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="title">excel中列标题</param>
        public void AddTitleMapping(string propertyName, string title)
        {
            TitleMapping[propertyName] = title;
        }

        #endregion
    }
}
