using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 导入配置
    /// </summary>
    public class ImportSetting
    {
        /// <summary>
        /// 导入头映射 (<c>nameof(A)</c>, <c>title</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>title</c> : Excel列标题</para>
        /// </summary>
        public Dictionary<string, string> TitleMapping { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 必须有数据的属性
        /// </summary>
        public List<string> RequiredProperties { get; set; } = new List<string>();

        /// <summary>
        /// 唯一性验证的属性
        /// </summary>
        public List<string> UniqueProperties { get; set; } = new List<string>();
    }
}
