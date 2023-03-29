using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 导出配置
    /// </summary>
    public class ExportSetting
    {
        /// <summary>
        /// 是否添加列标题
        /// </summary>
        public bool AddTitle { get; set; } = true;

        /// <summary>
        /// 要忽略的属性
        /// </summary>
        public List<string> IgnoreProperties { get; set; } = new List<string>();

        /// <summary>
        /// 要添加的属性
        /// </summary>
        public List<string> IncludeProperties { get; set; } = new List<string>();

        /// <summary>
        /// 导出头映射 (<c>nameof(A)</c>, <c>title</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>title</c> : Excel列标题</para>
        /// </summary>
        public Dictionary<string, string> TitleMapping { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 标题备注 (<c>nameof(A)</c>, <c>comment</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>comment</c> : Excel列标题上的备注</para>
        /// </summary>
        public Dictionary<string, string> TitleComment { get; set; } = new Dictionary<string, string>();
    }
}
