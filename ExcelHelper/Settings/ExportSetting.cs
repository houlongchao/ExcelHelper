using ExcelHelper.Settings;
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
        /// 导出位置
        /// </summary>
        public ExportLocation ExportLocation { get; set; } = ExportLocation.LastRow;

        /// <summary>
        /// 导出行位置坐标
        /// </summary>
        public int ExportRowIndex { get; set; }

        /// <summary>
        /// 导出列位置坐标
        /// </summary>
        public int ExportColumnIndex { get; set; }

        /// <summary>
        /// 导出头映射 (<c>nameof(A)</c>, <c>title</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>title</c> : Excel列标题</para>
        /// </summary>
        public Dictionary<string, string> TitleMapping { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 标题备注 (<c>nameof(A)</c>, <c>comment</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>comment</c> : Excel列标题上的备注</para>
        /// </summary>
        public Dictionary<string, string> TitleComment { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 要忽略的属性
        /// </summary>
        public List<string> IgnoreProperties { get; private set; } = new List<string>();

        /// <summary>
        /// 要添加的属性
        /// </summary>
        public List<string> IncludeProperties { get; private set; } = new List<string>();


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

        /// <summary>
        /// 添加excel标题备注信息
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="comment">excel标题备注信息</param>
        public void AddTitleComment(string propertyName, string comment)
        {
            TitleComment[propertyName] = comment;
        }

        /// <summary>
        /// 添加导出时要忽略的属性
        /// </summary>
        /// <param name="propertyNames">对象属性名称</param>
        public void AddIgnoreProperties(params string[] propertyNames)
        {
            IgnoreProperties.AddRange(propertyNames);
        }

        /// <summary>
        /// 指定导出时要导出的属性
        /// </summary>
        /// <param name="propertyNames">对象属性名称</param>
        public void AddIncludeProperties(params string[] propertyNames)
        {
            IncludeProperties.AddRange(propertyNames);
        }

        #endregion
    }
}
