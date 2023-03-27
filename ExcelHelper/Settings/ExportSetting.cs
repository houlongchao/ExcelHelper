using System.Collections.Generic;

namespace ExcelHelper.Settings
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
    }
}
