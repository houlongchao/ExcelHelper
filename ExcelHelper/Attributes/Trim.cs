using System.ComponentModel;

namespace ExcelHelper
{
    /// <summary>
    /// Trim类型
    /// </summary>
    public enum Trim
    {
        /// <summary>
        /// 不做处理
        /// </summary>
        [Description("不做处理")] None = 0,

        /// <summary>
        /// 处理两面
        /// </summary>
        [Description("处理两面")] All = 1,

        /// <summary>
        /// 处理前面
        /// </summary>
        [Description("处理前面")] Start = 2,

        /// <summary>
        /// 处理后面
        /// </summary>
        [Description("处理后面")] End = 3,
    }
}
