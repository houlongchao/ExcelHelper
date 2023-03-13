namespace ExcelHelper
{
    /// <summary>
    /// 导入限制扩展方法
    /// </summary>
    public static class ImportLimitExtensions
    {
        /// <summary>
        /// 检查导入限制
        /// </summary>
        /// <param name="importLimit"></param>
        /// <param name="value"></param>
        public static void CheckValue(this ImportLimitAttribute importLimit, object value)
        {
            if (importLimit == null || importLimit.Limits == null || importLimit.Limits.Length <= 0)
            {
                return;
            }

            foreach (var limit in importLimit.Limits)
            {
                if (limit.Equals(value))
                {
                    return;
                }
            }

            throw ImportException.New($"【{value}】is limit");
        }
    }
}
