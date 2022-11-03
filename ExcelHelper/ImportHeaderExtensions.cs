using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 导入头扩展方法
    /// </summary>
    public static class ImportHeaderExtensions
    {
        /// <summary>
        /// 是否是图片
        /// </summary>
        public static bool IsImage(this IEnumerable<ImportHeaderAttribute> importHeaders)
        {
            if (importHeaders == null)
            {
                return false;
            }

            foreach (var importHeader in importHeaders)
            {
                if (importHeader.IsImage)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
