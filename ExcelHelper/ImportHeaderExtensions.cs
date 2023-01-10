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

        /// <summary>
        /// 检查必须,如果设置了必须且没有数据则报错
        /// </summary>
        /// <returns></returns>
        public static void CheckRequired(this IEnumerable<ImportHeaderAttribute> importHeaders, object data)
        {
            if (importHeaders == null)
            {
                return;
            }

            foreach (var importHeader in importHeaders)
            {
                if (importHeader.IsRequired && string.IsNullOrEmpty(data?.ToString()))
                {
                    throw new ImportException($"{importHeader.Name} is Required!");
                }
            }
        }
    }
}
