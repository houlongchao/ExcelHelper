using System.IO;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// ExcelHelperBuilder 扩展方法
    /// </summary>
    public static partial class ExcelHelperBuilderExtension
    {
        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static IExcelReadHelper BuildRead(this ExcelHelperBuilder builder, string filePath)
        {
           return new ExcelReadHelper(builder, filePath);
        }

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="fileBytes"></param>
        /// <returns></returns>
        public static IExcelReadHelper BuildRead(this ExcelHelperBuilder builder, byte[] fileBytes)
        {
            return new ExcelReadHelper(builder, fileBytes);
        }

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static IExcelReadHelper BuildRead(this ExcelHelperBuilder builder, Stream stream)
        {
            return new ExcelReadHelper(builder, stream);
        }
    }

    public static partial class ExcelHelperBuilderExtension
    {
        /// <summary>
        /// 构建Excel文件
        /// </summary>
        /// <param name="builder"></param>
        /// <returns></returns>
        public static IExcelWriteHelper BuildWrite(this ExcelHelperBuilder builder)
        {
            return new ExcelWriteHelper(builder);
        }
    }
}
