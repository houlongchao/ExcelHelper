namespace ExcelHelper
{
    /// <summary>
    /// 导出异常
    /// </summary>
    public class ExportException : ExcelHelperException
    {
        /// <summary>
        /// 导出异常
        /// </summary>
        /// <param name="message"></param>
        public ExportException(string message) : base(message)
        {
        }

        /// <summary>
        /// 创建一个导出异常
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public static ExportException New(string message)
        {
            return new ExportException(message);
        }
    }
}
