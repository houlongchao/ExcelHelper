namespace ExcelHelper
{
    /// <summary>
    /// 导入异常
    /// </summary>
    public class ImportException : ExcelHelperException
    {
        /// <summary>
        /// 导入异常
        /// </summary>
        /// <param name="message"></param>
        public ImportException(string message) : base(message)
        {
        }

        /// <summary>
        /// 创建一个导入异常
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public static ImportException New(string message)
        {
            return new ImportException(message);
        }
    }
}
