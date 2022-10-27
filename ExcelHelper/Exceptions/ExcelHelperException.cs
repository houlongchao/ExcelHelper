using System;

namespace ExcelHelper
{
    /// <summary>
    /// ExcelHelper异常
    /// </summary>
    public class ExcelHelperException : Exception
    {
        /// <summary>
        ///  ExcelHelper异常
        /// </summary>
        /// <param name="message"></param>
        public ExcelHelperException(string message) : base(message)
        {
        }
    }
}
