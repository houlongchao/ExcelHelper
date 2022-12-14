using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 帮助类基类
    /// </summary>
    public abstract class BaseExcelReadHelper : IDisposable, IExcelReadHelper
    {
        private readonly Stream _excelStream;

        /// <summary>
        /// Excel 文件流
        /// </summary>
        public Stream FileStream => _excelStream;

        private readonly ExcelHelperBuilder _excelHelperBuilder;

        /// <summary>
        /// Excel 帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        /// <param name="stream">Excel 文件流</param>
        public BaseExcelReadHelper(ExcelHelperBuilder excelHelperBuilder, Stream stream)
        {
            _excelStream = stream;
            _excelHelperBuilder = excelHelperBuilder;
        }

        /// <summary>
        /// Excel 帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        /// <param name="fileBytes">Excel 文件字节数据</param>
        public BaseExcelReadHelper(ExcelHelperBuilder excelHelperBuilder, byte[] fileBytes) : this(excelHelperBuilder, new MemoryStream(fileBytes))
        {
        }

        /// <summary>
        /// Excel 帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        /// <param name="filePath">Excel 文件路径</param>
        public BaseExcelReadHelper(ExcelHelperBuilder excelHelperBuilder, string filePath) : this(excelHelperBuilder, File.ReadAllBytes(filePath))
        {
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public virtual void Dispose()
        {
            _excelStream.Dispose();
            _excelHelperBuilder.Dispose();
        }

        /// <summary>
        /// 获取所有Sheet信息
        /// </summary>
        /// <returns></returns>
        public abstract List<ExcelSheetInfo> GetAllSheets();

        /// <summary>
        /// 导入 Sheet 信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        public List<T> ImportSheet<T>(params string[] sheetNames) where T: new()
        {
            return GetExcelSheet(sheetNames)?.GetData<T>();
        }

        /// <inheritdoc/>
        public abstract IExcelSheet GetExcelSheet(params string[] sheetNames);
    }
}
