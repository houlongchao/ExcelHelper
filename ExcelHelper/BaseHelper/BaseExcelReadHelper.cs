using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 帮助类基类
    /// </summary>
    public abstract class BaseExcelReadHelper : IExcelReadHelper
    {
        private readonly Stream _excelStream;

        /// <summary>
        /// Excel 文件流
        /// </summary>
        public Stream FileStream => _excelStream;

        /// <summary>
        /// Excel 帮助类
        /// </summary>
        /// <param name="stream">Excel 文件流</param>
        public BaseExcelReadHelper(Stream stream)
        {
            _excelStream = stream;
        }

        /// <summary>
        /// Excel 帮助类
        /// </summary>
        /// <param name="fileBytes">Excel 文件字节数据</param>
        public BaseExcelReadHelper(byte[] fileBytes) : this(new MemoryStream(fileBytes))
        {
        }

        /// <summary>
        /// Excel 帮助类
        /// </summary>
        /// <param name="filePath">Excel 文件路径</param>
        public BaseExcelReadHelper(string filePath) : this(File.ReadAllBytes(filePath))
        {
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public virtual void Dispose()
        {
            _excelStream.Dispose();
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
            var result = GetExcelSheet(sheetNames)?.GetData<T>();

            return result ?? new List<T>();
        }

        /// <summary>
        /// 导入 Sheet 信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="importSetting"></param>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        public List<T> ImportSheet<T>(ImportSetting importSetting = null, params string[] sheetNames) where T : new()
        {
            var result = GetExcelSheet(sheetNames)?.GetData<T>(importSetting);

            return result ?? new List<T>();
        }

        /// <inheritdoc/>
        public abstract IExcelSheet GetExcelSheet(params string[] sheetNames);
    }
}
