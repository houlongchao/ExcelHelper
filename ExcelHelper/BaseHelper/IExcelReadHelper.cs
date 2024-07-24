using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 读取帮助类
    /// </summary>
    public interface IExcelReadHelper: IDisposable
    {
        /// <summary>
        /// Excel 文件流
        /// </summary>
        Stream FileStream { get; }

        /// <summary>
        /// 获取所有Sheet名称
        /// </summary>
        /// <returns></returns>
        List<ExcelSheetInfo> GetAllSheets();


        /// <summary>
        /// 导入指定 Sheet 页的数据，如果指定多个 Sheet 页则依次匹配，返回第一个匹配到的 Sheet 页数据 <br/>
        /// 如果没有指定名称，则解析第一个 sheet 页
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        List<T> ImportSheet<T>(params string[] sheetNames) where T: new();

        /// <summary>
        /// 导入指定 Sheet 页的数据，如果指定多个 Sheet 页则依次匹配，返回第一个匹配到的 Sheet 页数据 <br/>
        /// 如果没有指定名称，则解析第一个 sheet 页
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="importSetting"></param>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        List<T> ImportSheet<T>(ImportSetting importSetting = null, params string[] sheetNames) where T : new();

        /// <summary>
        /// 读取一个Sheet页，如果指定多个 Sheet 页则依次匹配，返回第一个匹配到的 Sheet 页数据 <br/>
        /// 如果没有指定名称，则解析第一个 sheet 页
        /// </summary>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        IExcelSheet GetExcelSheet(params string[] sheetNames);
    }
}
