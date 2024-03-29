﻿using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// Excel 读取帮助类
    /// </summary>
    public class ExcelReadHelper : BaseExcelReadHelper
    {
        private readonly IWorkbook _excel;

        /// <summary>
        /// NPOI IWorkbook
        /// </summary>
        public IWorkbook Excel => _excel;

        /// <summary>
        /// Excel 读取帮助类
        /// </summary>
        /// <param name="stream"></param>
        public ExcelReadHelper(Stream stream) : base(stream)
        {
            _excel = NpoiHelper.ReadExcel(FileStream);
        }

        /// <summary>
        /// Excel 读取帮助类
        /// </summary>
        /// <param name="fileBytes"></param>
        public ExcelReadHelper(byte[] fileBytes) : base(fileBytes)
        {
            _excel = NpoiHelper.ReadExcel(FileStream);
        }

        /// <summary>
        /// Excel 读取帮助类
        /// </summary>
        /// <param name="filePath"></param>
        public ExcelReadHelper(string filePath) : base(filePath)
        {
            _excel = NpoiHelper.ReadExcel(FileStream);
        }

        /// <inheritdoc/>
        public override void Dispose()
        {
            _excel.Close();
            base.Dispose();
        }

        /// <inheritdoc/>
        public override List<ExcelSheetInfo> GetAllSheets()
        {
            var result = new List<ExcelSheetInfo>();
            for (int i = 0; i < _excel.NumberOfSheets; i++)
            {
                result.Add(new ExcelSheetInfo(i, _excel.GetSheetName(i), _excel.IsSheetHidden(i)));
            }
            return result;
        }

        /// <inheritdoc/>
        public override IExcelSheet GetExcelSheet(params string[] sheetNames)
        {
            var sheet = _excel.GetSheet(sheetNames);

            if (sheetNames.Length <= 0)
            {
                sheet = _excel.GetSheetAt(0);
            }

            if (sheet == null)
            {
                return null;
            }

            return new ExcelSheet(sheet);
        }
    }
}
