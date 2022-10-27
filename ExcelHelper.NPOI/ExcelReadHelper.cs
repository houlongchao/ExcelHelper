using NPOI.SS.UserModel;
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
        /// Excel 读取帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        /// <param name="stream"></param>
        public ExcelReadHelper(ExcelHelperBuilder excelHelperBuilder, Stream stream) : base(excelHelperBuilder, stream)
        {
            _excel = NpoiHelper.ReadExcel(FileStream);
        }

        /// <summary>
        /// Excel 读取帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        /// <param name="fileBytes"></param>
        public ExcelReadHelper(ExcelHelperBuilder excelHelperBuilder, byte[] fileBytes) : base(excelHelperBuilder, fileBytes)
        {
            _excel = NpoiHelper.ReadExcel(FileStream);
        }

        /// <summary>
        /// Excel 读取帮助类
        /// </summary>
        /// <param name="excelHelperBuilder"></param>
        /// <param name="filePath"></param>
        public ExcelReadHelper(ExcelHelperBuilder excelHelperBuilder, string filePath) : base(excelHelperBuilder, filePath)
        {
            _excel = NpoiHelper.ReadExcel(FileStream);
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
        public override List<T> ImportSheet<T>(params string[] sheetNames)
        {
            var result = new List<T>();
            var sheet = _excel.GetSheet(sheetNames);
            
            if (sheetNames.Length <= 0)
            {
                sheet = _excel.GetSheetAt(0);
            }
            
            if (sheet == null)
            {
                return result;
            }

            // 获取导入模型属性信息字典
            var excelPropertyInfoNameDict = typeof(T).GetImportNamePropertyInfoDict();

            // 获取导入数据列对应的模型属性
            var excelPropertyInfoIndexDict = new Dictionary<int, ExcelPropertyInfo>();
            var titleRow = sheet.GetRow(0);
            foreach (var titleCell in titleRow)
            {
                var title = titleCell.GetData()?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    continue;
                }
                if (!excelPropertyInfoNameDict.ContainsKey(title))
                {
                    continue;
                }
                excelPropertyInfoIndexDict[titleCell.ColumnIndex] = excelPropertyInfoNameDict[title];
            }

            // 读取数据
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }

                var t = new T();
                foreach (var excelPropertyInfo in excelPropertyInfoIndexDict)
                {
                    var value = row.GetCell(excelPropertyInfo.Key).GetData();

                    excelPropertyInfo.Value.ImportLimit.CheckValue(value);

                    var actualValue = excelPropertyInfo.Value.ImportMappers.MappedToActual(value);

                    excelPropertyInfo.Value.PropertyInfo.SetValueAuto(t, actualValue);
                }
                
                result.Add(t);
            }

            return result;
        }
    }
}
