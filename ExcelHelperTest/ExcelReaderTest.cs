using ExcelHelper;
using NUnit.Framework;
using System.Collections.Generic;

namespace ExcelHelperTest
{
    /// <summary>
    /// Excel 读取器测试
    /// </summary>
    public abstract class ExcelReaderTest
    {
        protected IExcelReadHelper _excelHelper;

        /// <summary>
        /// 获取所有 Sheet
        /// </summary>
        [Test]
        public void Test_GetAllSheets()
        {
            var sheets = _excelHelper.GetAllSheets();
            Assert.IsNotNull(sheets);
        }

        /// <summary>
        /// 导入指定Sheet页数据
        /// </summary>
        [Test]
        public void Test_ImportSheet()
        {
            var sheets = _excelHelper.ImportSheet<DemoIO>();
            Assert.IsNotNull(sheets);

            var sheets2 = _excelHelper.ImportSheet<DemoIO>(new ImportSetting()
            {
                TitleMapping = new Dictionary<string, string>() { { nameof(DemoIO.A), "AA"} },
                RequiredProperties = new List<string>() { nameof(DemoIO.Image) },
                UniqueProperties = new List<string> { nameof(DemoIO.A) },
            });
            Assert.IsNotNull(sheets2);
        }
    }
}
