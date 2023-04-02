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

            var importSetting = new ImportSetting();
            importSetting.AddTitleMapping(nameof(DemoIO.A), "AA");
            importSetting.AddRequiredProperties(nameof(DemoIO.Image));
            importSetting.AddUniqueProperties(nameof(DemoIO.A));
            importSetting.AddLimitValues(nameof(DemoIO.A), "A1", "A2", "A3");
            importSetting.AddValueTrim(nameof(DemoIO.A), Trim.All);

            var sheets2 = _excelHelper.ImportSheet<DemoIO>(importSetting);
            Assert.IsNotNull(sheets2);
        }
    }
}
