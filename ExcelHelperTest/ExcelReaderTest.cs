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
            importSetting.AddRequiredProperties(nameof(DemoIO.A));
            importSetting.AddRequiredMessage(nameof(DemoIO.A), "AA是必须的");
            importSetting.AddUniqueProperties(nameof(DemoIO.A));
            importSetting.AddUniqueMessage(nameof(DemoIO.A), "AA必须唯一");
            importSetting.AddLimitValues(nameof(DemoIO.A), "A1", "A2", "A3");
            importSetting.AddLimitMessage(nameof(DemoIO.A), "AA数据非法");
            importSetting.AddValueTrim(nameof(DemoIO.A), Trim.All);
            importSetting.AddTitleMapping("BB", "B");

            var sheets2 = _excelHelper.ImportSheet<DemoIO>(importSetting);
            Assert.AreEqual(3, sheets2.Count);
            Assert.IsNotNull(sheets2);

            var sheets3 = _excelHelper.ImportSheet<Dictionary<string, object>>(importSetting);
            Assert.AreEqual(3, sheets3.Count);
            Assert.IsNotNull(sheets3);
        }
    }
}
