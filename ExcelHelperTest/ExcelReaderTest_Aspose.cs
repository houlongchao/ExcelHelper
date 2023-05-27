using ExcelHelper.Aspose;
using NUnit.Framework;

namespace ExcelHelperTest
{
    public class ExcelReaderTest_Aspose : ExcelReaderTest
    {
        [SetUp]
        public void Setup()
        {
            _excelHelper = new ExcelReadHelper("Excel.xlsx");
        }

    }
}
