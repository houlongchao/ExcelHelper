using ExcelHelper;
using ExcelHelper.NPOI;
using NUnit.Framework;

namespace ExcelHelperTest
{
    public class ExcelReaderTest_NPOI : ExcelReaderTest
    {
        [SetUp]
        public void Setup()
        {
            _excelHelper = new ExcelHelperBuilder().BuildRead("Excel.xlsx");
        }

    }
}
