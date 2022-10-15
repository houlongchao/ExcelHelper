using ExcelHelper;
using ExcelHelper.NPOI;
using NUnit.Framework;

namespace ExcelHelperTest
{
    public class ExcelWriterTest_NPOI : ExcelWriterTest
    {
        [SetUp]
        public void Setup()
        {
            _excelHelper = new ExcelHelperBuilder().BuildWrite();
        }

    }
}
