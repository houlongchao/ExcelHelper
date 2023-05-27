using ExcelHelper;
using ExcelHelper.NPOI;
using NUnit.Framework;

namespace ExcelHelperTest
{
    public class ExcelTempTest_NPOI : ExcelTempTest
    {
        [SetUp]
        public void Setup()
        {
            _excelHelper = new ExcelTempHelper();
        }

    }
}
