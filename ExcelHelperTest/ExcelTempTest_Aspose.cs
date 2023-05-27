using ExcelHelper;
using ExcelHelper.Aspose;
using NUnit.Framework;

namespace ExcelHelperTest
{
    public class ExcelTempTest_Aspose : ExcelTempTest
    {
        [SetUp]
        public void Setup()
        {
            _excelHelper = new ExcelTempHelper();
        }

    }
}
