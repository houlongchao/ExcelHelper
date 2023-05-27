using ExcelHelper;
using ExcelHelper.Aspose;
using NUnit.Framework;

namespace ExcelHelperTest
{
    public class ExcelWriterTest_Aspose : ExcelWriterTest
    {
        [SetUp]
        public void Setup()
        {
            _excelHelper = new ExcelWriteHelper();
        }

    }
}
