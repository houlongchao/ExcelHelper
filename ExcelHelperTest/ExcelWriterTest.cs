using ExcelHelper;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelperTest
{
    public abstract class ExcelWriterTest
    {
        protected IExcelWriteHelper _excelHelper;

        [Test]
        public void Test_ExportSheet()
        {
            var datas = new List<DemoIO>();
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2",
                C = "a"
            });
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2",
                C = "b"
            });
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2",
                C = "C3"
            }); 
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                C = "c",
                Status = Status.B,
            });
            _excelHelper.ExportSheet("test", datas);
            var bytes = _excelHelper.ToBytes();
            File.WriteAllBytes("test.xlsx", bytes);
        }
    }
}
