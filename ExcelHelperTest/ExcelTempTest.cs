using ExcelHelper;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelperTest
{
    public abstract class ExcelTempTest
    {
        protected IExcelTempHelper _excelHelper;

        [Test]
        public void Test_TempSheet()
        {
            var tempIo = new DemoTempIO()
            {
                A = "a",
                B = 1,
                C = DateTime.Now,
                Children = new List<DemoTempChild>()
                {
                    new DemoTempChild()
                    {
                        Name = "a",
                        Age = 1,
                    },
                    new DemoTempChild()
                    {
                        Name = "b",
                        Age = 2,
                    }
                }
            };
            var bytes = _excelHelper.SetData("Excel.xlsx", tempIo);
            File.WriteAllBytes("test.xlsx", bytes);

            var data = _excelHelper.GetData<DemoTempIO>("test.xlsx");
            Assert.That(tempIo.A, Is.EqualTo(data.A));
            Assert.That(tempIo.B, Is.EqualTo(data.B));
            Assert.That(tempIo.C.ToString("yyyy-MM-dd HH:mm:ss"), Is.EqualTo(data.C.ToString("yyyy-MM-dd HH:mm:ss")));
            Assert.That(tempIo.Children.Count, Is.EqualTo(data.Children.Count));
        }

        [Test]
        public void Test_TempSheet_Setting()
        {
            var tempIo = new DemoTempIO()
            {
                A = "A1",
                B = 1,
                C = DateTime.Now,
                D = "D",
                Children = new List<DemoTempChild>()
                {
                    new DemoTempChild()
                    {
                        Name = "A1",
                        Age = 1,
                        Other = "Other"
                    },
                    new DemoTempChild()
                    {
                        Name = "A2",
                        Age = 2,
                    }
                }
            };

            var tempSetting = new TempSetting();
            tempSetting.AddCellAddress(nameof(DemoTempIO.A), "A8");
            tempSetting.AddCellAddress(nameof(DemoTempIO.B), "B8");
            tempSetting.AddCellAddress(nameof(DemoTempIO.C), "C8");
            tempSetting.AddCellAddress(nameof(DemoTempIO.D), "D8");
            tempSetting.AddRequiredProperties(nameof(DemoTempIO.A));
            tempSetting.AddRequiredMessage(nameof(DemoTempIO.A), "A是必须的");
            tempSetting.AddUniqueProperties(nameof(DemoTempIO.A));
            tempSetting.AddUniqueMessage(nameof(DemoTempIO.A), "A必须唯一");
            tempSetting.AddLimitValues(nameof(DemoTempIO.A), "A1", "A2", "A3");
            tempSetting.AddLimitMessage(nameof(DemoTempIO.A), "AA数据非法");
            tempSetting.AddValueTrim(nameof(DemoTempIO.A), Trim.All);

            var childrenSetting = tempSetting.AddTempListSetting(nameof(DemoTempIO.Children), TempListType.Row, 10, 15);
            childrenSetting.AddItemIndex(nameof(DemoTempChild.Name), 0);
            childrenSetting.AddItemIndex(nameof(DemoTempChild.Age), 5);
            childrenSetting.AddItemIndex(nameof(DemoTempChild.Other), 3);
            childrenSetting.AddRequiredProperties(nameof(DemoTempChild.Name));
            //childrenSetting.AddRequiredMessage(nameof(DemoTempChild.Other), "Other是必须的");
            childrenSetting.AddUniqueProperties(nameof(DemoTempChild.Name));
            //childrenSetting.AddUniqueMessage(nameof(DemoTempChild.Name), "A必须唯一");
            childrenSetting.AddLimitValues(nameof(DemoTempChild.Name), "A1", "A2", "A3");
            //childrenSetting.AddLimitMessage(nameof(DemoTempChild.Name), "A数据非法");
            childrenSetting.AddValueTrim(nameof(DemoTempChild.Name), Trim.All);

            var bytes = _excelHelper.SetData("Excel.xlsx", tempIo, tempSetting: tempSetting);
            File.WriteAllBytes("test.xlsx", bytes);

            var data = _excelHelper.GetData<DemoTempIO>("test.xlsx", tempSetting: tempSetting);
            Assert.That(tempIo.A, Is.EqualTo(data.A));
            Assert.That(tempIo.B, Is.EqualTo(data.B));
            Assert.That(tempIo.C.ToString("yyyy-MM-dd HH:mm:ss"), Is.EqualTo(data.C.ToString("yyyy-MM-dd HH:mm:ss")));
            Assert.That(tempIo.Children.Count, Is.EqualTo(data.Children.Count));
        }
    }
}
