using ExcelHelper;
using System;
using System.Collections.Generic;

namespace ExcelHelperTest
{
    public class DemoTempIO
    {
        [Temp("A1")]
        public string A { get; set; }

        [Temp("B2")]
        public int B { get; set; }

        [Temp("C3")]
        public DateTime C { get; set; }

        public string D { get; set; }

        [TempList(TempListType.Row, 5, 8)]
        public List<DemoTempChild> Children { get; set; }
    }

    public class DemoTempChild
    {
        [TempListItem(1)]
        public string Name { get; set; }

        [TempListItem(2)]
        public int Age { get; set; }

        public string Other { get; set; }
    }
}
