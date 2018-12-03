using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace InSys.Office.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            using (var ms = new InSys.Office.ExcelWriter(true))
            {
                ms.FileName = Path.Combine(Environment.CurrentDirectory, Guid.NewGuid().ToString() + ".xlsx");

                ms.AddSheet("test");

                var row = ms.AddRow("test");

                ms.AddCell(row, 0, "Hello World");
                ms.AddCell(row, 1, int.MaxValue);
                ms.AddCell(row, 2, DateTime.Now);

                ms.Save();

            }


        }
    }
}
