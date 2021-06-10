using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace z.Office.Microsoft.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            using (var ms = new z.Office.Microsoft.ExcelWriter(false))
            {
                ms.FileName = Path.Combine(Environment.CurrentDirectory, Guid.NewGuid().ToString() + ".xls");

                ms.AddSheet("test");

                var row = ms.AddRow("test");

                ms.AddCell(row, 0, "Hello World");
                ms.AddCell(row, 1, int.MaxValue);
                var cell = ms.AddCell(row, 2, DateTime.Now);
                ms.AddCellComment(cell, "Hello from world");

                ms.Save();

            }


        }
    }
}
