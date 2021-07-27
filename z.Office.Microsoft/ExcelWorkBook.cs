using System;
using System.Collections.Generic;
using System.Text;
using z.Data;

namespace z.Office.Microsoft
{
    public class ExcelWorkBook : List<ExcelWorkSheet>
    {
        public readonly string Name;

        public ExcelWorkBook(string name)
        {
            Name = name;
        }
    }
}
