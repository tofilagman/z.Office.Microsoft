using z.Data;

namespace z.Office.Microsoft
{
    public class ExcelWorkSheet : PairCollection
    {
        public readonly string Name;

        public ExcelWorkSheet(string name)
        {
            Name = name;
        }
    }
}
