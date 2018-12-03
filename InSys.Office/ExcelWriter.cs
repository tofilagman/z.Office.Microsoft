using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace InSys.Office
{
    public class ExcelWriter : IDisposable
    {

        public List<string> ColumnLetters { get; private set; }

        IWorkbook hssworkbook;
        public string FileName { get; set; }
        public bool IsNewFormat { get; private set; }
        Dictionary<string, ISheet> Sheets;
        private Dictionary<string, int> RowIndex;
        public bool DeleteFileOnDisposed { private get; set; } = false;
        public Dictionary<string, HSSFCellStyle> HFontStyle;
        public Dictionary<string, XSSFCellStyle> XFontStyle;


        public ExcelWriter()
        {
            this.Sheets = new Dictionary<string, ISheet>();
            this.RowIndex = new Dictionary<string, int>();
            HFontStyle = new Dictionary<string, HSSFCellStyle>();
            XFontStyle = new Dictionary<string, XSSFCellStyle>();
            this.InitiateColumnCells();
        }

        public ExcelWriter(bool NewFormat) : this()
        {
            this.IsNewFormat = NewFormat;

            if (IsNewFormat)
            {
                this.hssworkbook = new XSSFWorkbook();
            }
            else
            {
                this.hssworkbook = new HSSFWorkbook();
            }
        }

        public ExcelWriter(string FileName) : this(Excel.IsFileInNewFormat(FileName))
        {
            this.FileName = FileName;
            using (FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                if (IsNewFormat)
                {
                    this.hssworkbook = new XSSFWorkbook(fs);
                }
                else
                {
                    this.hssworkbook = new HSSFWorkbook(fs);
                }
            }
        }

        public ExcelWriter(Stream fs, bool NewFormat) : this(NewFormat)
        {
            if (IsNewFormat)
            {
                this.hssworkbook = new XSSFWorkbook(fs);
            }
            else
            {
                this.hssworkbook = new HSSFWorkbook(fs);
            }
        }

        /// <summary>
        /// Gets the index of the last cell Contained in this row <b>PLUS ONE</b>. The result also
        /// happens to be the 1-based column number of the last cell.  This value can be used as a
        /// standard upper bound when iterating over cells:
        /// <pre>
        /// short minColIx = row.GetFirstCellNum();
        /// short maxColIx = row.GetLastCellNum();
        /// for(short colIx=minColIx; colIx&lt;maxColIx; colIx++) {
        /// Cell cell = row.GetCell(colIx);
        /// if(cell == null) {
        /// continue;
        /// }
        /// //... do something with cell
        /// }
        /// </pre>
        /// </summary>
        /// <returns>
        /// short representing the last logical cell in the row <b>PLUS ONE</b>,
        /// or -1 if the row does not contain any cells.
        /// </returns>
        public int CellCount(IRow row)
        {
            return row.LastCellNum;
        }

        public void ClearBook()
        {
            this.hssworkbook.Clear();
        }

        public void AddSheet(string SheetName)
        {
            Sheets.Add(SheetName, this.hssworkbook.CreateSheet(SheetName));
        }

        public IRow AddRow(string SheetName)
        {
            int idx = 0;
            if (!RowIndex.ContainsKey(SheetName))
            {
                RowIndex.Add(SheetName, idx);
            }
            else
            {
                idx = RowIndex[SheetName] + 1;
                RowIndex[SheetName] = idx;
            }

            IRow row = this.hssworkbook.GetSheet(SheetName).CreateRow(idx);

            return row;
        }

        public void AddCell(IRow row, int index, object value, Type type, string Style = "")
        {

            //row.CreateCell(index).SetCellValue(value);
            CreateCell(row, type, index, value);


            if (Style.Trim() != "")
            {
                if (this.IsNewFormat)
                {
                    if (this.XFontStyle.ContainsKey(Style))
                    {
                        row.Cells[index].CellStyle = this.XFontStyle[Style];
                    }
                }
                else if (this.HFontStyle.ContainsKey(Style))
                {
                    row.Cells[index].CellStyle = this.HFontStyle[Style];
                }
            }
        }

        public void AddCell(IRow row, int index, object value, string Style = "") => AddCell(row, index, value, typeof(string), Style);

        public void AddCellFormula(IRow row, int index, string value, string Style = "")
        {
            ICell cell = row.CreateCell(index, CellType.Formula);
            cell.SetCellFormula(value);

            if (Style.Trim() != "")
            {
                if (this.IsNewFormat)
                {
                    if (this.XFontStyle.ContainsKey(Style))
                    {
                        cell.CellStyle = this.XFontStyle[Style];
                    }
                }
                else if (this.HFontStyle.ContainsKey(Style))
                {
                    cell.CellStyle = this.HFontStyle[Style];
                }
            }
        }

        public void CreateCell(IRow row, Type type, int cellIndex, object value)
        {
            if (value == null)
            {
                row.CreateCell(cellIndex).SetCellValue(string.Empty);
                return;
            }

            if (type == typeof(Int32) ||
              type == typeof(Int16) ||
              type == typeof(Int64) ||
              type == typeof(string) ||
              value == DBNull.Value)
            {
                row.CreateCell(cellIndex).SetCellValue(value.ToString());
                return;
            }

            if (type == typeof(bool))
            {
                row.CreateCell(cellIndex).SetCellValue(Convert.ToBoolean(value));
            }

            if (type == typeof(DateTime))
            {
                row.CreateCell(cellIndex).SetCellValue(Convert.ToDateTime(value));
            }

            if (type == typeof(double) ||
                type == typeof(decimal))
            {
                row.CreateCell(cellIndex).SetCellValue(Convert.ToDouble(value));
            }
        }

        public void AddCellStyle(String StyleName, String FontName = "Arial", Int16 FontSize = 8,
           Boolean IsItalic = false, FontUnderlineType UnderlineType = FontUnderlineType.None,
           FontBoldWeight BoldWeight = FontBoldWeight.None, HorizontalAlignment HorizontalAlign = HorizontalAlignment.Left,
           VerticalAlignment VerticalAlign = VerticalAlignment.Top, BorderStyle TopBorder = BorderStyle.None,
           BorderStyle BottomBorder = BorderStyle.None, BorderStyle RightBorder = BorderStyle.None,
           BorderStyle LeftBorder = BorderStyle.None, IndexedColors FontColor = null,
           IndexedColors BackgroundColor = null)
        {
            IFont font = this.hssworkbook.CreateFont();
            font.Color = ((FontColor == null) ? IndexedColors.Black.Index : FontColor.Index);
            font.FontName = FontName;
            font.FontHeightInPoints = FontSize;
            font.IsItalic = IsItalic;
            if (font.Underline != FontUnderlineType.None)
            {
                font.Underline = UnderlineType;
            }
            font.Boldweight = (short)BoldWeight;

            if (this.IsNewFormat)
            {
                XSSFCellStyle style = (XSSFCellStyle)this.hssworkbook.CreateCellStyle();
                style.SetFont(font);
                style.Alignment = HorizontalAlign;
                style.VerticalAlignment = VerticalAlign;
                style.BorderTop = TopBorder;
                style.BorderBottom = BottomBorder;
                style.BorderRight = RightBorder;
                style.BorderLeft = LeftBorder;

                if (BackgroundColor != null)
                {
                    style.FillForegroundColor = BackgroundColor.Index;
                    style.FillPattern = FillPattern.SolidForeground;
                }

                if (!this.XFontStyle.ContainsKey(StyleName))
                {
                    this.XFontStyle.Add(StyleName, style);
                }
                else
                {
                    this.XFontStyle[StyleName] = style;
                }
            }
            else
            {
                HSSFCellStyle style2 = (HSSFCellStyle)this.hssworkbook.CreateCellStyle();
                style2.SetFont(font);
                style2.Alignment = HorizontalAlign;
                style2.VerticalAlignment = VerticalAlign;
                style2.BorderTop = TopBorder;
                style2.BorderBottom = BottomBorder;
                style2.BorderRight = RightBorder;
                style2.BorderLeft = LeftBorder;

                if (BackgroundColor != null)
                {
                    style2.FillForegroundColor = BackgroundColor.Index;
                    style2.FillPattern = FillPattern.SolidForeground;
                }
                if (!this.HFontStyle.ContainsKey(StyleName))
                {
                    this.HFontStyle.Add(StyleName, style2);
                }
                else
                {
                    this.HFontStyle[StyleName] = style2;
                }
            }
        }

        public void MergeCells(string SheetName, int firstRow, int lastRow, int firstCell, int lastCell)
        {
            var cra = new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, firstCell, lastCell);
            this.hssworkbook.GetSheet(SheetName).AddMergedRegion(cra);
        }

        /// <summary>
        /// Only supported on new Version of Excel
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="Password"></param>
        public void ProtectSheet(string SheetName, string Password)
        {
            if (!IsNewFormat)
                throw new InvalidOperationException("Password Protection doesn't supported in old version of Excel");

            var sheet = (XSSFSheet)hssworkbook.GetSheet(SheetName);
            sheet.LockDeleteColumns();
            sheet.LockDeleteRows();
            sheet.LockInsertColumns();
            sheet.LockInsertRows();
            sheet.ProtectSheet(Password);
            sheet.EnableLocking();
            ((XSSFWorkbook)hssworkbook).LockStructure();
        }

        public void Save()
        {
            foreach (KeyValuePair<string, ISheet> pair in this.Sheets)
            {
                int num = pair.Value.LastRowNum + 1;
                int num2 = 0;
                for (int i = 0; i < num; i++)
                {
                    int physicalNumberOfCells = pair.Value.GetRow(i).PhysicalNumberOfCells;
                    if (physicalNumberOfCells > num2)
                    {
                        num2 = physicalNumberOfCells;
                    }
                }
                for (int j = 0; j < num2; j++)
                {
                    pair.Value.AutoSizeColumn(j);
                }
            }

            using (MemoryStream ms = new MemoryStream())
            {

                if (this.IsNewFormat)
                {
                    ((XSSFWorkbook)this.hssworkbook).Write(ms);
                }
                else
                {
                    this.hssworkbook.Write(ms);
                }

                using (FileStream file = new FileStream(this.FileName, FileMode.Create, FileAccess.Write))
                {
                    ms.WriteTo(file);
                }
            }
        }

        public void SaveToStream(Stream ms)
        {
            foreach (KeyValuePair<string, ISheet> pair in this.Sheets)
            {
                int num = pair.Value.LastRowNum + 1;
                int num2 = 0;
                for (int i = 0; i < num; i++)
                {
                    int physicalNumberOfCells = pair.Value.GetRow(i).PhysicalNumberOfCells;
                    if (physicalNumberOfCells > num2)
                    {
                        num2 = physicalNumberOfCells;
                    }
                }
                for (int j = 0; j < num2; j++)
                {
                    pair.Value.AutoSizeColumn(j);
                }
            }

            //using (MemoryStream ms = new MemoryStream())
            //{

            if (this.IsNewFormat)
            {
                ((XSSFWorkbook)this.hssworkbook).Write(ms);
            }
            else
            {
                this.hssworkbook.Write(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);


            //}
        }

        protected void InitiateColumnCells()
        {
            var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            ColumnLetters = new List<string>();

            for (var i = 0; i < alphabet.Length; i++)
                ColumnLetters.Add(alphabet[i].ToString());

            for (var i = 0; i < alphabet.Length; i++)
                for (var o = 0; o < alphabet.Length; o++)
                    ColumnLetters.Add($"{ alphabet[i] }{ alphabet[o] }");
        }

        public string GetCellColumnLetter(int index)
        {
            if (index > ColumnLetters.Count)
                throw new InvalidOperationException("The index reached the maximum cell available");
            return ColumnLetters[index];
        }

        public void FreezePane(string SheetName, int colSplit, int rowSplit)
        {
            this.hssworkbook.GetSheet(SheetName).CreateFreezePane(colSplit, rowSplit);
        }

        public void FreezePane(string SheetName, int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            this.hssworkbook.GetSheet(SheetName).CreateFreezePane(colSplit, rowSplit, leftMostColumn, topRow);
        }

        public void SplitPane(string SheetName, int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            this.hssworkbook.GetSheet(SheetName).CreateSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane);
        }

        public void Dispose()
        {
            if (this.hssworkbook != null) this.hssworkbook.Dispose();
            if (DeleteFileOnDisposed && File.Exists(this.FileName)) File.Delete(this.FileName);
            GC.Collect();
            GC.SuppressFinalize(this);
        }
    }
}
