using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace InSys.Office
{
    /// <summary>
    /// LJ 20131209, NPOI
    /// Excel new Library, using NPOI
    /// Supported Microsoft Office Application without interop, current 2003 and latest formats
    /// </summary>
    public class Excel : IDisposable
    {
        #region Variables

        //HSSFWorkbook hssworkbook;
        protected IWorkbook hssworkbook;
        public string mFilename { get; set; }
        public bool IsNewFormat;

        private Dictionary<string, ICellStyle> FontStyle;

        #endregion

        #region Constructor

        public Excel()
        {
            this.FontStyle = new Dictionary<string, ICellStyle>();
        }
        /// <summary>
        /// Read/Write a Excel File
        /// </summary>
        /// <param name="Filename"></param>
        /// <param name="NewFile"></param>
        public Excel(string Filename, bool NewFile = true): this()
        {
            this.mFilename = Filename;
            this.IsNewFormat = Excel.IsFileInNewFormat(Filename);

            if (NewFile)
            {
                if (IsNewFormat)
                {
                    this.hssworkbook = new XSSFWorkbook();
                }
                else
                {
                    this.hssworkbook = new HSSFWorkbook();
                }
            }
            else
            {
                using (FileStream fs = new FileStream(mFilename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
        }

        public Excel(Stream fs, bool NewFormat): this()
        {
            this.IsNewFormat = NewFormat;
            if (IsNewFormat)
            {
                fs.Position = 0;
                this.hssworkbook = new XSSFWorkbook(fs);
            }
            else
            {
                this.hssworkbook = new HSSFWorkbook(fs);
            }

        }

        public void Dispose()
        {
            if (this.hssworkbook != null) this.hssworkbook.Dispose();
            GC.Collect();
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Read

        public DataTable Read(string sheet)
        {
            try
            {
                if (this.hssworkbook == null) throw new Exception("Do you intend to read a file? please specify file in constructor");

                ISheet sht = this.hssworkbook.GetSheet(sheet);

                if (sht == null) { throw new Exception("Sheet not found: " + sheet); }

                System.Collections.IEnumerator rows = sht.GetRowEnumerator();

                DataTable dt = new DataTable(sheet);
                //get header

                rows.MoveNext();

                IRow row = this.GetRow(rows.Current);

                for (int j = 0; j < row.LastCellNum; j++)
                {
                    if (row.GetCell(j) == null) throw new Exception("Column Header is undefined");
                    dt.Columns.Add(row.GetCell(j).ToString());
                }

                while (rows.MoveNext())
                {
                    row = this.GetRow(rows.Current); //(HSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();

                    for (int i = 0; i < dt.Columns.Count; i++) //row.LastCellNum
                    {
                        ICell cell = row.GetCell(i);

                        if (cell == null)
                        {
                            dr[i] = DBNull.Value;
                        }
                        else
                        {
                            dr[i] = handleCell(cell.CellType, cell);
                        }
                    }
                    dt.Rows.Add(dr);
                }

                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Excel Reading Error: " + ex.Message);
            }
        }

        private object handleCell(CellType type, ICell cell)
        {
            switch (type)
            {
                case CellType.String: return cell.StringCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    return cell.NumericCellValue;
                case CellType.Boolean: return cell.BooleanCellValue;
                case CellType.Formula: return handleCell(cell.CachedFormulaResultType, cell);
                case CellType.Blank: return DBNull.Value;
                default: return cell.ToString();
            }
        }

        public List<string> SheetsNames
        {
            get
            {
                List<string> str = new List<string>();

                int sht = this.hssworkbook.NumberOfSheets;

                for (int i = 0; i < sht; i++)
                {
                    str.Add(this.hssworkbook.GetSheetAt(i).SheetName);
                }

                return str;
            }
        }

        public DataSet Read()
        {
            try
            {
                int sht = this.hssworkbook.NumberOfSheets;
                DataSet ds = new DataSet("WorkBook");
                for (int i = 0; i < sht; i++)
                {
                    ds.Tables.Add(this.Read(this.hssworkbook.GetSheetAt(i).SheetName));
                }

                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected IRow GetRow(object Current)
        {
            if (this.IsNewFormat)
            {
                return (XSSFRow)Current;
            }
            else
            {
                return (HSSFRow)Current;
            }
        }

        #endregion

        #region Write

        public void Write(DataTable dt, string sheet = "Sheet1")
        {
            try
            {
                ISheet sht = this.hssworkbook.CreateSheet(sheet);

                //Create header
                IRow hrow = sht.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    hrow.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row = sht.CreateRow(i + 1);
                    for (int e = 0; e < dt.Columns.Count; e++)
                    {
                        if (dt.Rows[i][e].GetType() == typeof(bool))
                        {
                            row.CreateCell(e).SetCellValue(Convert.ToBoolean(dt.Rows[i][e]));
                        }

                        if (dt.Rows[i][e].GetType() == typeof(DateTime))
                        {
                            row.CreateCell(e).SetCellValue(Convert.ToDateTime(dt.Rows[i][e]));
                        }

                        if (dt.Rows[i][e].GetType() == typeof(double) ||
                            dt.Rows[i][e].GetType() == typeof(decimal))
                        {
                            row.CreateCell(e).SetCellValue(Convert.ToDouble(dt.Rows[i][e]));
                        }

                        if (dt.Rows[i][e].GetType() == typeof(Int32) ||
                            dt.Rows[i][e].GetType() == typeof(Int16) ||
                            dt.Rows[i][e].GetType() == typeof(Int64) ||
                            dt.Rows[i][e].GetType() == typeof(string) ||
                            dt.Rows[i][e] == DBNull.Value)
                        {
                            row.CreateCell(e).SetCellValue(dt.Rows[i][e].ToString());
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Write(DataSet ds)
        {
            try
            {
                foreach (DataTable dt in ds.Tables)
                {
                    this.Write(dt, dt.TableName);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region Workout

        public void AddData(DataTable table, string SheetName, int StartingCell = 0, int StartingRow = 0)
        {
            try
            {
                var shet = GetSheet(SheetName);

                for (var i = 0; i < table.Rows.Count; i++)
                {
                    IRow row = shet.CreateRow(i + StartingRow);
                    for (int e = 0; e < table.Columns.Count; e++)
                    {
                        CreateCell(row, table.Rows[i][e].GetType(), e + StartingCell, table.Rows[i][e]);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetCell(IRow row, Type type, int cellIndex, object value)
        {
            try
            {
                if (type == typeof(bool))
                {
                    row.GetCell(cellIndex).SetCellValue(Convert.ToBoolean(value));
                }

                if (type == typeof(DateTime))
                {
                    row.GetCell(cellIndex).SetCellValue(Convert.ToDateTime(value));
                }

                if (type == typeof(double) ||
                    type == typeof(decimal))
                {
                    row.GetCell(cellIndex).SetCellValue(Convert.ToDouble(value));
                }

                if (type == typeof(Int32) ||
                    type == typeof(Int16) ||
                    type == typeof(Int64) ||
                    type == typeof(string) ||
                    value == DBNull.Value)
                {
                    row.GetCell(cellIndex).SetCellValue(value.ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void AddHeader(DataColumnCollection dc, string SheetName, int StartingCell = 0, int StartingRow = 0)
        {
            try
            {
                var shet = GetSheet(SheetName);

                IRow row = shet.CreateRow(StartingRow);
                for (int e = 0; e < dc.Count; e++)
                {
                    CreateCell(row, typeof(string), e + StartingCell, dc[e].ColumnName);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CreateCell(IRow row, Type type, int cellIndex, object value)
        {
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

            if (type == typeof(Int32) ||
                type == typeof(Int16) ||
                type == typeof(Int64) ||
                type == typeof(string) ||
                value == DBNull.Value)
            {
                row.CreateCell(cellIndex).SetCellValue(value.ToString());
            }
        }

        public ISheet GetSheet(string SheetName)
        {
            if (this.SheetsNames.Contains(SheetName))
                return this.hssworkbook.GetSheet(SheetName);
            else
                return this.hssworkbook.CreateSheet(SheetName);
        }

        public void AddCellStyle(String StyleName, String FontName = "Arial", Int16 FontSize = 8,
          Boolean IsItalic = false, FontUnderlineType UnderlineType = FontUnderlineType.None,
          FontBoldWeight BoldWeight = FontBoldWeight.None, HorizontalAlignment HorizontalAlign = HorizontalAlignment.Left,
          VerticalAlignment VerticalAlign = VerticalAlignment.Top, BorderStyle TopBorder = BorderStyle.None,
          BorderStyle BottomBorder = BorderStyle.None, BorderStyle RightBorder = BorderStyle.None,
          BorderStyle LeftBorder = BorderStyle.None, IndexedColors FontColor = null,
           IndexedColors BackgroundColor = null, FillPattern fp = FillPattern.SolidForeground, short DataFormat = 0)
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
                style.DataFormat = DataFormat;

                if (BackgroundColor != null)
                {
                    style.FillForegroundColor = BackgroundColor.Index;
                    style.FillPattern = fp;
                }

                if (!this.FontStyle.ContainsKey(StyleName))
                {
                    this.FontStyle.Add(StyleName, style);
                }
                else
                {
                    this.FontStyle[StyleName] = style;
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
                style2.DataFormat = DataFormat;

                if (BackgroundColor != null)
                {
                    style2.FillForegroundColor = BackgroundColor.Index;
                    style2.FillPattern = fp;
                }

                if (!this.FontStyle.ContainsKey(StyleName))
                {
                    this.FontStyle.Add(StyleName, style2);
                }
                else
                {
                    this.FontStyle[StyleName] = style2;
                }
            }
        }

        public void SetCellStyle(IRow row, int index, string Style)
        {
            if (Style.Trim() != "")
            {
                row.Cells[index].CellStyle = FontStyle[Style];
            }
        }

        public IRow GetRow(string SheetName, int RowIndex)
        {
            try
            {
                var shte = GetSheet(SheetName);

                return shte.GetRow(RowIndex);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IRow CreateRow(string SheetName, int RowIndex)
        {
            try
            {
                var shte = GetSheet(SheetName);

                return shte.CreateRow(RowIndex);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void DeleteCell(IRow row, int v)
        {
            row.RemoveCell(row.GetCell(v));
        }

        public string GetExtension
        {
            get
            {
                return IsNewFormat ? ".xlsx" : ".xls";
            }
        }

        #endregion

        #region Save

        public virtual void Save()
        {
            if (this.mFilename == null || this.mFilename == "")
            {
                throw new Exception("Please Specify FileName Property");
            }

            //foreach (string pair in SheetsNames)
            //{
            //    var shet = hssworkbook.GetSheet(pair);
            //    int num = shet.LastRowNum + 1;
            //    int num2 = 0;
            //    for (int i = 0; i < num; i++)
            //    {
            //        int physicalNumberOfCells = shet.GetRow(i).PhysicalNumberOfCells;
            //        if (physicalNumberOfCells > num2)
            //        {
            //            num2 = physicalNumberOfCells;
            //        }
            //    }
            //    for (int j = 0; j < num2; j++)
            //    {
            //        shet.AutoSizeColumn(j);
            //    }
            //}

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

                using (FileStream file = new FileStream(this.mFilename, FileMode.Create, FileAccess.Write))
                {
                    ms.WriteTo(file);
                }
            }
        }

        //public void Save(string username, string Password)
        //{
        //    if (this.mFilename == null || this.mFilename == "")
        //    {
        //        throw new Exception("Please Specify FileName Property");
        //    }

        //    using (MemoryStream ms = new MemoryStream())
        //    {
        //        //this.hssworkbook.WriteProtectWorkbook(Password, username);
        //        this.hssworkbook.Write(ms);
        //        using (FileStream file = new FileStream(this.mFilename, FileMode.Create, FileAccess.Write))
        //        {
        //            ms.WriteTo(file);
        //        }
        //    }
        //}

        #endregion

        #region Static

        public static bool IsFileInNewFormat(string ExcelFile)
        {
            return (Path.GetExtension(ExcelFile).ToLower() == ".xls") ? false : true;
        }

        #endregion
    }
}
