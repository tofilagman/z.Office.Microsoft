using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using z.Data;

namespace z.Office.Microsoft
{
    /// <summary>
    /// LJ 20131209, NPOI
    /// Excel new Library, using NPOI
    /// Supported Microsoft Office Application without interop, current 2003 and latest formats
    /// </summary>
    public class ExcelReader : IDisposable
    {
        #region Variables

        //HSSFWorkbook hssworkbook;
        protected IWorkbook hssworkbook;
        public string mFilename { get; set; }
        public bool IsNewFormat;

        #endregion

        #region Constructor

        /// <summary>
        /// Read/Write a Excel File
        /// </summary>
        /// <param name="Filename"></param>
        /// <param name="NewFile"></param>
        public ExcelReader(string Filename)
        {
            this.mFilename = Filename;
            this.IsNewFormat = !(Path.GetExtension(Filename).ToLower() == ".xls");

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

        public ExcelReader(Stream fs, bool NewFormat)
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

        ~ExcelReader() => Dispose();

        public void Dispose()
        {
            if (this.hssworkbook != null) this.hssworkbook.Dispose();
            GC.Collect();
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Read

        /// <summary>
        /// Reads Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">defaults to 1 to skip the header</param>
        /// <returns></returns>
        public ExcelWorkSheet ReadSheet(string sheet, int rowIndex = 1)
        {
            try
            {
                if (this.hssworkbook == null) throw new Exception("Do you intend to read a file? please specify file in constructor");

                ISheet sht = this.hssworkbook.GetSheet(sheet);

                if (sht == null) { throw new Exception("Sheet not found: " + sheet); }

                System.Collections.IEnumerator rows = sht.GetRowEnumerator();

                for (var i = 0; i < rowIndex; i++)
                    rows.MoveNext();

                IRow row = this.GetRow(rows.Current);

                var cols = new Pair<int, string>();
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    if (row.GetCell(j) == null) throw new Exception("Column Header is undefined");
                    cols.Add(j, row.GetCell(j).ToString());
                }
                var pps = new ExcelWorkSheet(sht.SheetName);
                while (rows.MoveNext())
                {
                    row = this.GetRow(rows.Current);
                    var pp = new Pair();

                    for (int i = 0; i < cols.Keys.Count; i++)
                    {
                        ICell cell = row.GetCell(i);

                        if (cell == null)
                            pp.Add(cols[i], null);
                        else
                            pp.Add(cols[i], HandleCell(cell.CellType, cell));
                    }
                    pps.Add(pp);
                }

                return pps;
            }
            catch (Exception ex)
            {
                throw new Exception("Excel Reading Error: " + ex.Message);
            }
        }

        private object HandleCell(CellType type, ICell cell)
        {
            switch (type)
            {
                case CellType.String: return cell.StringCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    return cell.NumericCellValue;
                case CellType.Boolean: return cell.BooleanCellValue;
                case CellType.Formula: return HandleCell(cell.CachedFormulaResultType, cell);
                case CellType.Blank: return null;
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

        public ExcelWorkBook ReadBook()
        {
            try
            {
                int sht = this.hssworkbook.NumberOfSheets;
                var ds = new ExcelWorkBook("WorkBook");
                for (int i = 0; i < sht; i++)
                    ds.Add(this.ReadSheet(this.hssworkbook.GetSheetAt(i).SheetName));

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
    }
}
