using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.CustomProperties;

using commanet.Db;

namespace commanet.Excel
{
    public class XLWorkbook : IDisposable
    {
        #region Public Properties
        public List<XlWorksheet> Worksheets { get; } = new List<XlWorksheet>();

        public SQLDBConnection? Db { get; set; } = null;

        public string this[string RefAddress]
        {
            get
            {
                return GetCellValue<string>(RefAddress);
            }
            set
            {
                SetCellValue(RefAddress, value);
            }
        }

        public Dictionary<string, object> CustomProperties
        {
            get
            {
                var res = new Dictionary<string, object>();
                var props = Document.CustomFilePropertiesPart.Properties;
                foreach (var p in props.Elements<CustomDocumentProperty>())
                {
                    foreach (var pr in p.GetType().GetProperties())
                    {
                        if (pr.Name.ToString(CultureInfo.InvariantCulture)
                              .StartsWith("VT", StringComparison.Ordinal))
                        {
                            var v = (OpenXmlElement?)pr.GetValue(p);
                            if (v != null)
                            {
                                res.Add(p.Name, v.InnerText);
                                break;
                            }
                        }
                    }
                }
                return res;
            }
        }

        #endregion

        #region Public Methods

        public XLWorkbook(SQLDBConnection? db = null)
        {
            ms = new MemoryStream();
            Document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            Db = db;
        }

        ~XLWorkbook()
        {
            Dispose(false);
        }


        public static XLWorkbook Open(string FileName, bool Editable = true, SQLDBConnection? db = null)
        {
            var wb = new XLWorkbook()
            {
                Document = SpreadsheetDocument.Open(FileName, Editable),
                Db = db
            };
            wb.FillWorksheets();
            return wb;
        }

        public static XLWorkbook Open(Stream stream, bool Editable = true, SQLDBConnection? db = null)
        {
            var wb = new XLWorkbook()
            {
                Document = SpreadsheetDocument.Open(stream, Editable),
                Db = db
            };
            wb.FillWorksheets();
            return wb;
        }

        public void Close()
        {
            if (Document != null)
            {
                Document.Close();
                Document.Dispose();
                ms?.Dispose();
            }

        }

        #region Dispose implementation
        private bool isDisposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (isDisposed) return;

            if (disposing)
            {
                Document?.Close();
            }

            isDisposed = true;
        }
        #endregion

        public void Save()
        {
            Document.Save();
        }

        public void SaveAs(string FilePath)
        {
            Document.SaveAs(FilePath);
        }

        public SpreadsheetDocument Document { get; private set; }

        public void SetCellValue<T>(string RefAddr, T value)
        {
            var raddr = new XLRefAddress(this, RefAddr);
            SetCellValue(raddr.SheetName, raddr.ColumnIndex1, raddr.RowIndex1, value);
        }

        public void SetCellValue<T>(string SheetName, uint ColumnIdx, uint RowIdx, T value)
        {
            var cell = GetCell(SheetName, ColumnIdx, RowIdx);
            if (cell != null)
            {
                if ((cell.DataType != null && cell.DataType == CellValues.Date) || typeof(T) == typeof(DateTime))
                {
                    var v = Convert.ChangeType(value, typeof(DateTime), CultureInfo.InvariantCulture);
                    if (v != null)
                    {
                        var dt = (DateTime)v;
                        cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                    }
                }
                else if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var v = Convert.ChangeType(value, typeof(string), CultureInfo.InvariantCulture);
                    if (v != null)
                    {
                        var dt = (string)v;
                        var idx = int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);
                        var tbl = Document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                        var item = (SharedStringItem)tbl.ChildElements.Skip(idx).First();
                        item.Text = new Text(dt);
                    }
                }
                else if (typeof(T) == typeof(bool))
                {
                    var v = Convert.ChangeType(value, typeof(bool), CultureInfo.InvariantCulture);
                    if (v != null)
                    {
                        var dt = (bool)v;
                        cell.DataType = CellValues.Boolean;
                        #pragma warning disable CA1303 // Do not pass literals as localized parameters
                        cell.CellValue = dt ? new CellValue(EX_BOOL_TRUE) : new CellValue(EX_BOOL_FALSE);
                        #pragma warning restore CA1303 // Do not pass literals as localized parameters
                    }
                }
                else if (typeof(T) == typeof(double) || typeof(T) == typeof(int))
                {
                    var v = Convert.ChangeType(value, typeof(double), CultureInfo.InvariantCulture);
                    if (v != null)
                    {
                        var dt = (double)v;
                        cell.CellValue = new CellValue(dt.ToString(CultureInfo.InvariantCulture));
                    }
                }
                else
                {
                    if (typeof(T) == typeof(DateTime))
                    {
                        var v = Convert.ChangeType(value, typeof(DateTime), CultureInfo.InvariantCulture);
                        if (v != null)
                        {
                            var dt = (DateTime)v;
                            cell.CellValue = new CellValue(dt);
                        }
                    }
                    else
                    {
                        var v = Convert.ChangeType(value, typeof(string), CultureInfo.InvariantCulture);
                        if (v != null)
                        {
                            var dt = (string)v;
                            cell.CellValue = new CellValue(dt);
                            cell.DataType = CellValues.String;
                        }
                    }


                }
            }
        }

        public T GetCellValue<T>(string RefAddress)
        {
            T res = default;
            var cell = GetCell(RefAddress);
            if (cell != null)
            {
                var cellValue = cell.CellValue != null ? cell.CellValue.Text : "";
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var idx = int.Parse(cellValue, CultureInfo.InvariantCulture);
                    var tbl = Document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    var item = (SharedStringItem)tbl.ChildElements.Skip(idx).First();
                    res = (T)Convert.ChangeType(item.Text.InnerText, typeof(T), CultureInfo.InvariantCulture);
                }
                else if ((cell.DataType != null && cell.DataType == CellValues.Date) || typeof(T) == typeof(DateTime))
                {
                    var d = (double)Convert.ChangeType(cellValue, typeof(double), CultureInfo.InvariantCulture);
                    res = (T)Convert.ChangeType(DateTime.FromOADate(d), typeof(T), CultureInfo.InvariantCulture);
                }
                else if ((cell.DataType != null && cell.DataType == CellValues.Boolean) || typeof(T) == typeof(bool))
                {
                    var d = (cellValue == "1");
                    res = (T)Convert.ChangeType(d, typeof(T), CultureInfo.InvariantCulture);
                }
                else
                {

                    res = (T)Convert.ChangeType(cellValue, typeof(T), CultureInfo.InvariantCulture);
                }
            }
            if (res == null)
                throw new Exception($"Cell value not found in address {RefAddress}");
            return res;
        }

        public T GetCellValue<T>(string SheetName, uint ColumnIdx, uint RowIdx)
        {
            T res = default;
            var cell = GetCell(SheetName, ColumnIdx, RowIdx);
            if (cell != null)
            {
                var cellValue = cell.CellValue != null ? cell.CellValue.Text : "";
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var idx = int.Parse(cellValue, CultureInfo.InvariantCulture);
                    var tbl = Document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    var item = (SharedStringItem)tbl.ChildElements.Skip(idx).First();
                    res = (T)Convert.ChangeType(item.Text.InnerText, typeof(T), CultureInfo.InvariantCulture);
                }
                else if ((cell.DataType != null && cell.DataType == CellValues.Date) || typeof(T) == typeof(DateTime))
                {
                    var d = (double)Convert.ChangeType(cellValue, typeof(double), CultureInfo.InvariantCulture);
                    res = (T)Convert.ChangeType(DateTime.FromOADate(d), typeof(T), CultureInfo.InvariantCulture);
                }
                else if ((cell.DataType != null && cell.DataType == CellValues.Boolean) || typeof(T) == typeof(bool))
                {
                    var d = (cellValue == "1");
                    res = (T)Convert.ChangeType(d, typeof(T), CultureInfo.InvariantCulture);
                }
                else
                {

                    res = (T)Convert.ChangeType(cellValue, typeof(T), CultureInfo.InvariantCulture);
                }
            }
            if (res == null)
                throw new Exception($"Cell value not found in {SheetName}!{XLRefAddress.GetColumnName(ColumnIdx)}{RowIdx}");
            return res;
        }

        public XlWorksheet? FirstWorkSheet
        {
            get => Worksheets.Count > 0 ? Worksheets[0]
                                        : null;
        }

        public XlWorksheet? GetWorkSheet(string WorksheetName)
        {
            return Worksheets.Find(ws => ws.WorksheetName == WorksheetName);
        }

        public WorksheetPart? GetWorksheetPartByName(string sheetName)
        {
            IEnumerable<Sheet> sheets =
               Document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (!sheets.Any())
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 Document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        public string? GetNameDefRef(string Name)
        {
            string? res = null;
            var defNames = Document.WorkbookPart.Workbook.DefinedNames;
            foreach (DefinedName dn in defNames)
            {
                if (dn.Name.Value == Name)
                {
                    res = dn.Text;
                    break;
                }
            }
            return res;
        }

        public void FillArea(string RefAddress, object[][] data, bool Extend = true, bool Transposed = false)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            var range = new XLRefAddress(this, RefAddress);
            var ws = GetWorkSheet(range.SheetName);
            if (ws == null) return;
            var row = range.RowIndex1;
            var col = range.ColumnIndex1;

            for (int r = 0; r < data.Length; r++)
            {
                for (int c = 0; c < data[r].Length; c++)
                {
                    SetCellValue(range.SheetName, col, row, data[r][c]);
                    if (Transposed)
                    {
                        row = ws.GetNextRowIndex(row, col);
                        if (row > range.RowIndex2) break;
                    }
                    else
                    {
                        col = ws.GetNextColumnIndex(row, col);
                        if (col > range.ColumnIndex2) break;
                    }
                }
                if (Extend && r < data.Length - 1)
                {
                    if (Transposed)
                        ws.InsertCellsRight(col, range.RowIndex1, range.RowIndex2);
                    else
                        ws.InsertCellsBelow(row, range.ColumnIndex1, range.ColumnIndex2);
                }

                if (Transposed)
                {
                    row = range.RowIndex1;
                    col = ws.GetNextColumnIndex(row, col);
                    if (!Extend && col > range.ColumnIndex2) break;
                }
                else
                {
                    col = range.ColumnIndex1;
                    row = ws.GetNextRowIndex(row, col);
                    if (!Extend && row > range.RowIndex2) break;
                }
            }
        }

        public void FillArea(string RefAddress, string SQL, bool Extend = true, bool Transposed = false, params KeyValuePair<string, object>[] SqlParams)
        {
            CheckDb();
            if (Db == null) return; //Just to suppress static code analyzer warning
                                    //Actually this check is already performed in CheckDb() 
            var range = new XLRefAddress(this, RefAddress);
            var ws = GetWorkSheet(range.SheetName);
            if (ws == null) return;
            var row = range.RowIndex1;
            var col = range.ColumnIndex1;
            var prevRow = row;
            var prevCol = col;

            int r = 0;

            Db.ExecuteReader(SQL, rd =>
            {
                r++;
                // Do not go out of fixed area boundary
                if (!Extend && !Transposed && row > range.RowIndex2) return false;
                if (!Extend && Transposed && col > range.ColumnIndex2) return false;

                if (Extend && r > 1)
                {
                    if (Transposed)
                        ws.InsertCellsRight(prevCol, range.RowIndex1, range.RowIndex2);
                    else
                        ws.InsertCellsBelow(prevRow, range.ColumnIndex1, range.ColumnIndex2);
                }

                for (int c = 0; c < rd.FieldCount; c++)
                {
                    var v = (string)Convert.ChangeType(rd.GetValue(c), typeof(string), CultureInfo.InvariantCulture);
                    SetCellValue(range.SheetName, col, row, v);
                    if (Transposed)
                    {
                        row = ws.GetNextRowIndex(row, col);
                        if (row > range.RowIndex2) break;
                    }
                    else
                    {
                        col = ws.GetNextColumnIndex(row, col);
                        if (col > range.ColumnIndex2) break;
                    }
                }

                prevRow = row;
                prevCol = col;

                if (Transposed)
                {
                    row = range.RowIndex1;
                    col = ws.GetNextColumnIndex(row, col);
                }
                else
                {
                    col = range.ColumnIndex1;
                    row = ws.GetNextRowIndex(row, col);
                }

                return true;
            }, SqlParams);
        }

        public void FillCells<T>(T dataobj)
            where T : class
        {
            if (dataobj == null)
                throw new ArgumentNullException(nameof(dataobj));

            foreach (var p in dataobj.GetType().GetProperties())
            {
                var raddr = GetNameDefRef(p.Name);
                var v = p.GetValue(dataobj);
                if (raddr != null && v != null)
                {
                    var sv = (string)Convert.ChangeType(v, typeof(string), CultureInfo.InvariantCulture);
                    SetCellValue(raddr, sv);
                }
            }
        }

        public void FillCells(Dictionary<string, object> data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            foreach (var it in data)
            {
                var raddr = GetNameDefRef(it.Key);
                var v = it.Value;
                if (raddr != null && v != null)
                {
                    var sv = (string)Convert.ChangeType(v, typeof(string), CultureInfo.InvariantCulture);
                    SetCellValue(raddr, sv);
                }
            }
        }

        public void FillCells(string SQL, params KeyValuePair<string, object>[] SqlParams)
        {
            CheckDb();
            if (Db == null) return; //Just to suppress static code analyzer warning
                                    //Actually this check is already performed in CheckDb() 
            Db.ExecuteReader(SQL, rd =>
            {
                for (int i = 0; i < rd.FieldCount; i++)
                {
                    var sv = (string)Convert.ChangeType(rd.GetValue(i), typeof(string), CultureInfo.InvariantCulture);
                    var name = rd.GetName(i);
                    SetCellValue<string>(name, sv);
                }
                return false;
            }, SqlParams);
        }

        public void SetActiveSheet(string SheetName)
        {
            var wsIdx = Worksheets.FindIndex(ws => ws.WorksheetName == SheetName);
            if (wsIdx < 0) return;
            var wbv = Document.WorkbookPart
                             .Workbook
                             .Descendants<WorkbookView>().First();
            wbv.ActiveTab = (uint)wsIdx;
        }

        /// <summary>
        /// Set active sheet
        /// </summary>
        /// <param name="SheetIdx">Zero based sheet index</param>
        public void SetActiveSheet(uint SheetIdx)
        {
            var wbv = Document.WorkbookPart
                             .Workbook
                             .Descendants<WorkbookView>().First();
            wbv.ActiveTab = SheetIdx;
        }

        public void DeleteSheet(string SheetName)
        {
            var workbookPart = Document.WorkbookPart;

            // Get the SheetToDelete from workbook.xml
            var theSheet = workbookPart.Workbook.Descendants<Sheet>()
                                       .FirstOrDefault(s => s.Name == SheetName);

            if (theSheet == null)
            {
                return;
            }

            // Remove the sheet reference from the workbook.
            var worksheetPart = GetWorksheetPartByName(SheetName);
            theSheet.Remove();

            // Delete the worksheet part.
            workbookPart.DeletePart(worksheetPart);

            Worksheets.RemoveAll(sh => sh.WorksheetName == SheetName);
        }

        public object? GetCustomProperty(string PropertyName)
        {
            object? res = default;
            var props = CustomProperties;
            if (props.ContainsKey(PropertyName))
            {
                res = props[PropertyName];

            }
            return res;
        }


        #endregion

        #region Private Variables
        private const string EX_BOOL_TRUE = "1";
        private const string EX_BOOL_FALSE = "1";
        private readonly MemoryStream? ms = null;
        #endregion

        #region Private Methods
        private void FillWorksheets()
        {
            if (Document != null)
            {
                var sheets = Document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                foreach (var sh in sheets)
                {
                    Worksheets.Add(new XlWorksheet(this, sh.Name));
                }
            }
        }

        private Cell? GetCell(string RefAddress)
        {
            var range = new XLRefAddress(this, RefAddress);
            var ws = GetWorkSheet(range.SheetName);
            if (ws == null)
                throw new Exception($"XLFile: Sheet with name {range.SheetName} is not found in workbook");
            return ws.GetCell(range.ColumnName1, range.RowIndex1);
        }

        private Cell? GetCell(string SheetName, uint ColumnIdx, uint RowIdx)
        {
            var ws = GetWorkSheet(SheetName);
            if (ws == null)
                throw new Exception($"XLFile: Sheet with name {SheetName} is not found in workbook");
            return ws.GetCell(ColumnIdx, RowIdx);
        }

        private void CheckDb()
        {
            if (Db == null)
                #pragma warning disable CA1303 // Do not pass literals as localized parameters
                throw new Exception("XLWorkbook: Db property must be set before using database operations");
                #pragma warning restore CA1303 // Do not pass literals as localized parameters
            if (!Db.IsConnected)
                Db.Open();

        }

        #endregion

    }
}
