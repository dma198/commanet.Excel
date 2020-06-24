using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace commanet.Excel
{
    public class XlWorksheet
    {
        #region Public Properties
        public string WorksheetName { get; }
        #endregion

        #region Public Methods
        public XlWorksheet(XLWorkbook wb, string worksheetName)
        {
            workbook = wb;
            WorksheetName = worksheetName;
        }

        public List<MergeCell> MergedCells
        {
            get
            {
                var wsPart = workbook.GetWorksheetPartByName(WorksheetName);
                if (wsPart != null)
                {
                    var mCells = wsPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
                    return mCells != null ? mCells.Elements<MergeCell>().ToList()
                                         : new List<MergeCell>();
                }
                return new List<MergeCell>();
            }
        }

        public uint GetNextColumnIndex(uint RowIndex, uint ColumnIndex)
        {
            uint res = ColumnIndex + 1;
            foreach (var mc in MergedCells)
            {
                var maddr = new XLRefAddress(workbook, mc.Reference);
                if (RowIndex == maddr.RowIndex1 && ColumnIndex == maddr.ColumnIndex1)
                {
                    res = maddr.ColumnIndex2 + 1;
                    break;
                }
            }
            return res;
        }

        public uint GetNextRowIndex(uint RowIndex, uint ColumnIndex)
        {
            uint res = RowIndex + 1;
            foreach (var mc in MergedCells)
            {
                var maddr = new XLRefAddress(workbook, mc.Reference);
                if (RowIndex == maddr.RowIndex1 && ColumnIndex == maddr.ColumnIndex1)
                {
                    res = maddr.RowIndex2 + 1;
                    break;
                }
            }
            return res;
        }

        public void InsertCellsBelow(uint RowIndex, uint ColumnIndex1, uint ColumnIndex2)
        {
            var newRowIndex = GetNextRowIndex(RowIndex, ColumnIndex1);
            var rowStep = (newRowIndex - RowIndex);
            var rowCnt = RowCount;

            for (var c = ColumnIndex1; c <= ColumnIndex2; c++)
            {
                ShiftMergedCellsDown(newRowIndex, rowStep, c);
                for (var r = rowCnt + rowStep; r >= newRowIndex; r--)
                {
                    var srcRow = r - rowStep;

                    var cellSrc = GetCell(c, srcRow);
                    if (cellSrc != null)
                    {
                        var newCell = (Cell)cellSrc.Clone();
                        ReplaceCell(c, r, newCell);
                    }
                }
                if (rowStep > 1)
                    CloneRowMergedCellsDown(RowIndex, c, RowIndex + rowStep);
                ShiftNamedRangesDown(newRowIndex, rowStep, c);
            }
            ExtendNamedRangeDown(RowIndex, ColumnIndex1, ColumnIndex2, rowStep);
        }

        public void InsertCellsRight(uint ColumnIndex, uint RowIndex1, uint RowIndex2)
        {
            var newColIndex = GetNextColumnIndex(RowIndex1, ColumnIndex);
            var colStep = (newColIndex - ColumnIndex);
            var colCnt = ColCount;


            for (var r = RowIndex1; r <= RowIndex2; r++)
            {
                ShiftMergedCellsRight(newColIndex, colStep, r);
                for (var c = colCnt + colStep; c >= newColIndex; c--)
                {
                    var srcCol = c - colStep;
                    var cellSrc = GetCell(srcCol, r);
                    if (cellSrc != null)
                    {
                        var newCell = (Cell)cellSrc.Clone();
                        ReplaceCell(c, r, newCell);
                    }
                }
                if (colStep > 1)
                    CloneRowMergedCellsRight(r, ColumnIndex, ColumnIndex + colStep);

                ShiftNamedRangesRight(newColIndex, colStep, r);
            }

            ExtendNamedRangeRight(ColumnIndex, RowIndex1, RowIndex2, colStep);
        }

        public string? GetNamedRangeRef(string Name)
        {
            var defNames = workbook.Document.WorkbookPart.Workbook.DefinedNames;
            if (defNames != null)
            {
                foreach (DefinedName dn in defNames)
                {
                    if (dn.Name == Name)
                        return dn.Text;
                }
            }
            return null;
        }

        public Cell? GetCell(uint columnIndex, uint rowIndex)
        {
            return GetCell(XLRefAddress.GetColumnName(columnIndex), rowIndex);
        }

        public Cell? GetCell(string columnName, uint rowIndex)
        {
            Row? row = GetRow(rowIndex);

            if (row == null)
                return null;

            Cell? cell = null;
            var cells = row.Elements<Cell>().Where(c => string.Compare
                            (c.CellReference.Value, $"{columnName}{rowIndex}",
                             true, CultureInfo.InvariantCulture) == 0);

            if (cells != null && !cells.Any())
            {
                cell = new Cell()
                {
                    CellReference = $"{columnName}{rowIndex}",
                    DataType = CellValues.String,
                };
                row.Append(cell);

                SortRowByCellReference(row);

            }
            else
            {
                cell = cells?.First();
            }

            return cell;
        }


        #endregion

        #region Private Fields/Properties
        private readonly XLWorkbook workbook;

        private uint RowCount
        {
            get
            {
                var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
                if (worksheetPart != null)
                {
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    uint rcnt = 0;
                    foreach (var row in rows)
                        if (row.RowIndex > rcnt) rcnt = row.RowIndex;
                    return rcnt;
                }
                return 0;
            }
        }

        private uint ColCount
        {
            get
            {
                var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
                if (worksheetPart != null)
                {
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    uint ccnt = 0;
                    foreach (var row in rows)
                    {
                        var lastCol = row.Elements<Cell>().Last();
                        var addr = new XLRefAddress(workbook, lastCol.CellReference);
                        var colIdx = addr.ColumnIndex2;
                        if (colIdx > ccnt) ccnt = colIdx;
                    }
                    return ccnt;
                }
                return 0;
            }
        }

        #endregion

        #region Private Methods
        private void CloneRowMergedCellsDown(uint RowIndex, uint ColumnIndex, uint newRowIndex)
        {
            var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
            if (worksheetPart == null) return;
            var rowStep = (newRowIndex - RowIndex);
            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells != null)
            {
                var mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue);
                foreach (var mc in mergeCellsList)
                {
                    var maddr = new XLRefAddress(workbook, mc.Reference);
                    if (maddr.RowIndex1 == RowIndex && maddr.ColumnIndex1 == ColumnIndex)
                    {
                        maddr.RowIndex1 += rowStep;
                        maddr.RowIndex2 += rowStep;
                        var newMerge = (MergeCell)mc.Clone();
                        newMerge.Reference = maddr.RegerenceAddrNoSheet;
                        mergeCells.Append(newMerge);
                    }
                }
            }
        }

        private void CloneRowMergedCellsRight(uint RowIndex, uint ColumnIndex, uint newColIndex)
        {
            var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
            if (worksheetPart == null) return;
            var colStep = (newColIndex - ColumnIndex);
            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells != null)
            {
                var mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue);
                foreach (var mc in mergeCellsList)
                {
                    var maddr = new XLRefAddress(workbook, mc.Reference);
                    if (maddr.RowIndex1 == RowIndex && maddr.ColumnIndex1 == ColumnIndex)
                    {
                        maddr.ColumnIndex1 += colStep;
                        maddr.ColumnIndex2 += colStep;
                        var newMerge = (MergeCell)mc.Clone();
                        newMerge.Reference = maddr.RegerenceAddrNoSheet;
                        mergeCells.Append(newMerge);
                    }
                }
            }
        }

        private void ShiftMergedCellsDown(uint newRowIndex, uint rowStep, uint ColIndex)
        {
            var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
            if (worksheetPart == null) return;
            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells != null)
            {
                var mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue);
                foreach (var mc in mergeCellsList)
                {
                    var maddr = new XLRefAddress(workbook, mc.Reference);
                    if (maddr.RowIndex1 >= newRowIndex && maddr.ColumnIndex1 == ColIndex)
                    {
                        maddr.RowIndex1 += rowStep;
                        maddr.RowIndex2 += rowStep;
                        mc.Reference = maddr.RegerenceAddrNoSheet;
                    }
                }
            }
        }

        private void ShiftMergedCellsRight(uint newColIndex, uint colStep, uint RowIndex)
        {
            var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
            if (worksheetPart == null) return;
            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells != null)
            {
                var mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue);
                foreach (var mc in mergeCellsList)
                {
                    var maddr = new XLRefAddress(workbook, mc.Reference);
                    if (maddr.ColumnIndex1 >= newColIndex && maddr.RowIndex1 == RowIndex)
                    {
                        maddr.ColumnIndex1 += colStep;
                        maddr.ColumnIndex2 += colStep;
                        mc.Reference = maddr.RegerenceAddrNoSheet;
                    }
                }
            }
        }

        private void ExtendNamedRangeDown(uint RowIndex, uint ColIndex1, uint ColIndex2, uint rowStep)
        {
            var defNames = workbook.Document.WorkbookPart.Workbook.DefinedNames;
            if (defNames != null)
            {
                foreach (DefinedName dn in defNames)
                {
                    if (dn.Text == null || string.IsNullOrEmpty(dn.Text.Trim())) continue;
                    var maddr = new XLRefAddress(workbook, dn.Text);
                    if (maddr.SheetName != WorksheetName) continue;
                    if (RowIndex >= maddr.RowIndex1 && RowIndex <= maddr.RowIndex2 &&
                        ColIndex1 == maddr.ColumnIndex1 && ColIndex2 == maddr.ColumnIndex2)
                    {
                        maddr.RowIndex2 += rowStep;
                        dn.Text = maddr.RegerenceAddrFixedRowCols;
                    }
                }
            }
        }

        private void ExtendNamedRangeRight(uint ColIndex, uint RowIndex1, uint RowIndex2, uint colStep)
        {
            var defNames = workbook.Document.WorkbookPart.Workbook.DefinedNames;
            if (defNames != null)
            {
                foreach (DefinedName dn in defNames)
                {
                    if (dn.Text == null || string.IsNullOrEmpty(dn.Text.Trim())) continue;
                    var maddr = new XLRefAddress(workbook, dn.Text);
                    if (maddr.SheetName != WorksheetName) continue;
                    if (ColIndex >= maddr.ColumnIndex1 && ColIndex <= maddr.ColumnIndex2 &&
                        RowIndex1 == maddr.RowIndex1 && RowIndex2 == maddr.RowIndex2)
                    {
                        maddr.ColumnIndex2 += colStep;
                        dn.Text = maddr.RegerenceAddrFixedRowCols;
                    }
                }
            }
        }

        private void ShiftNamedRangesDown(uint newRowIndex, uint rowStep, uint ColIndex)
        {
            var defNames = workbook.Document.WorkbookPart.Workbook.DefinedNames;
            if (defNames != null)
            {
                foreach (DefinedName dn in defNames)
                {
                    if (dn.Text == null || string.IsNullOrEmpty(dn.Text.Trim())) continue;
                    var maddr = new XLRefAddress(workbook, dn.Text);
                    if (maddr.SheetName != WorksheetName) continue;
                    if (maddr.RowIndex1 >= newRowIndex && maddr.ColumnIndex1 == ColIndex)
                    {
                        maddr.RowIndex1 += rowStep;
                        maddr.RowIndex2 += rowStep;
                        dn.Text = maddr.RegerenceAddrFixedRowCols;
                    }
                }
            }

        }

        private void ShiftNamedRangesRight(uint newColIndex, uint colStep, uint RowIndex)
        {
            var defNames = workbook.Document.WorkbookPart.Workbook.DefinedNames;
            if (defNames != null)
            {
                foreach (DefinedName dn in defNames)
                {
                    if (dn.Text == null || string.IsNullOrEmpty(dn.Text.Trim())) continue;
                    var maddr = new XLRefAddress(workbook, dn.Text);
                    if (maddr.SheetName != WorksheetName) continue;
                    if (maddr.ColumnIndex1 >= newColIndex && maddr.RowIndex1 == RowIndex)
                    {
                        maddr.ColumnIndex1 += colStep;
                        maddr.ColumnIndex2 += colStep;
                        dn.Text = maddr.RegerenceAddrFixedRowCols;
                    }
                }
            }
        }

        private Row? GetRow(uint rowIndex)
        {
            var worksheetPart = workbook.GetWorksheetPartByName(WorksheetName);
            if (worksheetPart == null) return null;
            var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            var rowFound = worksheetPart.Worksheet.GetFirstChild<SheetData>().
                   Elements<Row>().Where(r => r.RowIndex == rowIndex);

            Row row;

            if (!rowFound.Any())
            {
                row = InsertRow(rowIndex, worksheetPart, null);
            }
            else
            {
                row = rowFound.First();
            }

            return row;
        }

        private void ReplaceCell(uint columnIndex, uint rowIndex, Cell newCell)
        {
            ReplaceCell(XLRefAddress.GetColumnName(columnIndex), rowIndex, newCell);
        }

        private void ReplaceCell(string columnName, uint rowIndex, Cell newCell)
        {
            newCell.CellReference = columnName + rowIndex;

            Row? row = GetRow(rowIndex);

            if (row == null) return;

            var cells = row.Elements<Cell>().Where(c => string.Compare
                            (c.CellReference.Value, columnName +
                             rowIndex, true, CultureInfo.InvariantCulture) == 0);

            Cell? oldCell = null;
            var toBeSorted = false;
            if (cells.Any()) oldCell = cells.First();
            if (oldCell == null)
            {
                row.Append(newCell);
                toBeSorted = true;
            }
            else
            {
                row.ReplaceChild(newCell, oldCell);
            }

            if (toBeSorted) SortRowByCellReference(row);
        }

        private static void SortRowByCellReference(Row row)
        {
            var sorted = row.Elements<Cell>().OrderBy(c => XLRefAddress.GetColumnIndex(c.CellReference)).ToArray();
            row.RemoveAllChildren<Cell>();
            row.Append(sorted);
        }

        private static Row InsertRow(uint rowIndex, WorksheetPart worksheetPart, Row? insertRow, bool isNewLastRow = false)
        {

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            Row? retRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex) : null;

            // If the worksheet does not contain a row with the specified row index, insert one.
            if (retRow != null)
            {
                // if retRow is not null and we are inserting a new row, then move all existing rows down.
                if (insertRow != null)
                {
                    UpdateRowIndexes(worksheetPart, rowIndex, false);
                    //UpdateMergedCellReferences(worksheetPart, rowIndex, false);
                    //UpdateHyperlinkReferences(worksheetPart, rowIndex, false);

                    // actually insert the new row into the sheet
                    retRow = sheetData.InsertBefore(insertRow, retRow);  // at this point, retRow still points to the row that had the insert rowIndex

                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString(CultureInfo.InvariantCulture);

                    foreach (Cell cell in retRow.Elements<Cell>())
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex, StringComparison.OrdinalIgnoreCase));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }
            }
            else
            {
                // Row doesn't exist yet, shifting not needed.
                // Rows must be in sequential order according to RowIndex. Determine where to insert the new row.
                Row? refRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex > rowIndex) : null;

                // use the insert row if it exists
                retRow = insertRow ?? new Row() { RowIndex = rowIndex };

                IEnumerable<Cell> cellsInRow = retRow.Elements<Cell>();

                if (cellsInRow.Any())
                {
                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString(CultureInfo.InvariantCulture);

                    foreach (Cell cell in cellsInRow)
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex, StringComparison.OrdinalIgnoreCase));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }

                sheetData.InsertBefore(retRow, refRow);
            }

            return retRow;
        }

        private static void UpdateRowIndexes(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {

            var rows = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= rowIndex);

            foreach (var row in rows)
            {
                var newIndex = (isDeletedRow ? row.RowIndex - 1 : row.RowIndex + 1);
                var curRowIndex = row.RowIndex.ToString();
                var newRowIndex = newIndex.ToString(CultureInfo.InvariantCulture);

                foreach (var cell in row.Elements<Cell>())
                {
                    // Update the references for the rows cells.
                    cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curRowIndex, newRowIndex, StringComparison.OrdinalIgnoreCase));
                }

                // Update the row index.
                row.RowIndex = newIndex;
            }
        }
        #endregion

    }
}
