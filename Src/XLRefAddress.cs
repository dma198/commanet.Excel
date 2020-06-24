using System;
using System.Text.RegularExpressions;
using System.Globalization;

namespace commanet.Excel
{
    public class XLRefAddress
    {
        public string SheetName { get; set; } = "";
        public uint RowIndex1 { get; set; } = 0;
        public string ColumnName1 { get; set; } = "";
        public uint RowIndex2 { get; set; } = 0;
        public string ColumnName2 { get; set; } = "";

        public uint ColumnIndex1
        {
            get => GetColumnIndex(ColumnName1);
            set => ColumnName1 = GetColumnName(value);
        }
        public uint ColumnIndex2
        {
            get => GetColumnIndex(ColumnName2);
            set => ColumnName2 = GetColumnName(value);
        }

        public XLRefAddress(XLWorkbook workbook, string address)
        {
            this.workbook = workbook;
            SheetName = "";
            DecodeRefAddr(address);
        }

        public static string GetColumnName(string cellName)
            => rxCol.Match(cellName).Value;
        public static uint GetRowIndex(string cellName)
            => uint.Parse(rxRow.Match(cellName).Value, CultureInfo.InvariantCulture);
        public static uint GetColumnIndex(string colName)
        {
            if (colName == null)
                throw new ArgumentNullException(nameof(colName));

            uint res = 0;
            for (int i = 0; i < colName.Length; i++)
            {
                var n = char.ToUpperInvariant(colName[i]) - 64;
                res += (uint)(n * Math.Pow(ALPHA_CNT, colName.Length - i - 1));
            }
            return res;
        }
        public static string GetColumnName(uint colIndex)
        {
            int idx = (int)colIndex - 1; 
            int qt = idx / 26;
            var s = char.ToUpperInvariant((char)(idx % 26 + 65))
                        .ToString(CultureInfo.InvariantCulture);
            if (qt > 0)
                return GetColumnName((uint)qt) + s;
            else
                return s;
        }

        public string RegerenceAddr
        {
            get
            {
                var addr = "";
                if (!string.IsNullOrEmpty(SheetName))
                    addr += (SheetName.Contains(' ', StringComparison.OrdinalIgnoreCase) ? $"'{SheetName}'" : SheetName) + "!";
                addr += RegerenceAddrNoSheet;
                return addr;
            }
        }



        public string RegerenceAddrFixedRowCols
        {
            get
            {
                var addr = "";
                if (!string.IsNullOrEmpty(SheetName.Trim()))
                    addr += (SheetName.Contains(' ', StringComparison.OrdinalIgnoreCase) ? $"'{SheetName}'" : SheetName) + "!";
                addr += RegerenceAddrNoSheetFixedRowCols;
                return addr;
            }
        }

        public string RegerenceAddrNoSheet
        {
            get
            {
                var addr = ColumnName1 + RowIndex1;
                if ((ColumnName1 != ColumnName2 && !string.IsNullOrEmpty(ColumnName2)) ||
                   (RowIndex1 != RowIndex2 && RowIndex2 > 0))
                {
                    addr += $":{ColumnName2}{RowIndex2}";
                }
                return addr;
            }
        }

        public string RegerenceAddrNoSheetFixedRowCols
        {
            get
            {
                var addr = '$' + ColumnName1 + '$' + RowIndex1;
                if ((ColumnName1 != ColumnName2 && !string.IsNullOrEmpty(ColumnName2)) ||
                   (RowIndex1 != RowIndex2 && RowIndex2 > 0))
                {
                    addr += $":${ColumnName2}${RowIndex2}";
                }
                return addr;
            }
        }

        private const uint ALPHA_CNT = 26;
        private readonly XLWorkbook workbook;
        private static readonly Regex rxSheet = new Regex("^[^!\n]+(?=![\\s\\S]*$)");
        private static readonly Regex rxCol = new Regex("\\$*[A-Z,a-z]+");
        private static readonly Regex rxRow = new Regex("\\$*[0-9]+");
        private const char CLIPADDR = '$';
        private const int MAX_COL = 16384;

        private void DecodeRefAddr(string RefAddr)
        {
            SheetName = "";
            RowIndex1 = 1;
            ColumnName1 = "A";
            RowIndex2 = 1;
            ColumnName2 = "A";

            var lRefAddr = RefAddr;
            var nAddr = workbook.GetNameDefRef(lRefAddr);
            if (nAddr != null)
            {
                lRefAddr = nAddr;
            }


            // Sheet token
            var m = rxSheet.Match(lRefAddr);
            if (m != null && m.Value != null && !string.IsNullOrEmpty(m.Value.Trim()))
                SheetName = m.Value.Trim(' ', '"', '\'');
            if (string.IsNullOrEmpty(SheetName))
            {
                var firstSheet = workbook.FirstWorkSheet;
                if (firstSheet != null) SheetName = firstSheet.WorksheetName;
            }

            //InSheet address 
            var ar = rxSheet.Split(lRefAddr);
            var sAddr = (ar.Length == 1 ? ar[0] : ar[1]).Trim(' ', '!');
            ar = sAddr.Split(':');
            if (ar.Length == 1)
            {
                ColumnName1 = rxCol.Match(sAddr).Value.Trim(CLIPADDR);
                ColumnName2 = ColumnName1;
                if (ColumnIndex1 > MAX_COL || !uint.TryParse(rxRow.Match(sAddr).Value.Trim(CLIPADDR), out uint rowidx))
                {
                    throw new Exception($"Address '{sAddr}' is wrong or named range not found");
                }
                else
                {
                    RowIndex1 = rowidx;
                    RowIndex2 = RowIndex1;
                }
            }
            else if (ar.Length == 2)
            {
                ColumnName1 = rxCol.Match(ar[0]).Value.Trim(CLIPADDR);
                ColumnName2 = rxCol.Match(ar[1]).Value.Trim(CLIPADDR);
                RowIndex1 = uint.Parse(rxRow.Match(ar[0]).Value.Trim(CLIPADDR), CultureInfo.InvariantCulture);
                RowIndex2 = uint.Parse(rxRow.Match(ar[1]).Value.Trim(CLIPADDR), CultureInfo.InvariantCulture);
            }
        }
    }
}