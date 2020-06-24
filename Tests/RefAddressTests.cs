using System;
using System.IO;
using System.Reflection;
using Xunit;

using commanet.Excel;

namespace OpenXML.Tests
{
    public class RefAddressTests
    {
        [Fact]
        public void TestAdresses()
        {
            // GetColumnIndex
            var idx = XLRefAddress.GetColumnIndex("A");
            if (idx != 1) throw new Exception("Address Error");
            idx = XLRefAddress.GetColumnIndex("AA");
            if (idx != 27) throw new Exception("Address Error");

            // Get
            var colName = XLRefAddress.GetColumnName(1);
            if (colName != "A") throw new Exception("Address Error");
            colName = XLRefAddress.GetColumnName(26);
            if (colName != "Z") throw new Exception("Address Error");
            colName = XLRefAddress.GetColumnName(27);
            if (colName != "AA") throw new Exception("Address Error");
        }

        [Fact]
        public void TestColumnNameIndexConversion()
        {
            var idx = XLRefAddress.GetColumnIndex("Z");
            if(idx != 26)
                throw new Exception("Column Name=>Index Conversion Error");

            idx = XLRefAddress.GetColumnIndex("AZ");
            if (idx != 52)
                throw new Exception("Column Name=>Index Conversion Error");

            idx = XLRefAddress.GetColumnIndex("ABZ");
            if (idx != 754)
                throw new Exception("Column Name=>Index Conversion Error");


            var cname = XLRefAddress.GetColumnName(26);
            if (cname != "Z")
                throw new Exception("Column Index=>Name Conversion Error");

            cname = XLRefAddress.GetColumnName(52);
            if (cname != "AZ")
                throw new Exception("Column Index=>Name Conversion Error");

            cname = XLRefAddress.GetColumnName(53);
            if (cname != "BA")
                throw new Exception("Column Index=>Name Conversion Error");

            cname = XLRefAddress.GetColumnName(754);
            if (cname != "ABZ")
                throw new Exception("Column Index=>Name Conversion Error");

            cname = XLRefAddress.GetColumnName(79);
            if (cname != "CA")
                throw new Exception("Column Index=>Name Conversion Error");

            cname = XLRefAddress.GetColumnName(677);
            if (cname != "ZA")
                throw new Exception("Column Index=>Name Conversion Error");

            for (uint i=1;i<=16000;i++)
            {
                cname = XLRefAddress.GetColumnName(i);
                foreach(var c in cname)
                {
                    if(!char.IsLetter(c))
                    {
                        throw new Exception("Column Index=>Name Conversion Error. Corrupted character");
                    }
                }
                var cidx = XLRefAddress.GetColumnIndex(cname);
                if(cidx != i)
                    throw new Exception("Column Index=>Name Conversion Error. Unexpected column name");
            }

        }
    }
}
