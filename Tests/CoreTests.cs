using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using Xunit;

using commanet.Excel;

namespace OpenXML.Tests
{
    public class CoreTests
    {
        [Fact]
        public void TestReadWriteCells()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);

            xls["A1"] = "Test A1";        // Cell is exists in original template            
            xls["A100"] = "Test A100";    // Cell is in row which does not exists in original template 
            xls["BB1"] = "Test BB1";      // Cell is in col which does not exists in original template 
            xls["Second Sheet!A1"] = "Second Sheet Test";

            xls.Save();
            xls.Close();

            // Check result

            xls = XLWorkbook.Open(outfile, false);
            if (xls["A1"] != "Test A1")
                throw new Exception("Value not saved");
            if (xls["A100"] != "Test A100")
                throw new Exception("Value not saved");
            if (xls["BB1"] != "Test BB1")
                throw new Exception("Value not saved");

            if (xls["Second Sheet!A1"] != "Second Sheet Test")
                throw new Exception("Can't read second sheet data");


            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }


        [Fact]
        public void TestCellTypes()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            xls.SetActiveSheet("Sheet1");

            xls["A1"] = "String";

            DateTime dt = DateTime.Now;
            xls.SetCellValue("B5", dt);
            xls.SetCellValue("A3", true);
            xls.SetCellValue("A4", 123);
            xls.SetCellValue("B4", 123.456);

            xls.Save();
            xls.Close();

            // Check result

            xls = XLWorkbook.Open(outfile, false);
            if (xls["A1"] != "String")
                throw new Exception("Value not saved");
            if (xls.GetCellValue<DateTime>("B5").CompareTo(dt) == 0)
                throw new Exception("Value not saved");

            if (xls.GetCellValue<bool>("A3") != true)
                throw new Exception("Value not saved");
            if (xls.GetCellValue<int>("A4") != 123)
                throw new Exception("Value not saved");
            if (xls.GetCellValue<double>("B4") != 123.456)
                throw new Exception("Value not saved");

            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestDefinedNames()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            xls["DefinedName"] = "TestDefinedName";
            xls.Save();
            xls.Close();
            // Check result
            xls = XLWorkbook.Open(outfile, false);
            if (xls["DefinedName"] != "TestDefinedName")
                throw new Exception("Value not saved");
            xls.Close();
            if (File.Exists(outfile)) File.Delete(outfile);
        }


        [Fact]
        public void TestFillAreaNotExtendedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] { 
                new object[]{ 11, 12, 13 },
                new object[]{ 21, 22, 23 },
                new object[]{ 31, 32, 33 } 
            };
            xls.FillArea("FixedGrid", data, false);
            xls.Save();
            xls.Close();
            // Check result
            xls = XLWorkbook.Open(outfile, false);
            if (xls.GetCellValue<int>("A8") != 11 || xls.GetCellValue<int>("B8") != 12 || xls.GetCellValue<int>("C8") != 13 ||
                xls.GetCellValue<int>("A9") != 21 || xls.GetCellValue<int>("B9") != 22 || xls.GetCellValue<int>("C9") != 23 ||
                xls.GetCellValue<int>("A10") != 31 || xls.GetCellValue<int>("B10") != 32 || xls.GetCellValue<int>("C10") != 33)
                throw new Exception("Value not saved not correct");

            xls.Close();
            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillAreaNotExtendedMergedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] { 
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 } 
            };
            xls.FillArea("FixedGridMergedCells", data, false);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("E8") != 11 || xls.GetCellValue<int>("G8") != 12 || xls.GetCellValue<int>("I8") != 13 ||
                xls.GetCellValue<int>("E10") != 21 || xls.GetCellValue<int>("G10") != 22 || xls.GetCellValue<int>("I10") != 23 ||
                xls.GetCellValue<int>("E12") != 31 || xls.GetCellValue<int>("G12") != 32 || xls.GetCellValue<int>("I12") != 33)
                throw new Exception("Value not saved not correct");

            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }


        [Fact]
        public void TestFillAreaExtendedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 } };
            xls.FillArea("ExtendedGrid", data);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("A16") != 11 || xls.GetCellValue<int>("B16") != 12 || xls.GetCellValue<int>("C16") != 13 ||
                xls.GetCellValue<int>("A17") != 21 || xls.GetCellValue<int>("B17") != 22 || xls.GetCellValue<int>("C17") != 23 ||
                xls.GetCellValue<int>("A18") != 31 || xls.GetCellValue<int>("B18") != 32 || xls.GetCellValue<int>("C18") != 33 ||
                // Check cells at right of extended grid to be sure that they are not destroyed
                xls.GetCellValue<string>("D16") != "R1" || xls.GetCellValue<string>("D17") != "R2" ||
                // Check if cells at the bottom of extended grid 
                xls.GetCellValue<string>("A19") != "Below Extended Grid")
                throw new Exception("Value not saved not correct");

            xls.Save();
            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillAreaExtendedSecondSheetArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] { 
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 } };
            xls.FillArea("Second Sheet!A4:C4", data);
            xls.Save();
            xls.Close();
            // Check result
            
            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("Second Sheet!A4") != 11 || xls.GetCellValue<int>("Second Sheet!B4") != 12 || xls.GetCellValue<int>("Second Sheet!C4") != 13 ||
                xls.GetCellValue<int>("Second Sheet!A5") != 21 || xls.GetCellValue<int>("Second Sheet!B5") != 22 || xls.GetCellValue<int>("Second Sheet!C5") != 23 ||
                xls.GetCellValue<int>("Second Sheet!A6") != 31 || xls.GetCellValue<int>("Second Sheet!B6") != 32 || xls.GetCellValue<int>("Second Sheet!C6") != 33)
                throw new Exception("Value not saved not correct");

            xls.Save();
            xls.Close();
            
            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillAreaExtendedMergedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";


            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 } };
            xls.FillArea("ExtendedGridMergedCells", data);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("E16") != 11 || xls.GetCellValue<int>("G16") != 12 || xls.GetCellValue<int>("I16") != 13 ||
                xls.GetCellValue<int>("E18") != 21 || xls.GetCellValue<int>("G18") != 22 || xls.GetCellValue<int>("I18") != 23 ||
                xls.GetCellValue<int>("E20") != 31 || xls.GetCellValue<int>("G20") != 32 || xls.GetCellValue<int>("I20") != 33 ||
                // Check cells at right of extended grid to be sure that they are not destroyed
                xls.GetCellValue<string>("K16") != "R3" || xls.GetCellValue<string>("K18") != "R4" || xls.GetCellValue<string>("K19") != "R5" ||
                // Check if cells at the bottom of extended grid 
                xls.GetCellValue<string>("E22") != "Below Extended Grid 2")
                throw new Exception("Value not saved not correct");

            xls.Save();
            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }


        [Fact]
        public void TestFillAreaFixedTransposedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 },
                new object[] { 41, 42, 43 } // Specially made overflow of target area - to be not filled 
            };
            xls.FillArea("FixedGridTransposed", data,false,true);
            xls.Save();
            xls.Close();
            // Check result
            
            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("A22") != 11 || xls.GetCellValue<int>("B22") != 21 || xls.GetCellValue<int>("C22") != 31 ||
                xls.GetCellValue<int>("A23") != 12 || xls.GetCellValue<int>("B23") != 22 || xls.GetCellValue<int>("C23") != 32 ||
                xls.GetCellValue<int>("A24") != 13 || xls.GetCellValue<int>("B24") != 23 || xls.GetCellValue<int>("C24") != 33 ||
                // Check cells at right of extended grid to be sure that they are not destroyed
                xls.GetCellValue<string>("D22") != "R7" || xls.GetCellValue<string>("D23") != "R8" || xls.GetCellValue<string>("D24") != "R9")
                throw new Exception("Value not saved not correct");
            xls.Save();
            xls.Close();
            
            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillAreaFixedTransposedMergedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 },
                new object[] { 41, 42, 43 } // Specially made overflow of target area - to be not filled 
            };
            xls.FillArea("FixedGridTransposedMerged", data, false, true);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("E22") != 11 || xls.GetCellValue<int>("G22") != 21 || xls.GetCellValue<int>("I22") != 31 ||
                xls.GetCellValue<int>("E24") != 12 || xls.GetCellValue<int>("G24") != 22 || xls.GetCellValue<int>("I24") != 32 ||
                xls.GetCellValue<int>("E26") != 13 || xls.GetCellValue<int>("G26") != 23 || xls.GetCellValue<int>("I26") != 33 ||
                // Check cells at right of extended grid to be sure that they are not destroyed
                xls.GetCellValue<string>("K22") != "R10" || xls.GetCellValue<string>("K23") != "R11" || xls.GetCellValue<string>("K24") != "R12")
                throw new Exception("Value not saved not correct");
            xls.Save();
            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }
        
        [Fact]
        public void TestFillAreaExtendedTransposedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 }};

            xls.FillArea("ExtendedGridTransposed", data, true, true);
            xls.Save();
            xls.Close();
            // Check result
            
            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("A30") != 11 || xls.GetCellValue<int>("B30") != 21 || xls.GetCellValue<int>("C30") != 31 ||
                xls.GetCellValue<int>("A31") != 12 || xls.GetCellValue<int>("B31") != 22 || xls.GetCellValue<int>("C31") != 32 ||
                xls.GetCellValue<int>("A32") != 13 || xls.GetCellValue<int>("B32") != 23 || xls.GetCellValue<int>("C32") != 33 ||
                // Check cells at right of extended grid to be sure that they are not destroyed
                xls.GetCellValue<string>("D30") != "R14" || xls.GetCellValue<string>("D31") != "R15" || xls.GetCellValue<string>("D32") != "R16" ||
                // Check below of grid 
                xls.GetCellValue<string>("A33") != "R17")
                throw new Exception("Value not saved not correct");
            xls.Save();
            xls.Close();
            
            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillAreaExtendedTransposedMergedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 } };
            xls.FillArea("ExtendedGridTransposedMerged", data, true, true);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);
            
            if (xls.GetCellValue<int>("A36") != 11 || xls.GetCellValue<int>("C36") != 21 || xls.GetCellValue<int>("E36") != 31 ||
                xls.GetCellValue<int>("A38") != 12 || xls.GetCellValue<int>("C38") != 22 || xls.GetCellValue<int>("E38") != 32 ||
                xls.GetCellValue<int>("A40") != 13 || xls.GetCellValue<int>("C40") != 23 || xls.GetCellValue<int>("E40") != 33 ||
                // Check cells at right of extended grid to be sure that they are not destroyed
                xls.GetCellValue<string>("G36") != "R18" || xls.GetCellValue<string>("G38") != "R19" || xls.GetCellValue<string>("G40") != "R20" ||
                // Check below of grid 
                xls.GetCellValue<string>("A42") != "R21")
                throw new Exception("Value not saved not correct");
            xls.Save();
            xls.Close();
            
            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillAreaCombinedArray()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new object[][] {
                new object[] { 11, 12, 13 },
                new object[] { 21, 22, 23 },
                new object[] { 31, 32, 33 } };
            xls.FillArea("FixedGrid", data, false);
            xls.FillArea("FixedGridMergedCells", data, false);
            xls.FillArea("ExtendedGrid", data);

            xls.FillArea("ExtendedGridMergedCells", data);            
            xls.FillArea("FixedGridTransposed", data, false, true);
            xls.FillArea("ExtendedGridTransposed", data, true, true);
            xls.FillArea("FixedGridTransposedMerged", data, false, true);                        
            xls.FillArea("ExtendedGridTransposedMerged", data, true, true);
            

            xls.Save();
            xls.Close();
            // Check result

            // Fixed Grid

            xls = XLWorkbook.Open(outfile,false);

            if (xls.GetCellValue<int>("A8") != 11 || xls.GetCellValue<int>("B8") != 12 || xls.GetCellValue<int>("C8") != 13 ||
                xls.GetCellValue<int>("A9") != 21 || xls.GetCellValue<int>("B9") != 22 || xls.GetCellValue<int>("C9") != 23 ||
                xls.GetCellValue<int>("A10")!= 31 || xls.GetCellValue<int>("B10") != 32 || xls.GetCellValue<int>("C10") != 33)
                throw new Exception("Fixed Grid is not correct");

            // Fixed Grid Merged Cells
            if (xls.GetCellValue<int>("E8") != 11 || xls.GetCellValue<int>("G8") != 12 || xls.GetCellValue<int>("I8") != 13 ||
                xls.GetCellValue<int>("E10") != 21 || xls.GetCellValue<int>("G10") != 22 || xls.GetCellValue<int>("I10") != 23 ||
                xls.GetCellValue<int>("E12") != 31 || xls.GetCellValue<int>("G12") != 32 || xls.GetCellValue<int>("I12") != 33)
                throw new Exception("Fixed Grid Merged Cells is not correct");

            // Extended Grid 
            if (xls.GetCellValue<int>("A16") != 11 || xls.GetCellValue<int>("B16") != 12 || xls.GetCellValue<int>("C16") != 13 ||
                xls.GetCellValue<int>("A17") != 21 || xls.GetCellValue<int>("B17") != 22 || xls.GetCellValue<int>("C17") != 23 ||
                xls.GetCellValue<int>("A18") != 31 || xls.GetCellValue<int>("B18") != 32 || xls.GetCellValue<int>("C18") != 33)
                throw new Exception("Extended Grid is not correct");

            // Extended Grid Merged Cells
            if (xls.GetCellValue<int>("E16") != 11 || xls.GetCellValue<int>("G16") != 12 || xls.GetCellValue<int>("I16") != 13 ||
                xls.GetCellValue<int>("E18") != 21 || xls.GetCellValue<int>("G18") != 22 || xls.GetCellValue<int>("I18") != 23 ||
                xls.GetCellValue<int>("E20") != 31 || xls.GetCellValue<int>("G20") != 32 || xls.GetCellValue<int>("I20") != 33)
                throw new Exception("Extended Grid Merged Cells is not correct");

            // Fixed Grid Transposed
            if (xls.GetCellValue<int>("A24") != 11 || xls.GetCellValue<int>("B24") != 21 || xls.GetCellValue<int>("C24") != 31 ||
                xls.GetCellValue<int>("A25") != 12 || xls.GetCellValue<int>("B25") != 22 || xls.GetCellValue<int>("C25") != 32 ||
                xls.GetCellValue<int>("A26") != 13 || xls.GetCellValue<int>("B26") != 23 || xls.GetCellValue<int>("C26") != 33)
                throw new Exception("Fixed Grid Transposed is not correct");

            // Fixed Grid Transposed Merged Cells
            if (xls.GetCellValue<int>("E26") != 11 || xls.GetCellValue<int>("G26") != 21 || xls.GetCellValue<int>("I26") != 31 ||
                xls.GetCellValue<int>("E28") != 12 || xls.GetCellValue<int>("G28") != 22 || xls.GetCellValue<int>("I28") != 32 ||
                xls.GetCellValue<int>("E30") != 13 || xls.GetCellValue<int>("G30") != 23 || xls.GetCellValue<int>("I30") != 33)
                throw new Exception("Fixed Grid Transposed Merged Cells is not correct");

            // Extended Grid Transposed
            if (xls.GetCellValue<int>("A32") != 11 || xls.GetCellValue<int>("B32") != 21 || xls.GetCellValue<int>("C32") != 31 ||
                xls.GetCellValue<int>("A33") != 12 || xls.GetCellValue<int>("B33") != 22 || xls.GetCellValue<int>("C33") != 32 ||
                xls.GetCellValue<int>("A34") != 13 || xls.GetCellValue<int>("B34") != 23 || xls.GetCellValue<int>("C34") != 33)
                throw new Exception("Extended Grid Transposed is not correct");

            // Extended Grid Transposed Merged Cells
            if (xls.GetCellValue<int>("A38") != 11 || xls.GetCellValue<int>("C38") != 21 || xls.GetCellValue<int>("E38") != 31 ||
                xls.GetCellValue<int>("A40") != 12 || xls.GetCellValue<int>("C40") != 22 || xls.GetCellValue<int>("E40") != 32 ||
                xls.GetCellValue<int>("A42") != 13 || xls.GetCellValue<int>("C42") != 23 || xls.GetCellValue<int>("E42") != 33)
                throw new Exception("Extended Grid Transposed Merged Cells is not correct");

            // Suround cells
            if(xls.GetCellValue<string>("D16") != "R1" || xls.GetCellValue<string>("D17") != "R2" ||
               xls.GetCellValue<string>("K16") != "R3" || xls.GetCellValue<string>("K18") != "R4" ||
               xls.GetCellValue<string>("K19") != "R5" || xls.GetCellValue<string>("D22") != "R7" ||
               xls.GetCellValue<string>("D23") != "R8" || xls.GetCellValue<string>("D24") != "R9" ||
               xls.GetCellValue<string>("K22") != "R10" || xls.GetCellValue<string>("K23") != "R11" ||
               xls.GetCellValue<string>("K24") != "R12" || xls.GetCellValue<string>("D32") != "R14" ||
               xls.GetCellValue<string>("D33") != "R15" || xls.GetCellValue<string>("D34") != "R16" ||
               xls.GetCellValue<string>("A35") != "R17" || xls.GetCellValue<string>("G38") != "R18" ||
               xls.GetCellValue<string>("G40") != "R19" || xls.GetCellValue<string>("G42") != "R20" ||
               xls.GetCellValue<string>("A44") != "R21")
                throw new Exception("Layout is not correct");


            xls.Close();
           
            if (File.Exists(outfile)) File.Delete(outfile);
        }


        private class TestClass
        {
            public int Field1 { get; set; }
            public int Field2 { get; set; }
            public int Field3 { get; set; }
        }

        [Fact]
        public void TestFillCellsFromObject()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new TestClass() { Field1 = 11, Field2 = 12, Field3 = 13};
            xls.FillCells(data);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("Field1") != 11 || xls.GetCellValue<int>("Field2") != 12 || xls.GetCellValue<int>("Field3") != 13)
                throw new Exception("Value not saved not correct");
            xls.Save();
            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }

        [Fact]
        public void TestFillCellsFromDictionary()
        {
            var fname = Path.GetFullPath(Path.Combine(
                                   Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                                   "..", "..", "..", "TestData", "OpenXmlTest.xlsx"));

            var outfile = Path.GetTempFileName() + ".xlsx";
            //var outfile = @"C:\tmp\OUT.xlsx";

            if (File.Exists(outfile)) File.Delete(outfile);

            // Perform data manipulation
            File.Copy(fname, outfile);
            var xls = XLWorkbook.Open(outfile);
            var data = new Dictionary<string, object>
            {
                { "Field1", 11 },
                { "Field2", 12 },
                { "Field3", 13 }
            };
            xls.FillCells(data);
            xls.Save();
            xls.Close();
            // Check result

            xls = XLWorkbook.Open(outfile, false);

            if (xls.GetCellValue<int>("Field1") != 11 || xls.GetCellValue<int>("Field2") != 12 || xls.GetCellValue<int>("Field3") != 13)
                throw new Exception("Value not saved not correct");
            xls.Save();
            xls.Close();

            if (File.Exists(outfile)) File.Delete(outfile);
        }


    }

}


