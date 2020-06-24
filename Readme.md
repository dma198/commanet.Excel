## commanet.Excel
--------------------------
Library for reading/writing Excel documents.
Based on Microsoft OpenXML SDK. Supports only XML-based formats:  *\*.xlsx* , *\*.xlsm*

#### Example of usage

```C#
    // Open existed Excel file
    var xls = XLWorkbook.Open(MyExcelFilePath);
    // Set cell value 
    xls["A1"] = "Hello Excel World!";
    // Close file (save) 
    xls.Close();
```

#### Cells/Ranges addressing

Full format of cells range adddress:

*[SheetName]![$]Column1[$]Row1[:[$]Column2[$]Row2]*

As address can be used named area.

Examples:

Address        | Converted to full addresss  
---------------|--------------------------------
A1             | Sheet1!A1:A1
A1:A2          | Sheet1!A1:A2
My Sheet!A1:C4 | 'My Sheet'!A1:C4
MyNamedArea    | *Address extracted from named area definition*


#### XLWorkbook class

All manipulations with Excel file perfromed by methods of this class.

**Document Open/Close/Save operations:**
```C#

// Create new workbook in memory 
public XLWorkbook(SQLDBConnection db = null);

// Open existed Excel file        
public static XLWorkbook Open(string FileName, bool Editable = true, SQLDBConnection db = null)

// Open Excel file in stream
public static XLWorkbook Open(Stream stream, bool Editable = true, SQLDBConnection db = null)

// Close Excel file (autosave changes)
public void Close()

// Save changes
public void Save()

// Save changes in another file
public void SaveAs(string FilePath)
```

**Access to cells:**
```C#
var xls = XLWorkbook.Open(MyExcelFilePath);

public string this[string RefAddress]{get;set;}
// Example
xls["A1"] = "Hello World";

public void SetCellValue<T>(string RefAddr, T value);
public void SetCellValue<T>(string SheetName, uint ColumnIdx, uint RowIdx, T value);
// Examples
xls.SetCellValue<int>("A1",123);
xls.SetCellValue<double>("Sheet2",1,1,123.456);

public void GetCellValue<T>(string RefAddr);
public void GetCellValue<T>(string SheetName, uint ColumnIdx, uint RowIdx);
// Examples
int iv = xls.SetCellValue<int>("A1");
int id = xls.SetCellValue<double>("A2");

// Filling cells range from 2-d jagged array
public void FillArea(string RefAddress, object[][] data, bool Extend = true, bool Transposed = false);
// Examples
var data = new int[][] { 
  new int[] { 11, 12, 13 }, 
  new int[] { 21, 22, 23 }, 
  new int[] { 31, 32, 33 } };
xls.FillArea("ExtendedGrid", data);

xls.FillArea("Sheet2!A1:C3", data);

//Filling cells from object properties
public void FillCells<T>(T obj);

//Example
public class TestClass
{
    public int Field1 { get; set; }
    public int Field2 { get; set; }
    public int Field3 { get; set; }
}

var xls = XLWorkbook.Open(outfile);
var data = new TestClass() { 
        Field1 = 11, 
        Field2 = 12, 
        Field3 = 13};

xls.FillCellss(data);

xls.Close();

//Filling cells from dictionary
public void FillCells(Dictionary<string,object> data);

//Example
var xls = XLWorkbook.Open(outfile);

var data = new Dictionary<string, object>();
data.Add("Field1", 11);
data.Add("Field2", 12);
data.Add("Field3", 13);

xls.FillCells(data);

xls.Close();
```

**Database Connection**

Database connection (optional) can be accessed by *XLWorkbook.Db* property.
Methods for filleing data like FillArea, FillCells use it natively.

**FillArea options**

*Extend* option. 

With this option provided cells area address should refer for one row of data.
For next data rows in worksheet will be inserted new rows after given area.
Formatting of new cells will be copied from original area cells.
Existed cells after given area will be shifted.

*Transposed* option.

If this option is *False* then filling cells area will be from Up to Down.
Otherwise data will be filled from Left to Right.

Example:
```C#
var data = new int[][] { 
  new int[] { 11, 12, 13 }, 
  new int[] { 21, 22, 23 }, 
  new int[] { 31, 32, 33 } };
// Not Transposed
xls.FillArea("A1:C3", data,false,false);
// Result in Excel
//    A   B   C
// 1  11  12 13
// 2  21  22 23 
// 3  31  32 33 

// Transposed
xls.FillArea("A1:C3", data,false,true);
// Result in Excel
//    A   B   C
// 1  11  21 31
// 2  12  22 32 
// 3  13  23 33 

``` 




**Worksheets operations**
```C#
// Set active (selected) sheet 
public void SetActiveSheet(string SheetName);
```

**Filling cells from SQL Database**
Database connection optionally can be passes in *XLWorkbook* constructor. 
If connection is provided then can be used methods for filling cells from SQL queries.

FillArea Method: 
```C#
public void FillArea(string RefAddress,string SQL,bool Extend = true, bool Transposed = false, params KeyValuePair<string,object>[] SqlParams);
// Example

db = new SQLDBConnection("sqlite", "", "", ":memory:");
db.Transaction(th =>
{
    th.ExecuteNonQuery("CREATE TABLE test(c1 INTEGER,c2 INTEGER, c3 INTEGER)");
    th.ExecuteNonQuery("INSERT INTO test VALUES(11,12,13)");
    th.ExecuteNonQuery("INSERT INTO test VALUES(21,22,23)");
    th.ExecuteNonQuery("INSERT INTO test VALUES(31,32,33)");
});

var xls = XLWorkbook.Open(outfile,true,db);
var SQL = "SELECT c1,c2,c3 FROM test ORDER BY c1";

// Fill named cells range
xls.FillArea("MyCellsRangeName",SQL);

// Fill addressed cells range
xls.FillArea("Sheet2!A1:C3",SQL);

xls.Close();

```

FillCells Metod:
```C#
public void FillCells(string SQL, params KeyValuePair<string, object>[] SqlParams);

//Example:
var xls = XLWorkbook.Open(outfile, true, db);

var SQL = "SELECT c1 AS G3, c2 AS C30, c3 AS \"Second Sheet!C1\" FROM test ORDER BY c1";
xls.FillCells(SQL);

var SQL = @"SELECT c1 AS MyNamedCell1,
                   c2 AS MyNamedCell2, 
                   c3 AS NyNamedCell3 
            FROM test ORDER BY c1";
xls.FillCells(SQL);

xls.Close();

```











