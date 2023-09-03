// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    [SetCulture("en-US")]
    public class LookupTests
    {
        private IXLWorksheet ws;

        #region Setup and teardown

        [OneTimeTearDown]
        public void Dispose()
        {
            ws.Workbook.Dispose();
        }

        [SetUp]
        public void Init()
        {
            ws = SetupWorkbook();
        }

        private IXLWorksheet SetupWorkbook()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            var data = new object[]
            {
                new {Id=1, OrderDate = DateTime.Parse("2015-01-06"), Region = "East", Rep = "Jones", Item = "Pencil", Units = 95, UnitCost = 1.99, Total = 189.05 },
                new {Id=2, OrderDate = DateTime.Parse("2015-01-23"), Region = "Central", Rep = "Kivell", Item = "Binder", Units = 50, UnitCost = 19.99, Total = 999.5},
                new {Id=3, OrderDate = DateTime.Parse("2015-02-09"), Region = "Central", Rep = "Jardine", Item = "Pencil", Units = 36, UnitCost = 4.99, Total = 179.64},
                new {Id=4, OrderDate = DateTime.Parse("2015-02-26"), Region = "Central", Rep = "Gill", Item = "Pen", Units = 27, UnitCost = 19.99, Total = 539.73},
                new {Id=5, OrderDate = DateTime.Parse("2015-03-15"), Region = "West", Rep = "Sorvino", Item = "Pencil", Units = 56, UnitCost = 2.99, Total = 167.44},
                new {Id=6, OrderDate = DateTime.Parse("2015-04-01"), Region = "East", Rep = "Jones", Item = "Binder", Units = 60, UnitCost = 4.99, Total = 299.4},
                new {Id=7, OrderDate = DateTime.Parse("2015-04-18"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 75, UnitCost = 1.99, Total = 149.25},
                new {Id=8, OrderDate = DateTime.Parse("2015-05-05"), Region = "Central", Rep = "Jardine", Item = "Pencil", Units = 90, UnitCost = 4.99, Total = 449.1},
                new {Id=9, OrderDate = DateTime.Parse("2015-05-22"), Region = "West", Rep = "Thompson", Item = "Pencil", Units = 32, UnitCost = 1.99, Total = 63.68},
                new {Id=10, OrderDate = DateTime.Parse("2015-06-08"), Region = "East", Rep = "Jones", Item = "Binder", Units = 60, UnitCost = 8.99, Total = 539.4},
                new {Id=11, OrderDate = DateTime.Parse("2015-06-25"), Region = "Central", Rep = "Morgan", Item = "Pencil", Units = 90, UnitCost = 4.99, Total = 449.1},
                new {Id=12, OrderDate = DateTime.Parse("2015-07-12"), Region = "East", Rep = "Howard", Item = "Binder", Units = 29, UnitCost = 1.99, Total = 57.71},
                new {Id=13, OrderDate = DateTime.Parse("2015-07-29"), Region = "East", Rep = "Parent", Item = "Binder", Units = 81, UnitCost = 19.99, Total = 1619.19},
                new {Id=14, OrderDate = DateTime.Parse("2015-08-15"), Region = "East", Rep = "Jones", Item = "Pencil", Units = 35, UnitCost = 4.99, Total = 174.65},
                new {Id=15, OrderDate = DateTime.Parse("2015-09-01"), Region = "Central", Rep = "Smith", Item = "Desk", Units = 2, UnitCost = 125, Total = 250},
                new {Id=16, OrderDate = DateTime.Parse("2015-09-18"), Region = "East", Rep = "Jones", Item = "Pen Set", Units = 16, UnitCost = 15.99, Total = 255.84},
                new {Id=17, OrderDate = DateTime.Parse("2015-10-05"), Region = "Central", Rep = "Morgan", Item = "Binder", Units = 28, UnitCost = 8.99, Total = 251.72},
                new {Id=18, OrderDate = DateTime.Parse("2015-10-22"), Region = "East", Rep = "Jones", Item = "Pen", Units = 64, UnitCost = 8.99, Total = 575.36},
                new {Id=19, OrderDate = DateTime.Parse("2015-11-08"), Region = "East", Rep = "Parent", Item = "Pen", Units = 15, UnitCost = 19.99, Total = 299.85},
                new {Id=20, OrderDate = DateTime.Parse("2015-11-25"), Region = "Central", Rep = "Kivell", Item = "Pen Set", Units = 96, UnitCost = 4.99, Total = 479.04},
                new {Id=21, OrderDate = DateTime.Parse("2015-12-12"), Region = "Central", Rep = "Smith", Item = "Pencil", Units = 67, UnitCost = 1.29, Total = 86.43},
                new {Id=22, OrderDate = DateTime.Parse("2015-12-29"), Region = "East", Rep = "Parent", Item = "Pen Set", Units = 74, UnitCost = 15.99, Total = 1183.26},
                new {Id=23, OrderDate = DateTime.Parse("2016-01-15"), Region = "Central", Rep = "Gill", Item = "Binder", Units = 46, UnitCost = 8.99, Total = 413.54},
                new {Id=24, OrderDate = DateTime.Parse("2016-02-01"), Region = "Central", Rep = "Smith", Item = "Binder", Units = 87, UnitCost = 15, Total = 1305},
                new {Id=25, OrderDate = DateTime.Parse("2016-02-18"), Region = "East", Rep = "Jones", Item = "Binder", Units = 4, UnitCost = 4.99, Total = 19.96},
                new {Id=26, OrderDate = DateTime.Parse("2016-03-07"), Region = "West", Rep = "Sorvino", Item = "Binder", Units = 7, UnitCost = 19.99, Total = 139.93},
                new {Id=27, OrderDate = DateTime.Parse("2016-03-24"), Region = "Central", Rep = "Jardine", Item = "Pen Set", Units = 50, UnitCost = 4.99, Total = 249.5},
                new {Id=28, OrderDate = DateTime.Parse("2016-04-10"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 66, UnitCost = 1.99, Total = 131.34},
                new {Id=29, OrderDate = DateTime.Parse("2016-04-27"), Region = "East", Rep = "Howard", Item = "Pen", Units = 96, UnitCost = 4.99, Total = 479.04},
                new {Id=30, OrderDate = DateTime.Parse("2016-05-14"), Region = "Central", Rep = "Gill", Item = "Pencil", Units = 53, UnitCost = 1.29, Total = 68.37},
                new {Id=31, OrderDate = DateTime.Parse("2016-05-31"), Region = "Central", Rep = "Gill", Item = "Binder", Units = 80, UnitCost = 8.99, Total = 719.2},
                new {Id=32, OrderDate = DateTime.Parse("2016-06-17"), Region = "Central", Rep = "Kivell", Item = "Desk", Units = 5, UnitCost = 125, Total = 625},
                new {Id=33, OrderDate = DateTime.Parse("2016-07-04"), Region = "East", Rep = "Jones", Item = "Pen Set", Units = 62, UnitCost = 4.99, Total = 309.38},
                new {Id=34, OrderDate = DateTime.Parse("2016-07-21"), Region = "Central", Rep = "Morgan", Item = "Pen Set", Units = 55, UnitCost = 12.49, Total = 686.95},
                new {Id=35, OrderDate = DateTime.Parse("2016-08-07"), Region = "Central", Rep = "Kivell", Item = "Pen Set", Units = 42, UnitCost = 23.95, Total = 1005.9},
                new {Id=36, OrderDate = DateTime.Parse("2016-08-24"), Region = "West", Rep = "Sorvino", Item = "Desk", Units = 3, UnitCost = 275, Total = 825},
                new {Id=37, OrderDate = DateTime.Parse("2016-09-10"), Region = "Central", Rep = "Gill", Item = "Pencil", Units = 7, UnitCost = 1.29, Total = 9.03},
                new {Id=38, OrderDate = DateTime.Parse("2016-09-27"), Region = "West", Rep = "Sorvino", Item = "Pen", Units = 76, UnitCost = 1.99, Total = 151.24},
                new {Id=39, OrderDate = DateTime.Parse("2016-10-14"), Region = "West", Rep = "Thompson", Item = "Binder", Units = 57, UnitCost = 19.99, Total = 1139.43},
                new {Id=40, OrderDate = DateTime.Parse("2016-10-31"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 14, UnitCost = 1.29, Total = 18.06},
                new {Id=41, OrderDate = DateTime.Parse("2016-11-17"), Region = "Central", Rep = "Jardine", Item = "Binder", Units = 11, UnitCost = 4.99, Total = 54.89},
                new {Id=42, OrderDate = DateTime.Parse("2016-12-04"), Region = "Central", Rep = "Jardine", Item = "Binder", Units = 94, UnitCost = 19.99, Total = 1879.06},
                new {Id=43, OrderDate = DateTime.Parse("2016-12-21"), Region = "Central", Rep = "Andrews", Item = "Binder", Units = 28, UnitCost = 4.99, Total = 139.72}
            };
            ws.FirstCell()
                .CellBelow()
                .CellRight()
                .InsertTable(data);

            return ws;
        }

        #endregion Setup and teardown

        [Test]
        public void Column()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            wb.AddWorksheet("Other");

            // If no argument, function uses the address of the cell that contains the formula
            Assert.AreEqual(4, ws.Cell("D1").SetFormulaA1("COLUMN()").Value);

            // With a reference, it returns the column number
            Assert.AreEqual(26, ws.Cell("A1").SetFormulaA1("COLUMN(Z14)").Value);

            // If a single column is used, return the column number 
            Assert.AreEqual(3, ws.Cell("A2").SetFormulaA1("COLUMN(C:C)").Value);

            // Return a horizontal array for multiple columns. Use SUM to verify content of an array since ROWS/COLUMNS don't work yet.
            Assert.AreEqual(3 + 4, ws.Cell("A3").SetFormulaA1("SUM(COLUMN(C:D))").Value);
            Assert.AreEqual(5 + 6 + 7, ws.Cell("A3").SetFormulaA1("SUM(COLUMN(E1:G10))").Value);

            // Not contiguous range (multiple areas) returns #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A4").SetFormulaA1("COLUMN((D5:G10,I8:K12))").Value);

            // Invalid references return #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A5").SetFormulaA1("COLUMN(NonExistent!F10)").Value);

            // Return column number even for different worksheet
            Assert.AreEqual(5, ws.Cell("A6").SetFormulaA1("COLUMN(Other!E7)").Value);

            // Unexpected types return error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A8").SetFormulaA1("COLUMN(TRUE)").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A7").SetFormulaA1("COLUMN(5)").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A8").SetFormulaA1("COLUMN(\"C5\")").Value);
            Assert.AreEqual(XLError.DivisionByZero, ws.Cell("A9").SetFormulaA1("COLUMN(#DIV/0!)").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A10").SetFormulaA1("COLUMN(\"C5\")").Value);
        }

        [Test]
        public void Columns_Blank_ReturnsValueError()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("COLUMNS(IF(TRUE,,))"));
        }

        [TestCase("0")]
        [TestCase("1")]
        [TestCase("99")]
        [TestCase("-10")]
        [TestCase("TRUE")]
        [TestCase("FALSE")]
        [TestCase("\"\"")]
        [TestCase("\"A\"")]
        [TestCase("\"Hello World\"")]
        public void Columns_ScalarValues_ReturnsOne(string value)
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr($"COLUMNS({value})"));
        }

        [Test]
        public void Columns_Error_ReturnsError()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("COLUMNS(#DIV/0!)"));
        }

        [TestCase("{1}", 1)]
        [TestCase("{1;2;3}", 1)]
        [TestCase("{1,2,3,4;5,6,7,8}", 4)]
        [TestCase("{TRUE,\"Z\";#DIV/0!,4}", 2)]
        public void Columns_Arrays_ReturnsNumberOfColumns(string array, int expectedColumnCount)
        {
            Assert.AreEqual(expectedColumnCount, XLWorkbook.EvaluateExpr($"COLUMNS({array})"));
        }

        [TestCase("A1", 1)]
        [TestCase("A1:A6", 1)]
        [TestCase("B2:D6", 3)]
        [TestCase("E7:AA14", 23)]
        public void Columns_References_ReturnsNumberOfColumns(string range, int expectedColumnCount)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            Assert.AreEqual(expectedColumnCount, sheet.Evaluate($"COLUMNS({range})"));
        }

        [Test]
        public void Columns_NonContiguousReferences_ReturnsReferenceError()
        {
            // Spec says #NULL!, but Excel says #REF!
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("COLUMNS((A1,C3))"));
        }

        [Test]
        public void Hlookup()
        {
            // Range lookup false
            var value = ws.Evaluate(@"=HLOOKUP(""Total"",Data!$B$2:$I$71,4,FALSE)");
            Assert.AreEqual(179.64, value);
        }

        [Test]
        public void Hyperlink()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            var cell = sheet.Cell("B3");
            cell.FormulaA1 = "HYPERLINK(\"http://github.com/ClosedXML/ClosedXML\")";
            Assert.AreEqual("http://github.com/ClosedXML/ClosedXML", cell.Value);
            Assert.False(cell.HasHyperlink);

            cell = sheet.Cell("B4");
            cell.FormulaA1 = "HYPERLINK(\"mailto:jsmith@github.com\", \"jsmith@github.com\")";
            Assert.AreEqual("jsmith@github.com", cell.Value);
            Assert.False(cell.HasHyperlink);

            cell = sheet.Cell("B5");
            cell.FormulaA1 = "HYPERLINK(\"[Test.xlsx]Sheet1!A5\", \"Cell A5\")";
            Assert.AreEqual("Cell A5", cell.Value);
            Assert.False(cell.HasHyperlink);
        }

        [Test]
        public void Index()
        {
            Assert.AreEqual("Kivell", ws.Evaluate(@"=INDEX(B2:J12, 3, 4)"));

            // We don't support optional parameter fully here yet.
            // Supposedly, if you omit e.g. the row number, then ROW() of the calling cell should be assumed
            // Assert.AreEqual("Gill", ws.Evaluate(@"=INDEX(B2:J12, , 4)"));

            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:I2, 4)"));

            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:I2, B6)"));

            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4)"));
            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4, 1)"));
            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4, 0)"));
            Assert.AreEqual(3, ws.Evaluate(@"=INDEX(B2:B20, 4, )"));

            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:J2, 1, 4)"));
            Assert.AreEqual("Rep", ws.Evaluate(@"=INDEX(B2:J2, , 4)"));
        }

        [Test]
        public void Index_WorksWithArrays()
        {
            Assert.AreEqual(1, ws.Evaluate(@"=INDEX({0,1,2}, 2)"));
        }

        [Test]
        public void Index_Exceptions()
        {
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:I10, 20, 1)"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:I10, 1, 10)"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:I2, 10)"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:I2, 4, 1)"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:I2, 4, )"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:B10, 20)"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:B10, 20, )"));
            Assert.AreEqual(XLError.CellReference, ws.Evaluate(@"INDEX(B2:B10, , 4)"));
        }

        [Test]
        public void Match()
        {
            Object value;
            value = ws.Evaluate(@"=MATCH(""Rep"", B2:I2, 0)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(E2, B2:I2, 0)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(""Rep"", A2:Z2, 0)");
            Assert.AreEqual(5, value);

            value = ws.Evaluate(@"=MATCH(""REP"", B2:I2, 0)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(95, B3:I3, 0)");
            Assert.AreEqual(6, value);

            value = ws.Evaluate(@"=MATCH(DATE(2015,1,6), B3:I3, 0)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(1.99, 3:3, 0)");
            Assert.AreEqual(8, value);

            value = ws.Evaluate(@"=MATCH(43, B:B, 0)");
            Assert.AreEqual(45, value);

            value = ws.Evaluate(@"=MATCH(""cENtraL"", D3:D45, 0)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(4.99, H:H, 0)");
            Assert.AreEqual(5, value);

            value = ws.Evaluate(@"=MATCH(""Rapture"", B2:I2, 1)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(22.5, B3:B45, 1)");
            Assert.AreEqual(22, value);

            value = ws.Evaluate(@"=MATCH(""Rep"", B2:I2)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(""Rep"", B2:I2, 1)");
            Assert.AreEqual(4, value);

            value = ws.Evaluate(@"=MATCH(40, G3:G6, -1)");
            Assert.AreEqual(2, value);

            value = ws.Evaluate(@"=MATCH(1, {0,1,2}, 0)");
            Assert.AreEqual(2, value);
        }

        [Test]
        public void Match_WorksWithArrays()
        {
            Object value;
            value = ws.Evaluate(@"=MATCH(1, {0,1,2}, 0)");
            Assert.AreEqual(2, value);
        }

        [Test]
        public void Match_Exceptions()
        {
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=MATCH(""Rep"", B2:I5)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=MATCH(""Dummy"", B2:I2, 0)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=MATCH(4.5,B3:B45,-1)"));
        }

        [Test]
        public void Row()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Data");
            wb.AddWorksheet("Other");

            // If no argument, function uses the address of the cell that contains the formula
            Assert.AreEqual(60, ws.Cell("M60").SetFormulaA1("ROW()").Value);

            // With a reference, it returns the row number
            Assert.AreEqual(12, ws.Cell("A1").SetFormulaA1("ROW(C12)").Value);

            // If a full row reference to a single row is used, return the row number 
            Assert.AreEqual(40, ws.Cell("A2").SetFormulaA1("ROW(40:40)").Value);

            // Return a vertical array for multiple rows. Use SUM to verify content of an array since ROWS/COLUMNS don't work yet.
            Assert.AreEqual(4 + 5 + 6 + 7, ws.Cell("A3").SetFormulaA1("SUM(ROW(4:7))").Value);
            Assert.AreEqual(2 + 3 + 4, ws.Cell("A4").SetFormulaA1("SUM(ROW(C2:Z4))").Value);

            // Not contiguous range (multiple areas) returns #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A5").SetFormulaA1("ROW((D5:G10,I8:K12))").Value);

            // Invalid references return #REF!
            Assert.AreEqual(XLError.CellReference, ws.Cell("A6").SetFormulaA1("ROW(NonExistent!F10)").Value);

            // Return row number even for different worksheet
            Assert.AreEqual(14, ws.Cell("A7").SetFormulaA1("ROW(Other!E14)").Value);

            // Unexpected types return error
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A8").SetFormulaA1("ROW(IF(TRUE,TRUE))").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A9").SetFormulaA1("ROW(IF(TRUE,5))").Value);
            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A10").SetFormulaA1("ROW(IF(TRUE,\"G15\"))").Value);
            Assert.AreEqual(XLError.DivisionByZero, ws.Cell("A11").SetFormulaA1("ROW(#DIV/0!)").Value);
        }

        [Test]
        public void Rows_Blank_ReturnsValueError()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("ROWS(IF(TRUE,,))"));
        }

        [TestCase("0")]
        [TestCase("1")]
        [TestCase("99")]
        [TestCase("-10")]
        [TestCase("TRUE")]
        [TestCase("FALSE")]
        [TestCase("\"\"")]
        [TestCase("\"A\"")]
        [TestCase("\"Hello World\"")]
        public void Rows_ScalarValues_ReturnsOne(string value)
        {
            Assert.AreEqual(1, XLWorkbook.EvaluateExpr($"ROWS({value})"));
        }

        [Test]
        public void Rows_Error_ReturnsError()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("ROWS(#DIV/0!)"));
        }

        [TestCase("{1}", 1)]
        [TestCase("{1;2;3}", 3)]
        [TestCase("{1,2,3,4;5,6,7,8;9,10,11,12}", 3)]
        [TestCase("{TRUE;#DIV/0!}", 2)]
        public void Rows_Arrays_ReturnsNumberOfRows(string array, int expectedColumnCount)
        {
            Assert.AreEqual(expectedColumnCount, XLWorkbook.EvaluateExpr($"ROWS({array})"));
        }

        [TestCase("C3", 1)]
        [TestCase("B3:E12", 10)]
        [TestCase("AA21:AC400", 380)]
        public void Rows_References_ReturnsNumberOfColumns(string range, int expectedColumnCount)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            Assert.AreEqual(expectedColumnCount, sheet.Evaluate($"ROWS({range})"));
        }

        [Test]
        public void Rows_NonContiguousReferences_ReturnsReferenceError()
        {
            // Spec says #NULL!, but Excel says #REF!
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("ROWS((A1,C3))"));
        }

        [Test]
        public void Vlookup()
        {
            // Range lookup false
            var value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,3,FALSE)");
            Assert.AreEqual("Central", value);

            value = ws.Evaluate("=VLOOKUP(DATE(2015,5,22),Data!C:I,7,FALSE)");
            Assert.AreEqual(63.68, value);

            value = ws.Evaluate(@"=VLOOKUP(""Central"",Data!D:E,2,FALSE)");
            Assert.AreEqual("Kivell", value);

            // Case insensitive lookup
            value = ws.Evaluate(@"=VLOOKUP(""central"",Data!D:E,2,FALSE)");
            Assert.AreEqual("Kivell", value);

            // Range lookup true
            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8,TRUE)");
            Assert.AreEqual(179.64, value);

            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8)");
            Assert.AreEqual(179.64, value);

            value = ws.Evaluate("=VLOOKUP(3,Data!$B$2:$I$71,8,)");
            Assert.AreEqual(179.64, value);

            value = ws.Evaluate("=VLOOKUP(14.5,Data!$B$2:$I$71,8,TRUE)");
            Assert.AreEqual(174.65, value);

            value = ws.Evaluate("=VLOOKUP(50,Data!$B$2:$I$71,8,TRUE)");
            Assert.AreEqual(139.72, value);
        }

        [Test]
        public void Vlookup_ElementNotFound_ReturnsNotAvailableError()
        {
            // Value not present in the range for exact search
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=VLOOKUP("""",Data!$B$2:$I$71,3,FALSE)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=VLOOKUP(50,Data!$B$2:$I$71,3,FALSE)"));

            // Value in approximate search that is lower than first element
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=VLOOKUP(-1,Data!$B$2:$I$71,2,TRUE)"));
        }

        [Test]
        public void Vlookup_UnexpectedArguments()
        {
            // Lookup value can't be an error
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("=VLOOKUP(#DIV/0!,B2:I71,1)"));

            // Text value can't be over 255 chars
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate($"=VLOOKUP(\"{new string('A', 256)}\",B2:I71,1)"));

            // Range can only be array or a reference. If other type, it returns the error #N/A
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=VLOOKUP(1,1,1)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=VLOOKUP(1,TRUE,1)"));

            // If range is a non-contiguous range, #N/A
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=VLOOKUP(1,(B2:I5,B6:I10),1)"));

            // The column index must be at most the same as width of the range. It is 9 here, but range is 8 cell wide.
            Assert.AreEqual(XLError.CellReference, ws.Evaluate("=VLOOKUP(20,B2:I71,9,FALSE)"));
            // The column index must be at least 1. It is 0 here.
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=VLOOKUP(20,B2:I71,0,FALSE)"));
        }

        [Test]
        public void Vlookup_ColumnIndexParameter_UsesValueSemantic()
        {
            // If column index is not a whole number, it is truncated, so here 1.9 is truncated to 1
            Assert.AreEqual(14.0, ws.Evaluate("=VLOOKUP(14,B2:I71,1.9)"));

            // Column index is evaluated using a VALUE semantic
            Assert.AreEqual(@"Jardine", ws.Evaluate("=VLOOKUP(3,B2:I71,\"2 5/2\")"));
        }

        [TestCase("\"TRUE\"")]
        [TestCase("1")]
        [TestCase("TRUE")]
        public void Vlookup_FlagParameter_CoercedToBoolean(string flagValue)
        {
            Assert.AreEqual(5.0, ws.Evaluate($"VLOOKUP(5,B2:I71,1,{flagValue})"));
        }

        [Test]
        public void Vlookup_BlankLookupValue_BehavesAsZero()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").InsertData(Enumerable.Range(-5, 10).Select(x => new object[] { x, $"Row with value {x}" }));

            var actual = worksheet.Evaluate("VLOOKUP(IF(TRUE,,),A1:B10,2)");

            Assert.AreEqual("Row with value 0", actual);
        }

        [Test]
        public void Vlookup_ApproximateSearch_OmitsValuesWithDifferentType()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = "0";
            worksheet.Cell("A2").Value = "1";
            worksheet.Cell("A3").Value = 1;
            worksheet.Cell("A4").Value = "0";
            worksheet.Cell("A5").Value = "text";
            worksheet.Cell("A6").Value = Blank.Value;
            worksheet.Cell("A7").Value = 2;
            worksheet.Cell("B1").InsertData(Enumerable.Range(1, 7).Select(x => $"Row {x}"));

            var actual = worksheet.Evaluate("VLOOKUP(1.9,A1:B7,2,TRUE)");
            Assert.AreEqual("Row 3", actual);
        }

        [Test]
        public void Vlookup_OnlyCellsWithDifferentType_ReturnsNotAvailable()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            Assert.AreEqual(XLError.NoValueAvailable, worksheet.Evaluate("VLOOKUP(1,A1,1,TRUE)"));
        }

        [Test]
        public void Vlookup_OnlyOneValueSurroundedByIgnoredTypes()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A3").Value = 5;

            Assert.AreEqual(5, worksheet.Evaluate("VLOOKUP(6,A1:A5,1,TRUE)"));
        }

        [Test]
        public void Vlookup_ResultAtTheHighestCellWithTrailingDifferentTypeAtTheEnd()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = 1;
            worksheet.Cell("A2").Value = 2;
            worksheet.Cell("A3").Value = 3;
            worksheet.Cell("A4").Value = Blank.Value;

            Assert.AreEqual(3, worksheet.Evaluate("VLOOKUP(3,A1:A4,1,TRUE)"));
        }

        [Test]
        public void Vlookup_ApproximateSearch_ReturnsLastRowForMultipleEqualValues()
        {
            var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            sheet.Cell("A1").Value = 1;
            sheet.Cell("A2").Value = 3;
            sheet.Cell("A3").Value = 3;
            sheet.Cell("A4").Value = 3;
            sheet.Cell("A5").Value = 3;
            sheet.Cell("A6").Value = 3;
            sheet.Cell("A7").Value = 3;
            sheet.Cell("A8").Value = 9;
            sheet.Cell("B1").InsertData(Enumerable.Range(1, 8));

            // If there is a section of values with same value, return the value at the highest row
            var actual = sheet.Evaluate("VLOOKUP(3, A1:B8, 2, TRUE)");
            Assert.AreEqual(7, actual);

            // If the last value is in the highest row, just return value outright
            actual = sheet.Evaluate("VLOOKUP(3, A2:B7, 2, TRUE)");
            Assert.AreEqual(7, actual);
        }

        [Test]
        public void Vlookup_CanSearchArrays()
        {
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("VLOOKUP(4, {1,2; 3,2; 5,3; 7,4}, 2)"));
        }

        [Test]
        public void Xlookup()
        {
            // Range lookup with only exact match (the default)
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(3,Data!$B$2:$B$71,Data!$D$2:$D$71,,0)"));
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(B5,Data!$B$2:$B$71,Data!$D$2:$D$71,,0)"));
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(B5:B5,Data!$B$2:$B$71,Data!$D$2:$D$71,,0)"));

            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(3,Data!$B$2:$B$71,Data!$D$2:$D$71)"));
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(B5,Data!$B$2:$B$71,Data!$D$2:$D$71)"));
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(B5:B5,Data!$B$2:$B$71,Data!$D$2:$D$71)"));

            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(3,Data!$B$2:$B$71,Data!$D$2:$D$71,,)"));
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(B5,Data!$B$2:$B$71,Data!$D$2:$D$71,,)"));
            Assert.AreEqual("Central", ws.Evaluate("=XLOOKUP(B5:B5,Data!$B$2:$B$71,Data!$D$2:$D$71,,)"));

            Assert.AreEqual(63.68, ws.Evaluate("=XLOOKUP(DATE(2015,5,22),Data!C:C,Data!I:I,,0)"));
            Assert.AreEqual(63.68, ws.Evaluate("=XLOOKUP(C11,Data!C:C,Data!I:I,,0)"));
            Assert.AreEqual(63.68, ws.Evaluate("=XLOOKUP(C11:C11,Data!C:C,Data!I:I,,0)"));

            Assert.AreEqual("Kivell", ws.Evaluate(@"=XLOOKUP(""Central"",Data!D:D,Data!E:E,,0)"));

            // Case insensitive lookup
            Assert.AreEqual("Kivell", ws.Evaluate(@"=XLOOKUP(""central"",Data!D:D,Data!E:E,,0)"));

            // Range lookup with trying for exact match but returning the next smaller item if exact match not found
            Assert.AreEqual(179.64, ws.Evaluate("=XLOOKUP(3,Data!$B$2:$B$71,Data!$I$2:$I$71,,-1)"));

            Assert.AreEqual(174.65, ws.Evaluate("=XLOOKUP(14.5,Data!$B$2:$B$71,Data!$I$2:$I$71,,-1)"));

            Assert.AreEqual(174.65, ws.Evaluate("=XLOOKUP(14.6,Data!$B$2:$B$71,Data!$I$2:$I$71,,-1)"));

            Assert.AreEqual(139.72, ws.Evaluate("=XLOOKUP(50,Data!$B$2:$B$71,Data!$I$2:$I$71,,-1)"));

            // Range lookup with trying for exact match but returning the next larger item if exact match not found
            Assert.AreEqual(179.64, ws.Evaluate("=XLOOKUP(3,Data!$B$2:$B$71,Data!$I$2:$I$71,,1)"));

            Assert.AreEqual(189.05, ws.Evaluate("=XLOOKUP(0,Data!$B$2:$B$71,Data!$I$2:$I$71,,1)"));

            Assert.AreEqual(250, ws.Evaluate("=XLOOKUP(14.4,Data!$B$2:$B$71,Data!$I$2:$I$71,,1)"));

            Assert.AreEqual(250, ws.Evaluate("=XLOOKUP(14.5,Data!$B$2:$B$71,Data!$I$2:$I$71,,1)"));
        }

        [Test]
        public void Xlookup_ReturnsReference()
        {
            using var wb = new XLWorkbook();
            var worksheet = (XLWorksheet)wb.AddWorksheet();
            worksheet.Cell("A1").Value = 1;
            worksheet.Cell("A2").Value = 2;
            worksheet.Cell("A3").Value = 3;
            worksheet.Cell("B1").Value = "Value1";
            worksheet.Cell("B2").Value = "Value2";
            worksheet.Cell("B3").Value = "Value3";

            // When evaluated, XLOOKUP should return the real value
            var actual = worksheet.Evaluate("XLOOKUP(2,A1:A3,B1:B3)");
            Assert.AreEqual("Value2", actual);

            // But when used in combination with other calculations, XLOOKUP needs to return a reference.
            // This is important so that calcs like =SUM(XLOOKUP(1,A1:A3,B1:B3)) work when XLOOKUP returns text.
            // IMPORTANT: This is different to how VLOOKUP and HLOOKUP work in Excel, where something like
            // =SUM(VLOOKUP(1,A1:B3,2)) returns #VALUE! if VLOOKUP returns text.
            var calcEngine = new XLCalcEngine(CultureInfo.CurrentCulture);
            var ctx = new CalcContext(calcEngine, CultureInfo.CurrentCulture, wb, worksheet, null);
            var value = calcEngine.EvaluateFormula("XLOOKUP(1,A1:A3,B1:B3)", ctx);
            Assert.That(value.IsReference);
            if (value.TryPickReference(out var reference, out _))
            {
                Assert.That(reference.IsSingleCell);
                Assert.AreEqual("B1:B1", reference.Areas.Single().ToString());
            }

            // Test XLOOKUP directly inside SUM case for good measure
            actual = worksheet.Evaluate("SUM(XLOOKUP(1,A1:A3,B1:B3))");
            Assert.AreEqual(0, actual);
        }

        [Test]
        public void Xlookup_ElementNotFound_ReturnsNotAvailableError()
        {
            // Value not present in the range for exact search
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=XLOOKUP("""",Data!$B$2:$B$71,Data!$D$2:$D$71)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=XLOOKUP(50,Data!$B$2:$B$71,Data!$D$2:$D$71)"));

            // Value in approximate search that is lower than first element
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"=XLOOKUP(-1,Data!$B$2:$B$71,Data!$C$2:$C$71,,-1)"));
        }

        [Test]
        public void Xlookup_ElementNotFound_ReturnsIfNotFoundValue()
        {
            // Value not present in the range for exact search
            Assert.AreEqual("Not Found", ws.Evaluate(@"=XLOOKUP("""",Data!$B$2:$B$71,Data!$D$2:$D$71,""Not Found"")"));
            Assert.AreEqual("Not Found", ws.Evaluate(@"=XLOOKUP(50,Data!$B$2:$B$71,Data!$D$2:$D$71,""Not Found"")"));

            // Value in approximate search that is lower than first element
            Assert.AreEqual("Not Found", ws.Evaluate(@"=XLOOKUP(-1,Data!$B$2:$B$71,Data!$C$2:$C$71,""Not Found"",-1)"));
        }

        [Test]
        public void Xlookup_UnexpectedArguments()
        {
            // Lookup value can't be an error
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("=XLOOKUP(#DIV/0!,B2:B71,C2:C71)"));

            // Text value can't be over 255 chars
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate($"=XLOOKUP(\"{new string('A', 256)}\",B2:B71,C2:C71)"));

            // Ranges are only allowed to be 1-dimensional
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(43,B2:C71,I2:I70)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(43,B2:B71,H2:I70)"));

            // The rules for what is allowed for the 2 range values is quite complicated
            Assert.AreEqual(1, ws.Evaluate("=XLOOKUP(1,1,1)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=XLOOKUP(1,2,1)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(1,B2:B71,1)"));
            // Assert.AreEqual("OrderDate", ws.Evaluate("=XLOOKUP(1,1,C2:C71)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(1,B2:B71,TRUE)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=XLOOKUP(1,TRUE,C2:C71)"));

            // If range is a non-contiguous range, #N/A
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=XLOOKUP(1,(B2:B5,B6:B10),C2:C71)"));
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("=XLOOKUP(1,B2:B71,(C2:C5,C6:C10))"));

            // The lengths of both ranges must be exactly the same, otherwise #VALUE!
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(43,B2:B71,I2:I5)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(43,B2:B71,I2:I70)"));
            Assert.AreEqual(XLError.IncompatibleValue, ws.Evaluate("=XLOOKUP(43,B2:B71,I2:I72)"));
        }

        [Test]
        public void Xlookup_BlankLookupValue_BehavesAsBlank()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").InsertData(Enumerable.Range(-5, 10).Select(x => new object[] { x, $"Row with value {x}" }));
            worksheet.Cell("B11").Value = "Row with blank value";

            var actual = worksheet.Evaluate("XLOOKUP(IF(TRUE,,),A1:A11,B1:B11)");

            Assert.AreEqual("Row with blank value", actual);
        }

        [Test]
        public void Xlookup_ApproximateSearch_OmitsValuesWithDifferentType()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = "0";
            worksheet.Cell("A2").Value = "1";
            worksheet.Cell("A3").Value = 1;
            worksheet.Cell("A4").Value = "0";
            worksheet.Cell("A5").Value = "text";
            worksheet.Cell("A6").Value = Blank.Value;
            worksheet.Cell("A7").Value = 2;
            worksheet.Cell("B1").InsertData(Enumerable.Range(1, 7).Select(x => $"Row {x}"));

            var actual = worksheet.Evaluate("XLOOKUP(1.9,A1:A7,B1:B7,,-1)");
            Assert.AreEqual("Row 3", actual);
        }

        [Test]
        public void Xlookup_OnlyCellsWithDifferentType_ReturnsNotAvailable()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            Assert.AreEqual(XLError.NoValueAvailable, worksheet.Evaluate("XLOOKUP(1,A1,A1,,-1)"));
        }

        [Test]
        public void Xlookup_OnlyOneValueSurroundedByIgnoredTypes()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A3").Value = 5;

            Assert.AreEqual(5, worksheet.Evaluate("XLOOKUP(6,A1:A5,A1:A5,,-1)"));
        }

        [Test]
        public void Xlookup_ResultAtTheHighestCellWithTrailingDifferentTypeAtTheEnd()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = 1;
            worksheet.Cell("A2").Value = 2;
            worksheet.Cell("A3").Value = 3;
            worksheet.Cell("A4").Value = Blank.Value;

            Assert.AreEqual(3, worksheet.Evaluate("XLOOKUP(3,A1:A4,A1:A4,,-1)"));
        }

        [Test]
        public void Xlookup_ApproximateSearch_ReturnsResultFromMultipleEqualValuesBasedOnSearchMode()
        {
            using var wb = new XLWorkbook();
            var worksheet = wb.AddWorksheet();
            worksheet.Cell("A1").Value = 1;
            worksheet.Cell("A2").Value = 3;
            worksheet.Cell("A3").Value = 3;
            worksheet.Cell("A4").Value = 3;
            worksheet.Cell("A5").Value = 3;
            worksheet.Cell("A6").Value = 3;
            worksheet.Cell("A7").Value = 3;
            worksheet.Cell("A8").Value = 9;
            worksheet.Cell("B1").InsertData(Enumerable.Range(1, 8));

            // If there is a section of values with same value, return the first value we find
            Assert.AreEqual(2, worksheet.Evaluate("XLOOKUP(3,A1:A8,B1:B8,,-1)"));
            Assert.AreEqual(2, worksheet.Evaluate("XLOOKUP(3,A1:A8,B1:B8,,-1,1)"));

            // If there is a section of values with same value, and we're searching from the bottom, then return the last value in the list
            Assert.AreEqual(7, worksheet.Evaluate("XLOOKUP(3,A1:A8,B1:B8,,-1,-1)"));
        }

        [Test]
        public void Xlookup_CanSearchArrays()
        {
            Assert.AreEqual(2, XLWorkbook.EvaluateExpr("XLOOKUP(4, {1; 3; 5; 7}, {2; 2; 3; 4}, , -1)"));
        }
    }
}
