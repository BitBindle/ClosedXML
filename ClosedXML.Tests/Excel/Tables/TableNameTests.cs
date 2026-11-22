using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Tables
{
    [TestFixture]
    public class TableNameTests
    {
        [TestCase("My table")]
        [TestCase("My\ttable")]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Name_with_whitespace_is_rejected(string name)
        {
            // ECMA-376 Part 1, 18.5.1.2: the name of a table shall not have spaces.
            using var wb = new XLWorkbook();
            var table = CreateTable(wb.AddWorksheet());

            var ex = Assert.Throws<ArgumentException>(() => table.Name = name);
            StringAssert.Contains("whitespace", ex!.Message);
            Assert.AreEqual("Table1", table.Name);
        }

        [TestCase("Sales-2024")]
        [TestCase("Sales(1)")]
        [TestCase("A&B")]
        [TestCase(@"Sales\2024")]
        [Issue("56")]
        public void Name_with_character_that_Excel_rejects_is_rejected(string name)
        {
            // Excel accepts only letters, digits, periods and underscores after the first character.
            using var wb = new XLWorkbook();
            var table = CreateTable(wb.AddWorksheet());

            var ex = Assert.Throws<ArgumentException>(() => table.Name = name);
            StringAssert.Contains("not a letter, a digit, a period or an underscore", ex!.Message);
            Assert.AreEqual("Table1", table.Name);
        }

        [TestCase(@"\Sales")]
        [TestCase("_Sales.2024")]
        [TestCase("Ventes_Été")]
        [Issue("56")]
        public void Save_keeps_valid_table_name(string name)
        {
            // The first character can be a backslash. Letters of all languages are valid.
            using var wb = new XLWorkbook();
            CreateTable(wb.AddWorksheet()).Name = name;
            using var ms = new MemoryStream();
            wb.SaveAs(ms);

            AssertSavedTableName(ms, name);
        }

        [Test]
        [Issue("56")]
        public void Save_keeps_loaded_table_name_with_a_space()
        {
            // The formulas and the pivot sources of the file use the name from the file.
            using var ms = CreateFileWithTableName("My sales");
            using var savedMs = new MemoryStream();
            using (var wb = new XLWorkbook(ms))
            {
                wb.SaveAs(savedMs);
            }

            AssertSavedTableName(savedMs, "My sales");
        }

        [TestCase("A1")]
        [TestCase("xfd1048576")]
        [TestCase("May2019")]
        [TestCase("R1C2")]
        [TestCase("r3c2")]
        [TestCase("RC")]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Name_that_is_a_cell_reference_is_rejected(string name)
        {
            // A formula could not tell a structured reference from a cell reference.
            using var wb = new XLWorkbook();
            var table = CreateTable(wb.AddWorksheet());

            var ex = Assert.Throws<ArgumentException>(() => table.Name = name);
            StringAssert.Contains("cell reference", ex!.Message);
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Name_of_workbook_defined_name_is_rejected()
        {
            // The name of a table is unique among the defined names of the workbook.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            wb.DefinedNames.Add("Sales", "Sheet1!$A$1:$A$5");
            var table = CreateTable(ws);

            var ex = Assert.Throws<ArgumentException>(() => table.Name = "sales");
            StringAssert.Contains("defined name", ex!.Message);
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Name_of_worksheet_defined_name_is_rejected()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            var ws2 = wb.AddWorksheet();
            ws2.DefinedNames.Add("Sales", "Sheet2!$A$1:$A$5");
            var table = CreateTable(ws1);

            Assert.Throws<ArgumentException>(() => table.Name = "Sales");
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Rename_that_changes_only_the_case_is_allowed()
        {
            using var wb = new XLWorkbook();
            var table = CreateTable(wb.AddWorksheet());
            table.Name = "Sales";

            table.Name = "SALES";

            Assert.AreEqual("SALES", table.Name);
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Valid_name_is_accepted()
        {
            using var wb = new XLWorkbook();
            var table = CreateTable(wb.AddWorksheet());

            table.Name = "Sales_2019.Q1";

            Assert.AreEqual("Sales_2019.Q1", table.Name);
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Generated_name_skips_defined_names()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            wb.DefinedNames.Add("Table1", "Sheet1!$A$1");

            var table = CreateTable(ws);

            Assert.AreEqual("Table2", table.Name);
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Load_keeps_table_name_with_a_space()
        {
            // Excel does not write such a name, but other producers can. The load must not fail.
            using var ms = CreateFileWithTableName("My sales");
            using var loaded = new XLWorkbook(ms);
            Assert.AreEqual("My sales", loaded.Worksheets.Single().Tables.Single().Name);
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1951")]
        public void Copy_keeps_table_name_from_file()
        {
            // The copy takes the name of an existing table. The name rules apply only to a new name.
            using var ms = CreateFileWithTableName("My sales");
            using var wb = new XLWorkbook(ms);
            var source = wb.Worksheets.Single();
            var table = source.Tables.Single();

            var tableCopy = table.CopyTo(wb.AddWorksheet("Table copy"));
            var sheetCopy = source.CopyTo("Sheet copy");
            using var otherWorkbook = new XLWorkbook();
            var otherSheet = source.CopyTo(otherWorkbook);
            using var addWorkbook = new XLWorkbook();
            addWorkbook.AddWorksheet(source);

            Assert.AreEqual("My sales", tableCopy.Name);
            Assert.AreEqual("My sales", sheetCopy.Tables.Single().Name);
            Assert.AreEqual("My sales", otherSheet.Tables.Single().Name);
            Assert.AreEqual("My sales", addWorkbook.Worksheets.Single().Tables.Single().Name);
        }

        /// <summary>
        /// Create a file with one table that has the name <paramref name="tableName"/>, without the name rules of Excel.
        /// </summary>
        private static MemoryStream CreateFileWithTableName(string tableName)
        {
            var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                CreateTable(wb.AddWorksheet()).Name = "Sales";
                wb.SaveAs(ms);
            }

            using (var package = SpreadsheetDocument.Open(ms, true))
            {
                var table = package.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table;
                table.Name = tableName;
                table.DisplayName = tableName;
                table.Save();
            }

            ms.Position = 0;
            return ms;
        }

        private static void AssertSavedTableName(Stream stream, string expectedName)
        {
            using var package = SpreadsheetDocument.Open(stream, false);
            var table = package.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
            Assert.AreEqual(expectedName, table.Name!.Value);
            Assert.AreEqual(expectedName, table.DisplayName!.Value);
        }

        private static IXLTable CreateTable(IXLWorksheet ws)
        {
            ws.Cell("A1").Value = "Number";
            ws.Cell("A2").Value = 1;
            return ws.Range("A1:A2").CreateTable();
        }
    }
}
