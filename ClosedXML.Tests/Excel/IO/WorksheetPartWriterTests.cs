using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class WorksheetPartWriterTests
{
    private const string RevisionNs = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

    [Test]
    [Issue("74")]
    public void Saved_sheet_has_only_the_rows_of_the_workbook()
    {
        // The writer doesn't load the rows of the old part, and the sheet data is the last
        // element of the worksheet. The worksheet keeps its attributes and ignored namespaces.
        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var worksheet = new Worksheet(
                new SheetData(
                    new Row(new Cell { CellReference = "A1", CellValue = new CellValue(1) }) { RowIndex = 1 },
                    new Row(new Cell { CellReference = "A2", CellValue = new CellValue(2) }) { RowIndex = 2 }));
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("xr", RevisionNs);
            worksheet.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "xr" };
            worksheet.SetAttribute(new OpenXmlAttribute("xr", "uid", RevisionNs, "{00000000-0000-0000-0000-000000000074}"));
            worksheetPart.Worksheet = worksheet;
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" }));
        }

        using var savedMs = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet(1);
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Clear();
            ws.Cell("B3").Value = 30;
            wb.SaveAs(savedMs);
        }

        savedMs.Position = 0;
        using (var document = SpreadsheetDocument.Open(savedMs, false))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var sheetData = worksheet.Elements<SheetData>().Single();
            var cells = sheetData.Descendants<Cell>().Select(c => $"{c.CellReference!.Value}={c.CellValue!.Text}");
            CollectionAssert.AreEqual(new[] { "A1=10", "B3=30" }, cells);
            Assert.AreEqual("{00000000-0000-0000-0000-000000000074}", worksheet.GetAttribute("uid", RevisionNs).Value);
            Assert.AreEqual("xr x14ac", worksheet.MCAttributes!.Ignorable!.Value);
        }
    }

    [Test]
    [Issue("74")]
    public void Table_total_label_is_saved_with_other_cells()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertTable(new[] { new { Name = "a", Value = 1 }, new { Name = "b", Value = 2 } });
            var table = ws.Table(0);
            table.ShowTotalsRow = true;
            table.Field("Name").TotalsRowLabel = "Total";
            table.Field("Value").TotalsRowFunction = XLTotalsRowFunction.Sum;
            ws.Cell("D5").Value = 7;
            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet(1);
            Assert.AreEqual("Total", ws.Cell("A4").Value);
            Assert.AreEqual("SUBTOTAL(109,[Value])", ws.Cell("B4").FormulaA1);
            Assert.AreEqual(7, ws.Cell("D5").Value);
        }
    }

    [Test]
    [Issue("74")]
    public void Table_total_cell_with_white_space_is_saved_as_a_value()
    {
        // A label of white space is not a label, so the writer writes the cell value.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertTable(new[] { new { Name = "a", Value = 1 }, new { Name = "b", Value = 2 } });
            ws.Table(0).ShowTotalsRow = true;
            ws.Cell("A4").Value = " ";
            ws.Cell("B5").Value = 3;
            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet(1);
            Assert.AreEqual(" ", ws.Cell("A4").Value);
            Assert.AreEqual(3d, ws.Cell("B5").GetDouble());
        }
    }
}
