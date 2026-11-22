using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class CalculationChainPartWriterTests
{
    private const string RevisionNs = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

    [Test]
    [Issue("74")]
    public void Chain_has_each_formula_cell_of_each_sheet()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws1 = wb.AddWorksheet();
            ws1.Cell("A1").FormulaA1 = "1+2";
            ws1.Cell("A2").Value = 5;
            ws1.Range("B2:C3").FormulaArrayA1 = "1";
            var ws2 = wb.AddWorksheet();
            ws2.Cell("XFD1048576").FormulaA1 = "3";
            wb.SaveAs(ms);
        }

        // The first cell of an array is in the chain two times, the first time with the array flag.
        CollectionAssert.AreEqual(
            new[] { "A1 1 False", "B2 1 True", "B2 1 False", "C2 1 False", "B3 1 False", "C3 1 False", "XFD1048576 2 False" },
            ReadChain(ms));
    }

    [Test]
    [Issue("74")]
    public void Chain_without_formulas_is_removed()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet().Cell("A1").FormulaA1 = "1";
            wb.SaveAs(ms);
        }

        using var savedMs = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet(1).Cell("A1").Value = 1;
            wb.SaveAs(savedMs);
        }

        savedMs.Position = 0;
        using var document = SpreadsheetDocument.Open(savedMs, false);
        Assert.IsNull(document.WorkbookPart!.CalculationChainPart);
    }

    [Test]
    [Issue("74")]
    public void Chain_after_a_sheet_delete_has_the_formulas_of_the_workbook()
    {
        // A sheet delete loads the DOM of the chain. The DOM must not replace the written chain.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet().Cell("A1").FormulaA1 = "1";
            wb.AddWorksheet().Cell("A1").FormulaA1 = "2";
            wb.SaveAs(ms);
        }

        using var savedMs = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet(1).Delete();
            wb.Worksheet(1).Cell("D4").FormulaA1 = "3";
            wb.SaveAs(savedMs);
        }

        CollectionAssert.AreEqual(new[] { "A1 2 False", "D4 2 False" }, ReadChain(savedMs));
    }

    [Test]
    [Issue("74")]
    public void Chain_keeps_the_extension_list()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet().Cell("A1").FormulaA1 = "1";
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var chain = document.WorkbookPart!.CalculationChainPart!.CalculationChain!;
            chain.AddNamespaceDeclaration("xr", RevisionNs);
            chain.Append(new ExtensionList(new Extension { Uri = "{00000000-0000-0000-0000-000000000074}" }));
        }

        using var savedMs = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet(1).Cell("B1").FormulaA1 = "2";
            wb.SaveAs(savedMs);
        }

        savedMs.Position = 0;
        using (var document = SpreadsheetDocument.Open(savedMs, false))
        {
            var chain = document.WorkbookPart!.CalculationChainPart!.CalculationChain!;
            var extensionList = (ExtensionList)chain.LastChild!;
            Assert.AreEqual("{00000000-0000-0000-0000-000000000074}", extensionList.Elements<Extension>().Single().Uri!.Value);

            // An extension can use a prefix of the root element in an attribute value.
            Assert.AreEqual(RevisionNs, chain.LookupNamespace("xr"));
        }

        CollectionAssert.AreEqual(new[] { "A1 1 False", "B1 1 False" }, ReadChain(savedMs));
    }

    [Test]
    [Issue("74")]
    public void Empty_chain_part_is_replaced()
    {
        // The SDK reads an empty part as a part without a root element.
        using var ms = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" }));
            workbookPart.AddNewPart<CalculationChainPart>();
        }

        using var savedMs = new MemoryStream();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet(1).Cell("A1").FormulaA1 = "1";
            wb.SaveAs(savedMs);
        }

        CollectionAssert.AreEqual(new[] { "A1 1 False" }, ReadChain(savedMs));
    }

    private static string[] ReadChain(MemoryStream ms)
    {
        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        return document.WorkbookPart!.CalculationChainPart!.CalculationChain!
            .Elements<CalculationCell>()
            .Select(c => $"{c.CellReference!.Value} {c.SheetId!.Value} {c.Array?.Value ?? false}")
            .ToArray();
    }
}
