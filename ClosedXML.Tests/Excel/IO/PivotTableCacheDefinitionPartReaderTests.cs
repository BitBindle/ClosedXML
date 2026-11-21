using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.IO;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class PivotTableCacheDefinitionPartReaderTests
{
    [Test]
    public void Cache_definition_part_without_root_element_throws()
    {
        // The pivotCacheDefinition element is required. ClosedXML can't load a part without it.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = "Name";
            data.Cell("A2").Value = "Value";
            var range = data.Range("A1:A2");
            wb.AddWorksheet("Pivot").PivotTables.Add("PivotTable", wb.Worksheet("Pivot").Cell("A1"), range);
            wb.SaveAs(ms);
        }

        using (var package = SpreadsheetDocument.Open(ms, true))
        {
            var cachePart = package.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            cachePart.FeedData(new MemoryStream());
        }

        ms.Position = 0;
        Assert.That(() => new XLWorkbook(ms), Throws
            .TypeOf<PartStructureException>()
            .With.Message.Contains("pivotCacheDefinition"));
    }
}
