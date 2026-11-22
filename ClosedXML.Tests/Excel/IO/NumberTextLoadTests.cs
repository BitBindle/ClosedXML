using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.IO;

/// <summary>
/// The load reads each number text of a file as the nearest double. The text <c>9007199254740991.5</c>
/// is at the midpoint between two doubles. The nearest double with the even mantissa is <c>0x4340000000000000</c>.
/// </summary>
[TestFixture]
internal class NumberTextLoadTests
{
    private const string NumberText = "9007199254740991.5";

    private const ulong NearestBits = 0x4340000000000000UL;

    [Test]
    [Issue("46")]
    public void Load_reads_nearest_double_of_pivot_cache_numbers()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            var data = ws.Cell("A1").InsertData(new object[]
            {
                ("Name", "Value"),
                ("Cake", 5),
            });
            var pt = ws.PivotTables.Add("pt", ws.Cell("D1"), data);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Value");
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var cachePart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            var sharedItems = cachePart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ElementAt(1).SharedItems!;
            sharedItems.MinValue = new DoubleValue { InnerText = NumberText };
            sharedItems.MaxValue = new DoubleValue { InnerText = NumberText };
            sharedItems.RemoveAllChildren();
            sharedItems.AppendChild(new NumberItem { Val = new DoubleValue { InnerText = NumberText } });
            sharedItems.Count = 1;

            var record = cachePart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Elements<PivotCacheRecord>().Single();
            record.Elements<NumberItem>().Single().Val = new DoubleValue { InnerText = NumberText };
        }

        using (var wb = new XLWorkbook(ms))
        {
            var cache = (XLPivotCache)wb.Worksheet(1).PivotTable("pt").PivotCache;
            var fieldValues = cache.GetFieldValues(1);
            Assert.AreEqual(NearestBits, GetBits(fieldValues.GetCellValues().Single().GetNumber()));
            Assert.AreEqual(NearestBits, GetBits(fieldValues.SharedItems[0].GetNumber()));
            Assert.AreEqual(NearestBits, GetBits(fieldValues.Stats.MinValue!.Value));
            Assert.AreEqual(NearestBits, GetBits(fieldValues.Stats.MaxValue!.Value));
        }
    }

    [Test]
    [Issue("46")]
    public void Load_reads_infinite_pivot_cache_limits()
    {
        // The type xsd:double has the texts "INF" and "-INF" for infinity.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            var data = ws.Cell("A1").InsertData(new object[]
            {
                ("Name", "Value"),
                ("Cake", 5),
            });
            var pt = ws.PivotTables.Add("pt", ws.Cell("D1"), data);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Value");
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var cachePart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            var sharedItems = cachePart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ElementAt(1).SharedItems!;
            sharedItems.MinValue = new DoubleValue { InnerText = "-INF" };
            sharedItems.MaxValue = new DoubleValue { InnerText = "INF" };
        }

        using (var wb = new XLWorkbook(ms))
        {
            var cache = (XLPivotCache)wb.Worksheet(1).PivotTable("pt").PivotCache;
            var stats = cache.GetFieldValues(1).Stats;
            Assert.AreEqual(double.NegativeInfinity, stats.MinValue);
            Assert.AreEqual(double.PositiveInfinity, stats.MaxValue);
        }
    }

    [Test]
    [Issue("46")]
    public void Load_reads_nearest_double_of_custom_property()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet();
            wb.CustomProperties.Add("Number", 1.5);
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var property = document.CustomFilePropertiesPart!.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>().Single();
            property.VTDouble!.Text = NumberText;
        }

        using (var wb = new XLWorkbook(ms))
        {
            var value = wb.CustomProperties.CustomProperty("Number").GetValue<double>();
            Assert.AreEqual(NearestBits, GetBits(value));
        }
    }

    private static ulong GetBits(double value)
    {
        return (ulong)BitConverter.DoubleToInt64Bits(value);
    }
}
