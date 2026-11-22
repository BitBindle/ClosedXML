using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables;

/// <summary>
/// Tests of cache fields that have no values in the cache records: calculated fields and group fields.
/// </summary>
[TestFixture]
internal class XLPivotCacheFieldTests
{
    [Test]
    [Issue("45")]
    public void Calculated_field_is_loaded_as_cache_field()
    {
        using var ms = CreateFileWithCalculatedField();

        using var wb = new XLWorkbook(ms);
        var pt = wb.Worksheet(1).PivotTable("pt");

        Assert.AreEqual(new[] { "Name", "Value", "Double" }, pt.PivotCache.FieldNames);
        Assert.AreEqual(new[] { "Value", "Double" }, pt.Values.Select(v => v.SourceName));
    }

    [Test]
    [Issue("45")]
    public void Calculated_field_is_saved_with_formula_and_without_record_values()
    {
        using var ms = CreateFileWithCalculatedField();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Save();
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var cachePart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            var cacheFields = cachePart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            Assert.AreEqual(3, cacheFields.Count);
            var calculatedField = cacheFields[2];
            Assert.AreEqual("Double", calculatedField.Name!.Value);
            Assert.AreEqual("Value*2", calculatedField.Formula!.Value);
            Assert.False(calculatedField.DatabaseField!.Value);
            Assert.Null(calculatedField.SharedItems);
            Assert.Null(cacheFields[0].Formula);
            Assert.Null(cacheFields[0].DatabaseField);

            var records = cachePart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Elements<PivotCacheRecord>().ToList();
            Assert.AreEqual(2, records.Count);
            Assert.True(records.All(r => r.ChildElements.Count == 2));

            var definition = document.WorkbookPart.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            Assert.AreEqual(3, definition.PivotFields!.Elements<PivotField>().Count());
        }

        using (var wb = new XLWorkbook(ms))
        {
            var pt = wb.Worksheet(1).PivotTable("pt");
            Assert.AreEqual(new[] { "Value", "Double" }, pt.Values.Select(v => v.SourceName));
        }
    }

    [Test]
    [Issue("45")]
    public void Refresh_keeps_calculated_field()
    {
        using var ms = CreateFileWithCalculatedField();
        using var wb = new XLWorkbook(ms);
        var pt = wb.Worksheet(1).PivotTable("pt");

        pt.PivotCache.Refresh();

        Assert.AreEqual(new[] { "Name", "Value", "Double" }, pt.PivotCache.FieldNames);
        Assert.AreEqual(new[] { "Value", "Double" }, pt.Values.Select(v => v.SourceName));
    }

    [Test]
    [Issue("45")]
    public void Refresh_removes_calculated_field_with_name_of_source_field()
    {
        using var ms = CreateFileWithCalculatedField();
        using var wb = new XLWorkbook(ms);
        var pt = wb.Worksheet(1).PivotTable("pt");
        wb.Worksheet(1).Cell("A1").Value = "Double";

        pt.PivotCache.Refresh();

        Assert.AreEqual(new[] { "Double", "Value" }, pt.PivotCache.FieldNames);
    }

    [Test]
    [Issue("45")]
    public void Group_field_is_saved_with_field_group_and_without_record_values()
    {
        using var ms = CreateFileWithGroupField();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Save();
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var cachePart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            var cacheFields = cachePart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            Assert.AreEqual(3, cacheFields.Count);
            Assert.AreEqual(2u, cacheFields[0].FieldGroup!.ParentId!.Value);
            var groupField = cacheFields[2];
            Assert.False(groupField.DatabaseField!.Value);
            Assert.AreEqual(0u, groupField.FieldGroup!.Base!.Value);
            Assert.AreEqual(2, groupField.FieldGroup.GetFirstChild<GroupItems>()!.ChildElements.Count);

            var records = cachePart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Elements<PivotCacheRecord>().ToList();
            Assert.True(records.All(r => r.ChildElements.Count == 2));
        }
    }

    [Test]
    [Issue("45")]
    public void Refresh_removes_group_field()
    {
        using var ms = CreateFileWithGroupField();
        using var wb = new XLWorkbook(ms);
        var pt = wb.Worksheet(1).PivotTable("pt");

        pt.PivotCache.Refresh();
        wb.Save();

        Assert.AreEqual(new[] { "Name", "Value" }, pt.PivotCache.FieldNames);
        using var document = SpreadsheetDocument.Open(ms, false);
        var cacheFields = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition!.CacheFields!.Elements<CacheField>();
        Assert.True(cacheFields.All(f => f.FieldGroup is null));
    }

    [TestCase("Double")]
    [TestCase("Name2")]
    [Issue("45")]
    public void Calculated_field_and_group_field_cannot_be_added_to_axis(string fieldName)
    {
        using var ms = CreateFileWithCalculatedAndGroupField();
        using var wb = new XLWorkbook(ms);
        var pt = wb.Worksheet(1).PivotTable("pt");

        Assert.Throws<InvalidOperationException>(() => pt.ColumnLabels.Add(fieldName));
        Assert.Throws<InvalidOperationException>(() => pt.ReportFilters.Add(fieldName));
    }

    [Test]
    [Issue("45")]
    public void Selected_value_of_group_field_filter_is_not_supported()
    {
        // The item of a group field is an index into the group items, not into the shared items.
        using var ms = CreateFileWithGroupField(groupFieldInFilter: true);
        using var wb = new XLWorkbook(ms);
        var filter = wb.Worksheet(1).PivotTable("pt").ReportFilters.Single();

        Assert.AreEqual("Name2", filter.SourceName);
        Assert.Throws<NotSupportedException>(() => _ = filter.SelectedValues);
        Assert.Throws<NotSupportedException>(() => filter.AddSelectedValue("Group1"));
    }

    [Test]
    [Issue("45")]
    public void Items_of_base_field_of_group_field_are_not_supported()
    {
        // A new shared item of the base field makes the groups of the group field incorrect.
        using var ms = CreateFileWithGroupField();
        using var wb = new XLWorkbook(ms);
        var pt = wb.Worksheet(1).PivotTable("pt");
        var nameField = pt.RowLabels.Get("Name");

        Assert.Throws<NotSupportedException>(() => nameField.StyleFormats.DataValuesFormat.AndWith(nameField, _ => true));
    }

    [Test]
    [Issue("45")]
    public void Cache_without_database_fields_has_no_records()
    {
        using var wb = new XLWorkbook();
        var cache = wb.PivotCachesInternal.Add(new XLPivotSourceConnection(0));
        cache.AddCachedField("Double", new XLPivotCacheValues(new XLPivotCacheSharedItems(), new XLPivotCacheValuesStats())
        {
            Formula = "Value*2",
            IsDatabaseField = false,
        });

        Assert.AreEqual(0, cache.RecordCount);
        Assert.IsEmpty(cache.GetDatabaseFieldIndexes());
    }

    [TestCase("shared-group-field.xlsx", 2)]
    [TestCase("shared-nested-dategroup.xlsx", 3)]
    [Issue("45")]
    public void File_with_group_fields_keeps_field_groups_after_save(string fileName, int fieldGroupCount)
    {
        using var ms = new MemoryStream();
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath($"TryToLoad.LO.xlsx.pivot_table.{fileName}")))
        using (var wb = new XLWorkbook(stream))
        {
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var cachePart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            var cacheFields = cachePart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
            Assert.AreEqual(fieldGroupCount, cacheFields.Count(f => f.FieldGroup is not null));
            var databaseFieldCount = cacheFields.Count(f => f.DatabaseField?.Value ?? true);
            var records = cachePart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Elements<PivotCacheRecord>();
            Assert.True(records.All(r => r.ChildElements.Count == databaseFieldCount));
        }

        Assert.DoesNotThrow(() => _ = new XLWorkbook(ms));
    }

    [Test]
    public void Error_record_value_is_saved_as_error_item()
    {
        // ECMA-376 uses the element 'e' for an error value in a record.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Value";
            ws.Cell("A2").Value = "Cake";
            ws.Cell("B2").Value = XLError.DivisionByZero;
            ws.PivotTables.Add("pt", ws.Cell("D1"), ws.Range("A1:B2")).RowLabels.Add("Name");
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var records = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotTableCacheRecordsPart!.PivotCacheRecords!;
            var errorItem = records.Elements<PivotCacheRecord>().Single().Elements<ErrorItem>().Single();
            Assert.AreEqual("#DIV/0!", errorItem.Val!.Value);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var cache = (XLPivotCache)wb.Worksheet(1).PivotTable("pt").PivotCache;
            Assert.AreEqual(new XLCellValue[] { XLError.DivisionByZero }, cache.GetFieldValues(1).GetCellValues());
        }
    }

    /// <summary>
    /// Creates a file with a calculated field <c>Double</c>. The pivot table has the field in the values.
    /// Excel saves a calculated field without shared items and the records have no values for it.
    /// </summary>
    private static MemoryStream CreateFileWithCalculatedField()
    {
        var ms = CreateFile();
        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var cacheDefinition = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition!;
            cacheDefinition.CacheFields!.AppendChild(new CacheField
            {
                Name = "Double",
                NumberFormatId = 0,
                Formula = "Value*2",
                DatabaseField = false,
            });
            cacheDefinition.CacheFields.Count = 3;

            var definition = document.WorkbookPart.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            definition.PivotFields!.AppendChild(new PivotField { DataField = true, ShowAll = false });
            definition.PivotFields.Count = 3;
            definition.DataFields!.AppendChild(new DataField { Name = "Sum of Double", Field = 2, BaseField = 0, BaseItem = 0 });
            definition.DataFields.Count = 2;
            definition.ColumnFields ??= new ColumnFields();
            definition.ColumnFields.AppendChild(new Field { Index = -2 });
            definition.ColumnFields.Count = (uint)definition.ColumnFields.ChildElements.Count;
        }

        return ms;
    }

    /// <summary>
    /// Creates a file with the calculated field <c>Double</c> and the group field <c>Name2</c>.
    /// The pivot table has no axis for the group field.
    /// </summary>
    private static MemoryStream CreateFileWithCalculatedAndGroupField()
    {
        var ms = CreateFileWithCalculatedField();
        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var cacheDefinition = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition!;
            cacheDefinition.CacheFields!.AppendChild(CreateGroupField());
            cacheDefinition.CacheFields.Count = 4;

            var definition = document.WorkbookPart.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            definition.PivotFields!.AppendChild(new PivotField { ShowAll = false });
            definition.PivotFields.Count = 4;
        }

        return ms;
    }

    /// <summary>
    /// Creates a file with a group field <c>Name2</c> that groups the items of the <c>Name</c> field.
    /// The records have no values for the group field.
    /// </summary>
    /// <param name="groupFieldInFilter">If true, the pivot table has the group field in the filters and the selected item is <c>Group2</c>.</param>
    private static MemoryStream CreateFileWithGroupField(bool groupFieldInFilter = false)
    {
        var ms = CreateFile();
        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var cacheDefinition = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition!;
            var cacheFields = cacheDefinition.CacheFields!.Elements<CacheField>().ToList();
            cacheFields[0].FieldGroup = new FieldGroup { ParentId = 2 };
            cacheDefinition.CacheFields!.AppendChild(CreateGroupField());
            cacheDefinition.CacheFields.Count = 3;

            var definition = document.WorkbookPart.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            if (groupFieldInFilter)
            {
                definition.PivotFields!.AppendChild(new PivotField(
                    new Items(new Item { Index = 0 }, new Item { Index = 1 }, new Item { ItemType = ItemValues.Default }) { Count = 3 })
                {
                    Axis = PivotTableAxisValues.AxisPage,
                    ShowAll = false,
                });
                definition.PageFields = new PageFields(new PageField { Field = 2, Item = 1, Hierarchy = -1 }) { Count = 1 };
            }
            else
            {
                definition.PivotFields!.AppendChild(new PivotField { ShowAll = false });
            }

            definition.PivotFields.Count = 3;
        }

        return ms;
    }

    /// <summary>
    /// Creates a group field <c>Name2</c>. It puts the item <c>Cake</c> into <c>Group1</c>
    /// and the item <c>Pie</c> into <c>Group2</c>.
    /// </summary>
    private static CacheField CreateGroupField()
    {
        return new CacheField(
            new FieldGroup(
                new DiscreteProperties(new FieldItem { Val = 0 }, new FieldItem { Val = 1 }) { Count = 2 },
                new GroupItems(new StringItem { Val = "Group1" }, new StringItem { Val = "Group2" }) { Count = 2 })
            {
                Base = 0,
            })
        {
            Name = "Name2",
            NumberFormatId = 0,
            DatabaseField = false,
        };
    }

    private static MemoryStream CreateFile()
    {
        var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "Value"),
            ("Cake", 7),
            ("Pie", 4),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("D1"), data);
        pt.RowLabels.Add("Name");
        pt.Values.Add("Value");
        wb.SaveAs(ms);
        return ms;
    }
}
