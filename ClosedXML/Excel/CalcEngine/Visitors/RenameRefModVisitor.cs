using ClosedXML.Parser;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A factory to rename named reference object (sheets, tables ect.).
/// </summary>
internal class RenameRefModVisitor : RefModVisitor
{
    private readonly Dictionary<string, string?>? _sheets;
    private readonly Dictionary<string, string>? _tables;

    /// <summary>
    /// A mapping of sheets, from old name (key) to a new name (value).
    /// The <c>null</c> value indicates sheet has been deleted.
    /// </summary>
    internal IReadOnlyDictionary<string, string?> Sheets
    {
        init => _sheets = value.ToDictionary(x => x.Key, x => x.Value, XLHelper.SheetComparer);
    }

    internal IReadOnlyDictionary<string, string> Tables
    {
        init => _tables = value.ToDictionary(x => x.Key, x => x.Value, XLHelper.NameComparer);
    }

    protected override ReadOnlySpan<char> ModifySheet(ModContext ctx, ReadOnlySpan<char> sheetName)
    {
        // The span of a null name is empty. An empty result means that the sheet has been deleted.
        if (_sheets is not null && _sheets.GetAlternateLookup<ReadOnlySpan<char>>().TryGetValue(sheetName, out var newName))
            return newName.AsSpan();

        return sheetName;
    }

    protected override ReadOnlySpan<char> ModifyTable(ModContext ctx, ReadOnlySpan<char> tableName)
    {
        if (_tables is not null && _tables.GetAlternateLookup<ReadOnlySpan<char>>().TryGetValue(tableName, out var newName))
            return newName;

        return tableName;
    }
}
