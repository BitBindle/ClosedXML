using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel.CalcEngine;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Excel
{
    using Ap;
    using ClosedXML.IO;
    using Drawings;
    using Op;

    public partial class XLWorkbook
    {
        private void Load(String file)
        {
            LoadSheets(file);
        }

        private void Load(Stream stream)
        {
            LoadSheets(stream);
        }

        private void LoadSheets(String fileName)
        {
            using (var dSpreadsheet = SpreadsheetDocument.Open(fileName, false))
                LoadSpreadsheetDocument(dSpreadsheet);
        }

        private void LoadSheets(Stream stream)
        {
            using (var dSpreadsheet = SpreadsheetDocument.Open(stream, false))
                LoadSpreadsheetDocument(dSpreadsheet);
        }

        private void LoadSheetsFromTemplate(String fileName)
        {
            using (var dSpreadsheet = SpreadsheetDocument.CreateFromTemplate(fileName))
                LoadSpreadsheetDocument(dSpreadsheet);

            // If we load a workbook as a template, we have to treat it as a "new" workbook.
            // The original file will NOT be copied into place before changes are applied
            // Hence all loaded RelIds have to be cleared
            ResetAllRelIds();
        }

        private void ResetAllRelIds()
        {
            foreach (var pc in PivotCachesInternal)
                pc.WorkbookCacheRelId = null;

            var sheetId = 1u;
            foreach (var ws in WorksheetsInternal)
            {
                // Ensure unique sheetId for each sheet. 
                ws.SheetId = sheetId++;
                ws.RelId = null;

                foreach (var pt in ws.PivotTables.Cast<XLPivotTable>())
                {
                    pt.CacheDefinitionRelId = null;
                    pt.RelId = null;
                }

                foreach (var picture in ws.Pictures.Cast<XLPicture>())
                    picture.RelId = null;

                foreach (var table in ws.Tables.Cast<XLTable>())
                    table.RelId = null;
            }
        }

        private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
        {
            var context = new LoadContext();
            ShapeIdManager = new XLIdManager();
            SetProperties(dSpreadsheet);

            var workbookPart = dSpreadsheet.WorkbookPart ?? throw PartStructureException.RequiredElementIsMissing("workbook");
            var workbook = workbookPart.Workbook ?? throw PartStructureException.RequiredElementIsMissing("workbook");

            LoadWorkbookTheme(workbookPart.ThemePart, this);

            if (dSpreadsheet.CustomFilePropertiesPart?.Properties is { } customFileProperties)
            {
                foreach (var m in customFileProperties.Elements<CustomDocumentProperty>())
                {
                    String? name = m.Name?.Value;

                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    if (m.VTLPWSTR != null)
                        CustomProperties.Add(name, m.VTLPWSTR.Text);
                    else if (m.VTFileTime != null)
                    {
                        CustomProperties.Add(name,
                                             DateTime.ParseExact(m.VTFileTime.Text, "yyyy'-'MM'-'dd'T'HH':'mm':'ssK",
                                                                 CultureInfo.InvariantCulture));
                    }
                    else if (m.VTDouble != null)
                    {
                        var number = Double.Parse(m.VTDouble.Text, CultureInfo.InvariantCulture);
                        CustomProperties.Add(name, number);
                    }
                    else if (m.VTBool != null)
                        CustomProperties.Add(name, m.VTBool.Text == "true");
                }
            }

            var wbProps = workbook.WorkbookProperties;
            if (wbProps != null)
                Use1904DateSystem = OpenXmlHelper.GetBooleanValueAsBool(wbProps.Date1904, false);

            var wbFilesharing = workbook.FileSharing;
            if (wbFilesharing != null)
            {
                FileSharing.ReadOnlyRecommended = OpenXmlHelper.GetBooleanValueAsBool(wbFilesharing.ReadOnlyRecommended, false);
                FileSharing.UserName = wbFilesharing.UserName?.Value;
            }

            LoadWorkbookProtection(workbook.WorkbookProtection, this);

            var calculationProperties = workbook.CalculationProperties;
            if (calculationProperties != null)
            {
                var calculateMode = calculationProperties.CalculationMode;
                if (calculateMode != null)
                    CalculateMode = calculateMode.Value.ToClosedXml();

                var calculationOnSave = calculationProperties.CalculationOnSave;
                if (calculationOnSave != null)
                    CalculationOnSave = calculationOnSave.Value;

                var forceFullCalculation = calculationProperties.ForceFullCalculation;
                if (forceFullCalculation != null)
                    ForceFullCalculation = forceFullCalculation.Value;

                var fullCalculationOnLoad = calculationProperties.FullCalculationOnLoad;
                if (fullCalculationOnLoad != null)
                    FullCalculationOnLoad = fullCalculationOnLoad.Value;

                var fullPrecision = calculationProperties.FullPrecision;
                if (fullPrecision != null)
                    FullPrecision = fullPrecision.Value;

                var referenceMode = calculationProperties.ReferenceMode;
                if (referenceMode != null)
                    ReferenceStyle = referenceMode.Value.ToClosedXml();
            }

            var efp = dSpreadsheet.ExtendedFilePropertiesPart;
            if (efp != null && efp.Properties != null)
            {
                if (efp.Properties.GetFirstChild<Company>() is { } company)
                    Properties.Company = company.Text;

                if (efp.Properties.GetFirstChild<Manager>() is { } manager)
                    Properties.Manager = manager.Text;
            }

            // Styles part is optional and each workbook must have a default style
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart is not null)
            {
                using var xmlReader = CreateTreeReader(stylesPart);
                var stylesReader = new StylesReader(xmlReader, Styles);
                stylesReader.Load();
            }
            else
            {
                Styles.Initialize();
            }

            // Spec says each package must have exactly one SST part, but some packages don't have that.
            List<OneOf<string, XLImmutableRichText>> sst = [];
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sstPart is not null)
            {
                using var sstXmlReader = CreateTreeReader(sstPart);
                var sstReader = new SstReader(sstXmlReader, Styles);
                sst = sstReader.ParseSst();
            }

            // TODO Styles: Verify the column width is same as DefaultColumnWidth even if normal style is missing
            ColumnWidth = XLHelper.CalculateColumnWidth(8, Format.Font, this);

            // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
            // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
            // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
            // That consistency point isn't required yet but could be taken advantage of in the future.
            var sheets = workbook.Sheets ?? throw PartStructureException.RequiredElementIsMissing("sheets");
            Int32 position = 0;
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                position++;
                var sheetName = dSheet.Name?.Value ?? throw PartStructureException.MissingAttribute("name");
                var sheetId = dSheet.SheetId?.Value ?? throw PartStructureException.MissingAttribute("sheetId");
                var sheetRelId = dSheet.Id?.Value;

                if (string.IsNullOrEmpty(sheetRelId))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    var emptySheet = WorksheetsInternal.Add(sheetName, position, sheetId);
                    if (dSheet.State != null)
                        emptySheet.Visibility = dSheet.State.Value.ToClosedXml();

                    continue;
                }

                // Although relationship to worksheet is most common, there can be other types
                // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
                // of unsupported sheets and copy them when saving. See Codeplex #6932.
                var worksheetPart = workbookPart.GetPartById(sheetRelId) as WorksheetPart;
                if (worksheetPart == null)
                {
                    UnsupportedSheets.Add(new UnsupportedSheet { SheetId = sheetId, Position = position });
                    continue;
                }

                var ws = WorksheetsInternal.Add(sheetName, position, sheetId);
                ws.RelId = sheetRelId;

                if (dSheet.State != null)
                    ws.Visibility = dSheet.State.Value.ToClosedXml();
            }

            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                var sheetName = dSheet.Name?.Value ?? throw PartStructureException.MissingAttribute("name");
                var sheetRelId = dSheet.Id?.Value;

                if (string.IsNullOrEmpty(sheetRelId))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    continue;
                }

                // Although relationship to worksheet is most common, there can be other types
                // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
                // of unsupported sheets and copy them when saving. See Codeplex #6932.
                var worksheetPart = workbookPart.GetPartById(sheetRelId) as WorksheetPart;
                if (worksheetPart == null)
                {
                    continue;
                }

                if (!WorksheetsInternal.TryGetWorksheet(sheetName, out var ws))
                {
                    // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                    continue;
                }

                var worksheetPartReader = new WorksheetPartReader();
                worksheetPartReader.LoadWorksheet(ws, worksheetPart, sst, context);

                ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

                #region LoadTables

                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts)
                {
                    var relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                    var dTable = tableDefinitionPart.Table ?? throw PartStructureException.RequiredElementIsMissing("table");

                    String reference = dTable.Reference?.Value ?? throw PartStructureException.MissingAttribute("ref");
                    String tableName = dTable.Name?.Value ?? dTable.DisplayName?.Value ?? string.Empty;
                    if (String.IsNullOrWhiteSpace(tableName))
                        throw new InvalidDataException("The table name is missing.");

                    // Other producers can write a name that breaks a rule of Excel. Keep it as it is.
                    var tableRange = ws.Range(reference) ?? throw PartStructureException.InvalidAttributeFormat("ref", reference);
                    var xlTable = (XLTable)ws.Table(tableRange, tableName, addToTables: true, setAutofilter: false, checkExcelNameRules: false);
                    var tableColumns = dTable.TableColumns ?? throw PartStructureException.RequiredElementIsMissing("tableColumns");
                    xlTable.RelId = relId;

                    // Add columns to the table
                    foreach (TableColumn tableColumn in tableColumns)
                    {
                        var fieldName = GetTableColumnName(tableColumn.Name?.Value ?? throw PartStructureException.MissingAttribute("name"));
                        var xlField = xlTable.AddField(fieldName);

                        if (tableColumn.HeaderRowDifferentialFormattingId is { } headerDxfId)
                            xlField.HeaderFormatValue = Styles.DifferentialFormats[checked((int)headerDxfId.Value)];

                        if (tableColumn.DataFormatId is { } dataDxfId)
                            xlField.DataFormatValue = Styles.DifferentialFormats[checked((int)dataDxfId.Value)];

                        if (tableColumn.TotalsRowDifferentialFormattingId is { } totalsDxfId)
                            xlField.TotalFormatValue = Styles.DifferentialFormats[checked((int)totalsDxfId.Value)];
                    }

                    if (dTable.HeaderRowCount != null && dTable.HeaderRowCount == 0)
                    {
                        xlTable._showHeaderRow = false;
                    }
                    else
                    {
                        xlTable.InitializeAutoFilter();
                    }

                    if (dTable.TotalsRowCount != null && dTable.TotalsRowCount.Value > 0)
                        xlTable._showTotalsRow = true;

                    if (dTable.TableStyleInfo != null)
                    {
                        if (dTable.TableStyleInfo.ShowFirstColumn != null)
                            xlTable.EmphasizeFirstColumn = dTable.TableStyleInfo.ShowFirstColumn.Value;
                        if (dTable.TableStyleInfo.ShowLastColumn != null)
                            xlTable.EmphasizeLastColumn = dTable.TableStyleInfo.ShowLastColumn.Value;
                        if (dTable.TableStyleInfo.ShowRowStripes != null)
                            xlTable.ShowRowStripes = dTable.TableStyleInfo.ShowRowStripes.Value;
                        if (dTable.TableStyleInfo.ShowColumnStripes != null)
                            xlTable.ShowColumnStripes = dTable.TableStyleInfo.ShowColumnStripes.Value;
                        if (dTable.TableStyleInfo.Name?.Value is { } themeName)
                        {
                            var theme = XLTableTheme.FromName(themeName);
                            if (theme != null)
                                xlTable.Theme = theme;
                            else
                                xlTable.Theme = new XLTableTheme(themeName);
                        }
                        else
                            xlTable.Theme = XLTableTheme.None;
                    }

                    if (dTable.AutoFilter != null)
                    {
                        xlTable.ShowAutoFilter = true;
                        AutoFilterReader.LoadAutoFilterColumns(dTable.AutoFilter, xlTable.AutoFilter);
                    }
                    else
                        xlTable.ShowAutoFilter = false;

                    if (xlTable.ShowTotalsRow)
                    {
                        foreach (var tableColumn in tableColumns.Cast<TableColumn>())
                        {
                            var tableColumnName = GetTableColumnName(tableColumn.Name?.Value ?? throw PartStructureException.MissingAttribute("name"));
                            if (tableColumn.TotalsRowFunction != null)
                                xlTable.Field(tableColumnName).TotalsRowFunction =
                                    tableColumn.TotalsRowFunction.Value.ToClosedXml();

                            if (tableColumn.TotalsRowFormula != null)
                                xlTable.Field(tableColumnName).TotalsRowFormulaA1 =
                                    tableColumn.TotalsRowFormula.Text;

                            if (tableColumn.TotalsRowLabel != null)
                                xlTable.Field(tableColumnName).TotalsRowLabel = tableColumn.TotalsRowLabel.Value;
                        }
                        if (xlTable.AutoFilter != null)
                            xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                                                    xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                                                    xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
                    }
                    else if (xlTable.AutoFilter != null)
                        xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);
                }

                #endregion LoadTables

                LoadDrawings(worksheetPart, ws);

                #region LoadComments

                if (worksheetPart.WorksheetCommentsPart != null)
                {
                    var root = worksheetPart.WorksheetCommentsPart.Comments ?? throw PartStructureException.RequiredElementIsMissing("comments");
                    var authorsElement = root.GetFirstChild<Authors>() ?? throw PartStructureException.RequiredElementIsMissing("authors");
                    var commentList = root.GetFirstChild<CommentList>() ?? throw PartStructureException.RequiredElementIsMissing("commentList");
                    var authors = authorsElement.ChildElements.OfType<Author>().ToList();
                    var comments = commentList.ChildElements.OfType<Comment>().ToList();

                    // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                    var shapes = GetCommentShapes(worksheetPart);

                    for (var i = 0; i < comments.Count; i++)
                    {
                        var c = comments[i];

                        XElement? shape = null;
                        if (i < shapes.Count)
                            shape = shapes[i];

                        // find cell by reference
                        var commentReference = c.Reference?.Value ?? throw PartStructureException.MissingAttribute("ref");
                        var cell = ws.Cell(commentReference) ?? throw PartStructureException.InvalidAttributeFormat("ref", commentReference);

                        var shapeIdText = (shape?.Attribute("id")?.Value).AsSpan();
                        if (shapeIdText.StartsWith("_x0000_s", StringComparison.Ordinal))
                            shapeIdText = shapeIdText[8..];

                        int? shapeId = int.TryParse(shapeIdText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sid) ? sid : null;
                        var xlComment = cell.CreateComment(shapeId);

                        var authorId = c.AuthorId?.Value ?? throw PartStructureException.MissingAttribute("authorId");
                        xlComment.Author = XStringConvert.Decode(authors[(int)authorId].InnerText);
                        ShapeIdManager.Add(xlComment.ShapeId);

                        var commentText = c.GetFirstChild<CommentText>() ?? throw PartStructureException.RequiredElementIsMissing("text");
                        var runs = commentText.Elements<Run>();
                        foreach (var run in runs)
                        {
                            var runProperties = run.RunProperties;
                            var runText = run.Text ?? throw PartStructureException.RequiredElementIsMissing("t");
                            String text = XStringConvert.Decode(runText.InnerText).FixNewLines();
                            var rt = xlComment.AddText(text);
                            OpenXmlHelper.LoadFont(runProperties, rt);
                        }

                        if (shape != null)
                        {
                            LoadShapeProperties(xlComment, shape);

                            var clientData = shape.Elements().First(e => e.Name.LocalName == "ClientData");
                            LoadClientData(xlComment, clientData);

                            // The textbox element is optional (CT_Shape of VML).
                            var textBox = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "textbox");
                            if (textBox != null)
                                LoadTextBox(xlComment, textBox);

                            var alt = shape.Attribute("alt");
                            if (alt != null) xlComment.Style.Web.SetAlternateText(alt.Value);

                            LoadColorsAndLines(xlComment, shape);

                            //var insetmode = (string)shape.Attributes().First(a=> a.Name.LocalName == "insetmode");
                            //xlComment.Style.Margins.Automatic = insetmode != null && insetmode.Equals("auto");
                        }
                    }
                }

                #endregion LoadComments
            }

            var bookViews = workbook.BookViews;
            if (bookViews != null && bookViews.FirstOrDefault() is WorkbookView workbookView)
            {
                if (workbookView.ActiveTab == null || !workbookView.ActiveTab.HasValue)
                {
                    Worksheets.First().SetTabActive().Unhide();
                }
                else
                {
                    UnsupportedSheet? unsupportedSheet =
                        UnsupportedSheets.FirstOrDefault(us => us.Position == (Int32)(workbookView.ActiveTab.Value + 1));
                    if (unsupportedSheet != null)
                        unsupportedSheet.IsActive = true;
                    else
                    {
                        Worksheet((Int32)(workbookView.ActiveTab.Value + 1)).SetTabActive();
                    }
                }
            }
            LoadDefinedNames(workbook);

            // Read cache definition before table definition
            foreach (var pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
            {
                var pivotCache = PivotTableCacheDefinitionPartReader.Load(workbookPart, pivotTableCacheDefinitionPart, this);
                if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart is { } recordsPart)
                {
                    using var reader = CreateTreeReader(recordsPart);
                    var recordsReader = new PivotCacheRecordsReader(reader, pivotCache);
                    recordsReader.ReadRecordsToCache();
                }
            }

            // Delay loading of pivot tables until all sheets have been loaded
            foreach (var dSheet in sheets.OfType<Sheet>())
            {
                var sheetRelId = dSheet.Id?.Value;
                if (string.IsNullOrEmpty(sheetRelId))
                {
                    // Some non-Excel producers create sheets with empty relId.
                    continue;
                }

                // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
                var worksheetPart = workbookPart.GetPartById(sheetRelId) as WorksheetPart;

                if (worksheetPart is not null)
                {
                    var sheetName = dSheet.Name?.Value ?? throw PartStructureException.MissingAttribute("name");
                    var ws = (XLWorksheet)WorksheetsInternal.Worksheet(sheetName);

                    foreach (var pivotTablePart in worksheetPart.PivotTableParts)
                    {
                        PivotTableDefinitionPartReader.Load(workbookPart, pivotTablePart, worksheetPart, ws, context);
                    }
                }
            }
        }

        private void LoadDrawings(WorksheetPart wsPart, XLWorksheet ws)
        {
            if (wsPart.DrawingsPart?.WorksheetDrawing is { } worksheetDrawing)
            {
                var drawingsPart = wsPart.DrawingsPart;
                var unnamedPictures = new List<XLPicture>();

                foreach (var anchor in worksheetDrawing.ChildElements)
                {
                    var imgId = GetImageRelIdFromAnchor(anchor);

                    //If imgId is null, we're probably dealing with a TextBox (or another shape) instead of a picture
                    if (imgId == null) continue;

                    var imagePart = drawingsPart.GetPartById(imgId);
                    using (var stream = imagePart.GetStream())
                    {
                        // The picture owns the copy, because the part stream can be forward only.
                        var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        var vsdp = GetPropertiesFromAnchor(anchor) ?? throw PartStructureException.RequiredElementIsMissing("cNvPr");
                        var pictureId = vsdp.Id?.Value ?? throw PartStructureException.MissingAttribute("id");

                        var picture = ((XLPictures)ws.Pictures).AddOwned(ms, vsdp.Name?.Value ?? String.Empty, Convert.ToInt32(pictureId));
                        picture.RelId = imgId;
                        if (String.IsNullOrEmpty(vsdp.Name))
                            unnamedPictures.Add(picture);

                        Xdr.ShapeProperties spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
                        picture.Placement = XLPicturePlacement.FreeFloating;

                        if (spPr.Transform2D?.Extents?.Cx?.Value is { } cx)
                            picture.Width = ConvertFromEnglishMetricUnits(cx, ws.Workbook.DpiX);

                        if (spPr.Transform2D?.Extents?.Cy?.Value is { } cy)
                            picture.Height = ConvertFromEnglishMetricUnits(cy, ws.Workbook.DpiY);

                        if (anchor is Xdr.AbsoluteAnchor absoluteAnchor)
                        {
                            var position = absoluteAnchor.Position ?? throw PartStructureException.RequiredElementIsMissing("pos");
                            picture.MoveTo(
                                ConvertFromEnglishMetricUnits(position.X?.Value ?? throw PartStructureException.MissingAttribute("x"), ws.Workbook.DpiX),
                                ConvertFromEnglishMetricUnits(position.Y?.Value ?? throw PartStructureException.MissingAttribute("y"), ws.Workbook.DpiY)
                            );
                        }
                        else if (anchor is Xdr.OneCellAnchor oneCellAnchor)
                        {
                            var from = LoadMarker(ws, oneCellAnchor.FromMarker ?? throw PartStructureException.RequiredElementIsMissing("from"));
                            picture.MoveTo(from.Cell, from.Offset);
                        }
                        else if (anchor is Xdr.TwoCellAnchor twoCellAnchor)
                        {
                            var from = LoadMarker(ws, twoCellAnchor.FromMarker ?? throw PartStructureException.RequiredElementIsMissing("from"));
                            var to = LoadMarker(ws, twoCellAnchor.ToMarker ?? throw PartStructureException.RequiredElementIsMissing("to"));

                            if (twoCellAnchor.EditAs == null || !twoCellAnchor.EditAs.HasValue || twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell)
                            {
                                picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
                            {
                                var shapeProperties = twoCellAnchor.Descendants<Xdr.ShapeProperties>().FirstOrDefault();
                                if (shapeProperties?.Transform2D?.Offset is { X.Value: var offsetX, Y.Value: var offsetY })
                                {
                                    picture.MoveTo(
                                        ConvertFromEnglishMetricUnits(offsetX, ws.Workbook.DpiX),
                                        ConvertFromEnglishMetricUnits(offsetY, ws.Workbook.DpiY)
                                    );
                                }
                            }
                            else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
                            {
                                picture.MoveTo(from.Cell, from.Offset);
                            }
                        }
                    }
                }

                // A picture that loads later can have the generated name of an unnamed picture.
                ws.Pictures.RenameDuplicates(unnamedPictures);
            }
        }

        private static Int32 ConvertFromEnglishMetricUnits(long emu, double resolution)
        {
            return Convert.ToInt32(emu * resolution / 914400);
        }

        private static XLMarker LoadMarker(XLWorksheet ws, Xdr.MarkerType marker)
        {
            var row = Math.Min(XLHelper.MaxRowNumber, Math.Max(1, MarkerValue(marker.RowId, "row") + 1));
            var column = Math.Min(XLHelper.MaxColumnNumber, Math.Max(1, MarkerValue(marker.ColumnId, "col") + 1));
            return new XLMarker(
                ws.Cell(row, column),
                new System.Drawing.Point(
                    ConvertFromEnglishMetricUnits(MarkerValue(marker.ColumnOffset, "colOff"), ws.Workbook.DpiX),
                    ConvertFromEnglishMetricUnits(MarkerValue(marker.RowOffset, "rowOff"), ws.Workbook.DpiY)
                )
            );

            static Int32 MarkerValue(OpenXmlLeafTextElement? element, String elementName)
            {
                return Convert.ToInt32((element ?? throw PartStructureException.RequiredElementIsMissing(elementName)).InnerText);
            }
        }

        #region Comment Helpers

        private static IList<XElement> GetCommentShapes(WorksheetPart worksheetPart)
        {
            // Cannot get this to return Vml.Shape elements
            foreach (var vmlPart in worksheetPart.VmlDrawingParts)
            {
                using (var stream = vmlPart.GetStream(FileMode.Open))
                {
                    var xdoc = XDocumentExtensions.Load(stream);
                    if (xdoc == null)
                        continue;

                    var root = xdoc.Root?.Element("xml") ?? xdoc.Root;

                    if (root == null)
                        continue;

                    var shapes = root
                        .Elements(XName.Get("shape", "urn:schemas-microsoft-com:vml"))
                        .Where(e => e.Attribute("type")?.Value is { } type && IsCommentShapeType(type))
                        .ToList();

                    if (shapes != null)
                        return shapes;
                }
            }

            throw new ArgumentException("Could not load comments file");

            static bool IsCommentShapeType(string type)
            {
                return type.StartsWith('#') &&
                       (type.AsSpan(1).SequenceEqual(XLConstants.Comment.ShapeTypeId) ||
                        type.AsSpan(1).SequenceEqual(XLConstants.Comment.AlternateShapeTypeId));
            }
        }

        #endregion Comment Helpers

        private String GetTableColumnName(string name)
        {
            return name.Replace("_x000a_", Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");
        }

        private void LoadColorsAndLines<T>(IXLDrawing<T> drawing, XElement shape)
        {
            var strokeColor = shape.Attribute(@"strokecolor");
            if (strokeColor is not null)
                drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);

            var strokeWeight = shape.Attribute(@"strokeweight");
            if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
                drawing.Style.ColorsAndLines.LineWeight = lineWeight;

            var fillColor = shape.Attribute(@"fillcolor");
            if (fillColor is not null)
                drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);

            var fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
            if (fill != null)
            {
                var opacity = fill.Attribute("opacity");
                if (opacity != null)
                {
                    drawing.Style.ColorsAndLines.FillTransparency = ParseOpacity(opacity.Value);
                }
            }

            var stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
            if (stroke != null)
            {
                var opacity = stroke.Attribute("opacity");
                if (opacity != null)
                {
                    drawing.Style.ColorsAndLines.LineTransparency = ParseOpacity(opacity.Value);
                }

                var dashStyle = stroke.Attribute("dashstyle");
                if (dashStyle != null)
                {
                    var dashStyleVal = dashStyle.Value.ToLowerInvariant();
                    if (dashStyleVal == "1 1" || dashStyleVal == "shortdot")
                    {
                        var endCap = stroke.Attribute("endcap");
                        if (endCap != null && endCap.Value == "round")
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.RoundDot;
                        else
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.SquareDot;
                    }
                    else
                    {
                        switch (dashStyleVal)
                        {
                            case "dash": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.Dash; break;
                            case "dashdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.DashDot; break;
                            case "longdash": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDash; break;
                            case "longdashdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDot; break;
                            case "longdashdotdot": drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDotDot; break;
                        }
                    }
                }

                var lineStyle = stroke.Attribute("linestyle");
                if (lineStyle != null)
                {
                    var lineStyleVal = lineStyle.Value.ToLowerInvariant();
                    switch (lineStyleVal)
                    {
                        case "single": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.Single; break;
                        case "thickbetweenthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickBetweenThin; break;
                        case "thickthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickThin; break;
                        case "thinthick": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThick; break;
                        case "thinthin": drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThin; break;
                    }
                }
            }

            // An opacity with the suffix 'f' is a fixed-point number. The unit is 1/65536.
            static double ParseOpacity(string opacity)
            {
                var span = opacity.AsSpan();
                return span.EndsWith('f')
                    ? double.Parse(span[..^1], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture) / 65536.0
                    : double.Parse(span, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
            }
        }

        private void LoadTextBox<T>(IXLDrawing<T> xlDrawing, XElement textBox)
        {
            var attStyle = textBox.Attribute("style");
            if (attStyle != null) LoadTextBoxStyle(xlDrawing, attStyle);

            var attInset = textBox.Attribute("inset");
            if (attInset != null) LoadTextBoxInset(xlDrawing, attInset);
        }

        private void LoadTextBoxInset<T>(IXLDrawing<T> xlDrawing, XAttribute attInset)
        {
            var inset = attInset.Value.AsSpan();
            // The fifth range takes all the text after the fourth value.
            Span<Range> split = stackalloc Range[5];
            if (inset.Split(split, ',') < 4)
                throw PartStructureException.InvalidAttributeFormat("inset", attInset.Value);

            xlDrawing.Style.Margins.Left = GetInsetInInches(inset[split[0]], DpiX);
            xlDrawing.Style.Margins.Top = GetInsetInInches(inset[split[1]], DpiY);
            xlDrawing.Style.Margins.Right = GetInsetInInches(inset[split[2]], DpiX);
            xlDrawing.Style.Margins.Bottom = GetInsetInInches(inset[split[3]], DpiY);
        }

        /// <summary>
        /// List of all VML length units and their conversion. Key is a name, value is a conversion
        /// function to EMU. See <a href="https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-units">documentation</a>.
        /// </summary>
        /// <remarks>
        /// OI-29500 says <em>Office also uses EMUs throughout VML as a valid unit system</em>.
        /// Relative units conversions are guesstimated by how Excel 2022 behaves for inset
        /// attribute of <c>TextBox</c> element of a note/comment. Generally speaking, Excel
        /// converts relative values to physical length (e.g. <c>px</c> to <c>pt</c>) and saves
        /// them as such. The <c>ex</c>/<c>em</c> units are not interpreted as described in the
        /// doc, but as 1/90th or an inch. The <c>%</c> seems to be always 0.
        /// </remarks>
        private static readonly Dictionary<string, Func<double, double, Emu?>> VmlLengthUnits = new()
        {
            {"in", (value, _) => Emu.From(value, AbsLengthUnit.Inch) },
            {"cm", (value, _) => Emu.From(value, AbsLengthUnit.Centimeter) },
            {"mm", (value, _) => Emu.From(value, AbsLengthUnit.Millimeter) },
            {"pt", (value, _) => Emu.From(value, AbsLengthUnit.Point) },
            {"pc", (value, _) => Emu.From(value, AbsLengthUnit.Pica) },
            {"emu", (value, _) => Emu.From(value , AbsLengthUnit.Emu) },
            {"px", (value, dpi) => Emu.From(value / dpi, AbsLengthUnit.Inch) },
            {"em", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
            {"ex", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
            {"%", (_, _) => Emu.ZeroPt },
        };

        private static double GetInsetInInches(ReadOnlySpan<char> value, double dpi)
        {
            var unit = value.Trim();
            foreach (var (unitName, conversion) in VmlLengthUnits)
            {
                if (unit.EndsWith(unitName, StringComparison.Ordinal) && Double.TryParse(unit[..^unitName.Length], NumberStyles.Float, CultureInfo.InvariantCulture, out var unitValue))
                {
                    var insetEmu = conversion(unitValue, dpi) ?? Emu.ZeroPt;
                    return insetEmu.To(AbsLengthUnit.Inch);
                }
            }

            // Excel treats no/unexpected unit as 0
            return 0;
        }

        private static void LoadTextBoxStyle<T>(IXLDrawing<T> xlDrawing, XAttribute attStyle)
        {
            var style = attStyle.Value.AsSpan();
            foreach (var pairRange in style.Split(';'))
            {
                if (!TrySplitStyleProperty(style[pairRange], out var attribute, out var value))
                    continue;

                var isVertical = false;
                if (attribute.Equals("mso-fit-shape-to-text", StringComparison.OrdinalIgnoreCase))
                {
                    xlDrawing.Style.Size.SetAutomaticSize(value is "t");
                }
                else if (attribute.Equals("mso-layout-flow-alt", StringComparison.OrdinalIgnoreCase))
                {
                    if (value is "bottom-to-top") xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop);
                    else if (value is "top-to-bottom") xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.Vertical);
                }
                else if (attribute.Equals("layout-flow", StringComparison.OrdinalIgnoreCase))
                {
                    isVertical = value is "vertical";
                }
                else if (attribute.Equals("mso-direction-alt", StringComparison.OrdinalIgnoreCase))
                {
                    if (value is "auto") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.Context;
                }
                else if (attribute.Equals("direction", StringComparison.OrdinalIgnoreCase))
                {
                    if (value is "RTL") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.RightToLeft;
                }

                if (isVertical && xlDrawing.Style.Alignment.Orientation == XLDrawingTextOrientation.LeftToRight)
                    xlDrawing.Style.Alignment.Orientation = XLDrawingTextOrientation.TopToBottom;
            }
        }

        /// <summary>
        /// Split a property of a VML style, for example <c>width:108pt</c>. The property must have
        /// exactly one colon.
        /// </summary>
        private static bool TrySplitStyleProperty(ReadOnlySpan<char> property, out ReadOnlySpan<char> name, out ReadOnlySpan<char> value)
        {
            var colonIndex = property.IndexOf(':');
            if (colonIndex < 0 || property[(colonIndex + 1)..].Contains(':'))
            {
                name = value = default;
                return false;
            }

            name = property[..colonIndex].Trim();
            value = property[(colonIndex + 1)..].Trim();
            return true;
        }

        private void LoadClientData<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var anchor = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor");
            if (anchor != null) LoadClientDataAnchor<T>(drawing, anchor);

            LoadDrawingPositioning<T>(drawing, clientData);
            LoadDrawingProtection<T>(drawing, clientData);

            var visible = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Visible");
            drawing.Visible = visible != null &&
                              (string.IsNullOrEmpty(visible.Value) ||
                               visible.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase));

            LoadDrawingHAlignment<T>(drawing, clientData);
            LoadDrawingVAlignment<T>(drawing, clientData);
        }

        private void LoadDrawingHAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var textHAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextHAlign");
            if (textHAlign != null)
                drawing.Style.Alignment.Horizontal = Enum.Parse<XLDrawingHorizontalAlignment>(textHAlign.Value, ignoreCase: true);
        }

        private void LoadDrawingVAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var textVAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
            if (textVAlign != null)
                drawing.Style.Alignment.Vertical = Enum.Parse<XLDrawingVerticalAlignment>(textVAlign.Value, ignoreCase: true);
        }

        private void LoadDrawingProtection<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var lockedElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Locked");
            var lockTextElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "LockText");
            var locked = string.Equals(lockedElement?.Value, "true", StringComparison.OrdinalIgnoreCase);
            var lockText = string.Equals(lockTextElement?.Value, "true", StringComparison.OrdinalIgnoreCase);
            drawing.Style.Protection.Locked = locked;
            drawing.Style.Protection.LockText = lockText;
        }

        private static void LoadDrawingPositioning<T>(IXLDrawing<T> drawing, XElement clientData)
        {
            var moveWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
            var sizeWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
            var moveWithCells = !string.Equals(moveWithCellsElement?.Value, "true", StringComparison.OrdinalIgnoreCase);
            var sizeWithCells = !string.Equals(sizeWithCellsElement?.Value, "true", StringComparison.OrdinalIgnoreCase);
            if (moveWithCells && !sizeWithCells)
                drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
            else if (moveWithCells && sizeWithCells)
                drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
            else
                drawing.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
        }

        private static void LoadClientDataAnchor<T>(IXLDrawing<T> drawing, XElement anchor)
        {
            var anchorValue = anchor.Value.AsSpan();
            Span<Range> location = stackalloc Range[5];
            if (anchorValue.Split(location, ',') < 4)
                throw PartStructureException.IncorrectElementFormat("Anchor");

            drawing.Position.Column = int.Parse(anchorValue[location[0]], NumberStyles.Integer, CultureInfo.InvariantCulture) + 1;
            drawing.Position.ColumnOffset = Double.Parse(anchorValue[location[1]], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture) / 7.5;
            drawing.Position.Row = int.Parse(anchorValue[location[2]], NumberStyles.Integer, CultureInfo.InvariantCulture) + 1;
            drawing.Position.RowOffset = Double.Parse(anchorValue[location[3]], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
        }

        private void LoadShapeProperties<T>(IXLDrawing<T> xlDrawing, XElement shape)
        {
            if (shape.Attribute("style") is not { } style)
                return;

            var styleValue = style.Value.AsSpan();
            foreach (var pairRange in styleValue.Split(';'))
            {
                if (!TrySplitStyleProperty(styleValue[pairRange], out var attribute, out var value))
                    continue;

                if (attribute.Equals("visibility", StringComparison.OrdinalIgnoreCase))
                {
                    xlDrawing.Visible = value.Equals("visible", StringComparison.OrdinalIgnoreCase);
                }
                else if (attribute.Equals("width", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryGetPtValue(value, out var ptWidth))
                        xlDrawing.Style.Size.Width = ptWidth / 7.5;
                }
                else if (attribute.Equals("height", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryGetPtValue(value, out var ptHeight))
                        xlDrawing.Style.Size.Height = ptHeight;
                }
                else if (attribute.Equals("z-index", StringComparison.OrdinalIgnoreCase))
                {
                    if (Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var zOrder))
                        xlDrawing.ZOrder = zOrder;
                }
            }
        }

        private static readonly (string Unit, double PointsPerUnit)[] KnownPtUnits =
        {
            ("pt", 1.0),
            ("in", 72.0),
            ("mm", 72.0 / 25.4),
        };

        /// <summary>
        /// Parse a VML length in points. A value without a known unit is a number of points.
        /// </summary>
        private static bool TryGetPtValue(ReadOnlySpan<char> value, out double result)
        {
            value = value.Trim();
            var pointsPerUnit = 1.0;
            foreach (var (unit, unitPoints) in KnownPtUnits)
            {
                if (value.EndsWith(unit, StringComparison.Ordinal))
                {
                    value = value[..^unit.Length];
                    pointsPerUnit = unitPoints;
                    break;
                }
            }

            if (Double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                result *= pointsPerUnit;
                return true;
            }

            result = 0d;
            return false;
        }

        private void LoadDefinedNames(Workbook workbook)
        {
            if (workbook.DefinedNames == null) return;

            foreach (var definedName in workbook.DefinedNames.OfType<DefinedName>())
            {
                var name = definedName.Name?.Value ?? throw PartStructureException.MissingAttribute("name");
                var visible = true;
                if (definedName.Hidden != null) visible = !BooleanValue.ToBoolean(definedName.Hidden);

                var localSheetId = -1;
                if (definedName.LocalSheetId?.HasValue ?? false) localSheetId = Convert.ToInt32(definedName.LocalSheetId.Value);

                if (name == "_xlnm.Print_Area")
                {
                    var fixedNames = validateDefinedNames(definedName.Text.Split(','));
                    foreach (string area in fixedNames)
                    {
                        if (area.Contains("["))
                        {
                            var ws = WorksheetsInternal.FirstOrDefault<XLWorksheet>(w => w.SheetId == (localSheetId + 1));
                            if (ws != null)
                            {
                                ws.PageSetup.PrintAreas.Add(area);
                            }
                        }
                        else
                        {
                            ParseReference(area, out String sheetName, out String sheetArea);
                            if (!(sheetArea.Equals("#REF") || sheetArea.EndsWith("#REF!") || sheetArea.Length == 0 || sheetName.Length == 0))
                                WorksheetsInternal.Worksheet(sheetName).PageSetup.PrintAreas.Add(sheetArea);
                        }
                    }
                }
                else if (name == "_xlnm.Print_Titles")
                {
                    LoadPrintTitles(definedName);
                }
                else
                {
                    string text = definedName.Text;

                    var comment = definedName.Comment;
                    if (localSheetId == -1)
                    {
                        if (DefinedNamesInternal.All<XLDefinedName>(nr => nr.Name != name))
                            DefinedNamesInternal.Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                    }
                    else
                    {
                        if (Worksheet(localSheetId + 1).DefinedNames.All(nr => nr.Name != name))
                            ((XLDefinedNames)Worksheet(localSheetId + 1).DefinedNames).Add(name, text, comment, validateName: false, validateRangeAddress: false).Visible = visible;
                    }
                }
            }
        }

        [GeneratedRegex(@"\A('?).*\1!.*\z")]
        private static partial Regex DefinedNameRegex { get; }

        private IEnumerable<String> validateDefinedNames(IEnumerable<String> definedNames)
        {
            var sb = new StringBuilder();
            foreach (string testName in definedNames)
            {
                if (sb.Length > 0)
                    sb.Append(',');

                sb.Append(testName);

                Match matchedValidPattern = DefinedNameRegex.Match(sb.ToString());
                if (matchedValidPattern.Success)
                {
                    yield return sb.ToString();
                    sb = new StringBuilder();
                }
            }

            if (sb.Length > 0)
                yield return sb.ToString();
        }

        private void LoadPrintTitles(DefinedName definedName)
        {
            var areas = validateDefinedNames(definedName.Text.Split(','));
            foreach (var item in areas)
            {
                if (this.Range(item) != null)
                    SetColumnsOrRowsToRepeat(item);
            }
        }

        private void SetColumnsOrRowsToRepeat(string area)
        {
            ParseReference(area, out String sheetName, out String sheetArea);
            sheetArea = sheetArea.Replace("$", "");

            if (sheetArea.Equals("#REF")) return;
            if (IsColReference(sheetArea))
                WorksheetsInternal.Worksheet(sheetName).PageSetup.SetColumnsToRepeatAtLeft(sheetArea);
            if (IsRowReference(sheetArea))
                WorksheetsInternal.Worksheet(sheetName).PageSetup.SetRowsToRepeatAtTop(sheetArea);
        }

        // either $A:$X => true or $1:$99 => false
        private static bool IsColReference(string sheetArea)
        {
            return sheetArea.All(c => c == ':' || char.IsLetter(c));
        }

        private static bool IsRowReference(string sheetArea)
        {
            return sheetArea.All(c => c == ':' || char.IsNumber(c));
        }

        private static void ParseReference(string item, out string sheetName, out string sheetArea)
        {
            var sections = item.Trim().Split('!');
            if (sections.Count() == 1)
            {
                sheetName = string.Empty;
                sheetArea = item;
            }
            else
            {
                sheetName = string.Join("!", sections.Take(sections.Length - 1)).UnescapeSheetName();
                sheetArea = sections[sections.Length - 1];
            }
        }


        private static void LoadWorkbookTheme(ThemePart? tp, XLWorkbook wb)
        {
            if (tp is null)
                return;

            var colorScheme = tp.Theme?.ThemeElements?.ColorScheme;
            if (colorScheme is not null)
            {
                var background1 = colorScheme.Light1Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(background1))
                {
                    wb.Theme.Background1 = XLColor.FromHexRgb(background1);
                }
                var text1 = colorScheme.Dark1Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(text1))
                {
                    wb.Theme.Text1 = XLColor.FromHexRgb(text1);
                }
                var background2 = colorScheme.Light2Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(background2))
                {
                    wb.Theme.Background2 = XLColor.FromHexRgb(background2);
                }
                var text2 = colorScheme.Dark2Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(text2))
                {
                    wb.Theme.Text2 = XLColor.FromHexRgb(text2);
                }
                var accent1 = colorScheme.Accent1Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent1))
                {
                    wb.Theme.Accent1 = XLColor.FromHexRgb(accent1);
                }
                var accent2 = colorScheme.Accent2Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent2))
                {
                    wb.Theme.Accent2 = XLColor.FromHexRgb(accent2);
                }
                var accent3 = colorScheme.Accent3Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent3))
                {
                    wb.Theme.Accent3 = XLColor.FromHexRgb(accent3);
                }
                var accent4 = colorScheme.Accent4Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent4))
                {
                    wb.Theme.Accent4 = XLColor.FromHexRgb(accent4);
                }
                var accent5 = colorScheme.Accent5Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent5))
                {
                    wb.Theme.Accent5 = XLColor.FromHexRgb(accent5);
                }
                var accent6 = colorScheme.Accent6Color?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(accent6))
                {
                    wb.Theme.Accent6 = XLColor.FromHexRgb(accent6);
                }
                var hyperlink = colorScheme.Hyperlink?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(hyperlink))
                {
                    wb.Theme.Hyperlink = XLColor.FromHexRgb(hyperlink);
                }
                var followedHyperlink = colorScheme.FollowedHyperlinkColor?.RgbColorModelHex?.Val?.Value;
                if (!string.IsNullOrEmpty(followedHyperlink))
                {
                    wb.Theme.FollowedHyperlink = XLColor.FromHexRgb(followedHyperlink);
                }
            }
        }

        private static void LoadWorkbookProtection(WorkbookProtection? wp, XLWorkbook wb)
        {
            if (wp == null) return;

            wb.Protection.IsProtected = true;

            var algorithmName = wp.WorkbookAlgorithmName?.Value ?? string.Empty;
            if (String.IsNullOrEmpty(algorithmName))
            {
                wb.Protection.PasswordHash = wp.WorkbookPassword?.Value ?? string.Empty;
                wb.Protection.Base64EncodedSalt = string.Empty;
            }
            else if (DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName))
            {
                wb.Protection.Algorithm = DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
                wb.Protection.PasswordHash = wp.WorkbookHashValue?.Value ?? string.Empty;
                wb.Protection.SpinCount = wp.WorkbookSpinCount?.Value ?? 0;
                wb.Protection.Base64EncodedSalt = wp.WorkbookSaltValue?.Value ?? string.Empty;
            }

            wb.Protection.AllowElement(XLWorkbookProtectionElements.Structure, !OpenXmlHelper.GetBooleanValueAsBool(wp.LockStructure, false));
            wb.Protection.AllowElement(XLWorkbookProtectionElements.Windows, !OpenXmlHelper.GetBooleanValueAsBool(wp.LockWindows, false));
        }

        private void SetProperties(SpreadsheetDocument dSpreadsheet)
        {
            var p = dSpreadsheet.PackageProperties;
            Properties.Author = p.Creator;
            Properties.Category = p.Category;
            Properties.Comments = p.Description;
            if (p.Created != null)
                Properties.Created = p.Created.Value;
            if (p.Modified != null)
                Properties.Modified = p.Modified.Value;
            Properties.Keywords = p.Keywords;
            Properties.LastModifiedBy = p.LastModifiedBy;
            Properties.Status = p.ContentStatus;
            Properties.Subject = p.Subject;
            Properties.Title = p.Title;
        }

        private XmlTreeReader CreateTreeReader(OpenXmlPart openXmlPart)
        {
            var stream = openXmlPart.GetStream(FileMode.Open);
            return new XmlTreeReader(stream, XmlToEnumMapper.Instance, StrictAttributeParsing);
        }
    }
}
