using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;

namespace ClosedXML.Excel.IO
{
    internal class CalculationChainPartWriter
    {
        internal static void GenerateContent(WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
        {
            var calculationChainPart = workbookPart.CalculationChainPart;
            ExtensionList? extensionList = null;
            var namespaceDeclarations = new List<(string Prefix, string NamespaceUri)>();
            if (calculationChainPart is not null)
            {
                // A delete of a worksheet can load the DOM of the part. The package writes a loaded
                // DOM to the part on save, so the DOM would replace the content from this writer.
                calculationChainPart.UnloadRootElement();
                extensionList = LoadExtensionList(calculationChainPart, namespaceDeclarations);
            }
            else
            {
                calculationChainPart = workbookPart.AddNewPart<CalculationChainPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
            }

            var settings = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = XLHelper.NoBomUTF8
            };

            var cellCount = 0;
            using (var xml = XmlWriter.Create(calculationChainPart.GetStream(FileMode.Create), settings))
            {
                xml.WriteStartDocument();
                xml.WriteStartElement("x", "calcChain", Main2006SsNs);

                // The extensions can use the prefixes of the old root element in attribute values.
                foreach (var (prefix, namespaceUri) in namespaceDeclarations)
                    xml.WriteAttributeString("xmlns", prefix, XmlnsNs, namespaceUri);

                // The longest reference is XFD1048576.
                var reference = new char[10];
                foreach (var worksheet in workbook.WorksheetsInternal)
                {
                    var sheetId = (int)worksheet.SheetId;
                    foreach (var c in worksheet.Internals.CellsCollection.GetCells())
                    {
                        if (c.Formula is not { } formula)
                            continue;

                        if (formula.Type == FormulaType.DataTable)
                        {
                            // Do nothing, Excel doesn't generate calc chain for data table
                        }
                        else if (c.HasArrayFormula)
                        {
                            if (formula.Range == default)
                                formula.Range = new Area(c.Point);

                            var formulaRange = formula.Range;
                            if (formulaRange.FirstPoint == c.Point)
                            {
                                WriteCell(xml, reference, c.Point, sheetId, isArray: true);
                                cellCount++;

                                // The first cell of the array is in the chain two times, the second
                                // time without the array flag.
                                foreach (var point in formulaRange)
                                {
                                    WriteCell(xml, reference, point, sheetId, isArray: false);
                                    cellCount++;
                                }
                            }
                        }
                        else
                        {
                            WriteCell(xml, reference, c.Point, sheetId, isArray: false);
                            cellCount++;
                        }
                    }
                }

                extensionList?.WriteTo(xml);

                xml.WriteEndElement(); // calcChain
            }

            // The schema requires at least one cell in the chain.
            if (cellCount == 0)
                workbookPart.DeletePart(calculationChainPart);
        }

        private static void WriteCell(XmlWriter xml, char[] reference, Point point, int sheetId, bool isArray)
        {
            xml.WriteStartElement("c", Main2006SsNs);

            var length = point.Format(reference);
            xml.WriteStartAttribute("r");
            xml.WriteRaw(reference, 0, length);
            xml.WriteEndAttribute();

            xml.WriteAttribute("i", sheetId);

            if (isArray)
                xml.WriteAttribute("a", true);

            xml.WriteEndElement(); // c
        }

        /// <summary>
        /// Load the extension list of the part, but not the cells. The writer writes the cells from
        /// the workbook.
        /// </summary>
        /// <param name="calculationChainPart">The old part.</param>
        /// <param name="namespaceDeclarations">The method adds the prefixed namespace declarations
        /// of the root element to this list. It doesn't add the prefix that the writer uses.</param>
        private static ExtensionList? LoadExtensionList(CalculationChainPart calculationChainPart, List<(string Prefix, string NamespaceUri)> namespaceDeclarations)
        {
            using var stream = calculationChainPart.GetStream(FileMode.Open, FileAccess.Read);

            // The SDK reads a part with fewer than 4 bytes as an empty part.
            if (stream.CanSeek && stream.Length < 4)
                return null;

            // The OpenXmlReader makes an element for each cell, so use the XmlReader.
            using var reader = XmlReader.Create(stream);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                if (reader.Depth == 0)
                {
                    ReadNamespaceDeclarations(reader, namespaceDeclarations);
                }
                else if (reader.Depth == 1 &&
                         reader.LocalName == "extLst" &&
                         reader.NamespaceURI == Main2006SsNs)
                {
                    return new ExtensionList(reader.ReadOuterXml());
                }
            }

            return null;
        }

        private static void ReadNamespaceDeclarations(XmlReader reader, List<(string Prefix, string NamespaceUri)> namespaceDeclarations)
        {
            if (!reader.MoveToFirstAttribute())
                return;

            do
            {
                // The writer uses the prefix x for the main namespace, and doesn't need the default namespace.
                if (reader.NamespaceURI == XmlnsNs && reader.Prefix == "xmlns" && reader.LocalName != "x")
                    namespaceDeclarations.Add((reader.LocalName, reader.Value));
            } while (reader.MoveToNextAttribute());

            reader.MoveToElement();
        }
    }
}
