using ClosedXML.Extensions;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Office.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCFDataBarConverterExtension : IXLCFConverterExtension
    {
        private static readonly IReadOnlyDictionary<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues, ConditionalFormattingValueObjectTypeValues> CFValueToTypeMap =
            new Dictionary<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues, ConditionalFormattingValueObjectTypeValues>
            {
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Max, ConditionalFormattingValueObjectTypeValues.AutoMax },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Min, ConditionalFormattingValueObjectTypeValues.AutoMin },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Number, ConditionalFormattingValueObjectTypeValues.Numeric },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Percent, ConditionalFormattingValueObjectTypeValues.Percent },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Percentile, ConditionalFormattingValueObjectTypeValues.Percentile },
                { DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues.Formula, ConditionalFormattingValueObjectTypeValues.Formula },
            };

        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context)
        {
            ConditionalFormattingRule conditionalFormattingRule = new ConditionalFormattingRule()
            {
                Type = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValues.DataBar,
                Id = ((XLConditionalFormat)cf).Id.WrapInBraces()
            };

            var options = ((XLConditionalFormat)cf).DataBarOptions;
            DataBar dataBar = new DataBar()
            {
                MinLength = options.MinLength,
                MaxLength = options.MaxLength,
                Gradient = cf.Gradient,
                ShowValue = !cf.ShowBarOnly,
                Border = OpenXmlHelper.GetBooleanValue(options.Border, false),
                NegativeBarColorSameAsPositive = OpenXmlHelper.GetBooleanValue(options.NegativeBarColorSameAsPositive, false),
                NegativeBarBorderColorSameAsPositive = OpenXmlHelper.GetBooleanValue(options.NegativeBarBorderColorSameAsPositive, true),
            };
            if (options.Direction != XLCFDataBarDirection.Context)
                dataBar.Direction = options.Direction.ToOpenXml();
            if (options.AxisPosition != XLCFDataBarAxisPosition.Automatic)
                dataBar.AxisPosition = options.AxisPosition.ToOpenXml();

            var cfMinType = cf.ContentTypes.TryGetValue(1, out var contentType1)
                ? GetCFType(contentType1.ToOpenXml())
                : ConditionalFormattingValueObjectTypeValues.AutoMin;
            var cfMin = new ConditionalFormattingValueObject { Type = cfMinType };
            if (cf.Values.Any() && cf.Values[1]?.Value is { } minValue)
            {
                cfMin.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
                cfMin.Append(new Formula() { Text = minValue });
            }

            var cfMaxType = cf.ContentTypes.TryGetValue(2, out var contentType2)
                ? GetCFType(contentType2.ToOpenXml())
                : ConditionalFormattingValueObjectTypeValues.AutoMax;
            var cfMax = new ConditionalFormattingValueObject { Type = cfMaxType };
            if (cf.Values.Count >= 2 && cf.Values[2]?.Value is { } maxValue)
            {
                cfMax.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
                cfMax.Append(new Formula() { Text = maxValue });
            }

            var negativeFillColor = new NegativeFillColor { Rgb = cf.Colors[1].Color.ToHex() };
            if (cf.Colors.Count == 2)
            {
                negativeFillColor = new NegativeFillColor { Rgb = cf.Colors[2].Color.ToHex() };
            }

            // The schema gives the order of the elements.
            dataBar.Append(cfMin);
            dataBar.Append(cfMax);
            if (options.BorderColor is not null)
                dataBar.Append(new BorderColor().FromClosedXMLColor<BorderColor>(options.BorderColor));

            dataBar.Append(negativeFillColor);
            if (options.NegativeBorderColor is not null)
                dataBar.Append(new NegativeBorderColor().FromClosedXMLColor<NegativeBorderColor>(options.NegativeBorderColor));

            if (options.AxisColor is not null)
                dataBar.Append(new BarAxisColor().FromClosedXMLColor<BarAxisColor>(options.AxisColor));

            conditionalFormattingRule.Append(dataBar);

            return conditionalFormattingRule;
        }

        private static ConditionalFormattingValueObjectTypeValues GetCFType(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObjectValues value)
        {
            return CFValueToTypeMap[value];
        }
    }
}
