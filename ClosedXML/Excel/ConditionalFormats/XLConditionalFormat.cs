using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormat : IXLDxfContainer, IXLConditionalFormat
    {
        private readonly XLWorksheet _worksheet;
        private sealed class NoRangeCfComparer : IEqualityComparer<XLConditionalFormat>
        {
            public bool Equals(XLConditionalFormat? xx, XLConditionalFormat? yy)
            {
                if (ReferenceEquals(xx, yy)) return true;
                if (ReferenceEquals(xx, null)) return false;
                if (ReferenceEquals(yy, null)) return false;
                if (xx.GetType() != yy.GetType()) return false;

                var xxValues = xx.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
                var yyValues = yy.Values.Values.Where(v => v == null || !v.IsFormula).Select(v => v?.Value);
                var xxFormulas = xx.Areas.Count > 0 ? xx.Values.Values.OfType<XLFormula>().Where(v => v.IsFormula).Select(f => ((XLCell)xx.Range.FirstCell()).GetFormulaR1C1(f.Value)) : null;
                var yyFormulas = yy.Areas.Count > 0 ? yy.Values.Values.OfType<XLFormula>().Where(v => v.IsFormula).Select(f => ((XLCell)yy.Range.FirstCell()).GetFormulaR1C1(f.Value)) : null;
                var xStyle = xx.FormatValue;
                var yStyle = yy.FormatValue;
                return Equals(xStyle, yStyle)
                    && xx.ConditionalFormatType == yy.ConditionalFormatType
                    && xx.TimePeriod == yy.TimePeriod
                    && xx.IconSetStyle == yy.IconSetStyle
                    && xx.Operator == yy.Operator
                    && xx.Bottom == yy.Bottom
                    && xx.Percent == yy.Percent
                    && xx.ReverseIconOrder == yy.ReverseIconOrder
                    && xx.StopIfTrue == yy.StopIfTrue
                    && xx.ShowIconOnly == yy.ShowIconOnly
                    && xx.ShowBarOnly == yy.ShowBarOnly
                    && xx.Gradient == yy.Gradient
                    && xx.DataBarOptions == yy.DataBarOptions
                    && SetEquals(xxValues, yyValues)
                    && SetEquals(xxFormulas, yyFormulas)
                    && Equals(xx.Colors, yy.Colors)
                    && Equals(xx.ContentTypes, yy.ContentTypes)
                    && Equals(xx.IconSetOperators, yy.IconSetOperators);
            }

            public int GetHashCode(XLConditionalFormat obj)
            {
                var xx = obj;

                // Equal formats must have the same hash code. The hash uses only the values that Equals
                // compares directly. A dictionary contributes its count, because Equals compares its items.
                // A data bar or a color scale has no style.
                unchecked
                {
                    var hashCode = xx.FormatValue?.GetHashCode() ?? 0;
                    hashCode = (hashCode * 397) ^ xx.Colors.Count;
                    hashCode = (hashCode * 397) ^ xx.ContentTypes.Count;
                    hashCode = (hashCode * 397) ^ xx.IconSetOperators.Count;
                    hashCode = (hashCode * 397) ^ (int)xx.ConditionalFormatType;
                    hashCode = (hashCode * 397) ^ (int)xx.TimePeriod;
                    hashCode = (hashCode * 397) ^ (int)xx.IconSetStyle;
                    hashCode = (hashCode * 397) ^ (int)xx.Operator;
                    hashCode = (hashCode * 397) ^ xx.Bottom.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.Percent.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.ReverseIconOrder.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.ShowIconOnly.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.ShowBarOnly.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.Gradient.GetHashCode();
                    hashCode = (hashCode * 397) ^ xx.StopIfTrue.GetHashCode();
                    return hashCode;
                }
            }

            private static bool SetEquals<T>(IEnumerable<T>? first, IEnumerable<T>? second)
            {
                // A format without areas has no formulas to compare. It is equal only to another format
                // without areas, so formats with areas and formats without areas are not merged.
                if (first is null || second is null)
                    return first is null && second is null;

                return new HashSet<T>(second, EqualityComparer<T>.Default)
                    .SetEquals(first);
            }

            private static bool Equals<TValue>(Dictionary<int, TValue> x, Dictionary<int, TValue> y)
            {
                if (x.Count != y.Count)
                    return false;
                if (x.Keys.Except(y.Keys).Any())
                    return false;
                if (y.Keys.Except(x.Keys).Any())
                    return false;
                var valueComparer = EqualityComparer<TValue>.Default;
                foreach (var pair in x)
                    if (!valueComparer.Equals(pair.Value, y[pair.Key]))
                        return false;

                return true;
            }
        }

        internal static IEqualityComparer<XLConditionalFormat> NoRangeComparer { get; } = new NoRangeCfComparer();

        #region Constructors

        internal XLConditionalFormat(XLWorksheet worksheet, XLAreaList areaList)
        {
            _worksheet = worksheet;
            Id = Guid.NewGuid();
            Areas = areaList;
            Values = new XLDictionary<XLFormula?>();
            Colors = new XLDictionary<XLColor>();
            ContentTypes = new XLDictionary<XLCFContentType>();
            IconSetOperators = new XLDictionary<XLCFIconSetOperator>();
        }

        /// <summary>
        /// Copy ctor.
        /// </summary>
        internal XLConditionalFormat(XLWorksheet worksheet, XLConditionalFormat other, XLAreaList areaList)
            : this(worksheet, areaList)
        {
            var otherDxf = other.FormatValue;
            FormatValue = otherDxf is not null
                ? _worksheet.Workbook.Styles.GetRegisteredDxFormat(otherDxf, static x => x)
                : null;
            ConditionalFormatType = other.ConditionalFormatType;
            TimePeriod = other.TimePeriod;
            IconSetStyle = other.IconSetStyle;
            Operator = other.Operator;
            Bottom = other.Bottom;
            Percent = other.Percent;
            ReverseIconOrder = other.ReverseIconOrder;
            ShowIconOnly = other.ShowIconOnly;
            ShowBarOnly = other.ShowBarOnly;
            Gradient = other.Gradient;
            DataBarOptions = other.DataBarOptions;
            StopIfTrue = other.StopIfTrue;

            var sourceAnchor = other.Areas[0].FirstPoint;
            var targetAnchor = Areas[0].FirstPoint;
            foreach (var (key, originalValue) in other.Values)
                Values.Add(key, originalValue?.GetAdjustedCopy(sourceAnchor, targetAnchor));

            Colors = other.Colors.CopyDictionary();
            ContentTypes = other.ContentTypes.CopyDictionary();
            IconSetOperators = other.IconSetOperators.CopyDictionary();
        }

        #endregion Constructors

        public Guid Id { get; internal set; }

        /// <summary>
        /// Priority of formatting rule. Lower values have higher priority than higher values.
        /// Minimum value is 1. It is basically used for ordering of CF during saving.
        /// </summary>
        internal Int32 Priority { get; set; }

        public XLDxfValue? FormatValue { get; set; }

        internal XLDxFormat Format => new(_worksheet.Workbook.Styles, this);

        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        public XLDictionary<XLFormula?> Values { get; }

        /// <summary>
        /// Get a value of the <see cref="Values"/>. The value must not be null.
        /// </summary>
        internal XLFormula GetRequiredValue(int index)
        {
            return Values[index] ?? throw new InvalidOperationException($"The conditional format has no value at index {index}.");
        }

        public XLDictionary<XLColor> Colors { get; }

        public XLDictionary<XLCFContentType> ContentTypes { get; }

        public XLDictionary<XLCFIconSetOperator> IconSetOperators { get; }

        public IXLRange Range
        {
            get => _worksheet.Range(Areas[0]);
            set => Areas = XLAreaList.FromRange(_worksheet, value);
        }

        public IEnumerable<IXLRange> Ranges
        {
            get
            {
                var ranges = new XLRanges(_worksheet);
                foreach (var area in Areas)
                    ranges.Add(_worksheet.Range(area));

                return ranges;
            }
            set => Areas = XLAreaList.FromRanges(_worksheet, value);
        }

        public XLConditionalFormatType ConditionalFormatType { get; set; }

        public XLTimePeriod TimePeriod { get; set; }

        public XLIconSetStyle IconSetStyle { get; set; }

        public XLCFOperator Operator { get; set; }

        public Boolean Bottom { get; set; }

        public Boolean Percent { get; set; }

        public Boolean ReverseIconOrder { get; set; }

        public Boolean ShowIconOnly { get; set; }

        public Boolean ShowBarOnly { get; set; }

        public Boolean Gradient { get; set; } = true;

        /// <summary>
        /// The options of a data bar that only an Excel 2010 rule has. The record is immutable, so a copy can share it.
        /// </summary>
        internal XLCFDataBarOptions DataBarOptions { get; set; } = XLCFDataBarOptions.Default;

        public Boolean StopIfTrue { get; set; }

        internal XLAreaList Areas { get; set; }

        public IXLConditionalFormat SetStopIfTrue()
        {
            return SetStopIfTrue(true);
        }

        public IXLConditionalFormat SetStopIfTrue(bool value)
        {
            this.StopIfTrue = value;
            return this;
        }

        public IXLConditionalFormat CopyTo(IXLWorksheet targetSheet)
        {
            if (targetSheet == Range?.Worksheet)
                throw new InvalidOperationException("Cannot copy conditional format to the worksheet it already belongs to.");
            var newCf = new XLConditionalFormat((XLWorksheet)targetSheet, this, Areas);
            targetSheet.ConditionalFormats.Add(newCf);
            return newCf;
        }

        public IXLStyle WhenIsBlank()
        {
            ConditionalFormatType = XLConditionalFormatType.IsBlank;
            return Style;
        }

        public IXLStyle WhenNotBlank()
        {
            ConditionalFormatType = XLConditionalFormatType.NotBlank;
            return Style;
        }

        public IXLStyle WhenIsError()
        {
            ConditionalFormatType = XLConditionalFormatType.IsError;
            return Style;
        }

        public IXLStyle WhenNotError()
        {
            ConditionalFormatType = XLConditionalFormatType.NotError;
            return Style;
        }

        public IXLStyle WhenDateIs(XLTimePeriod timePeriod)
        {
            TimePeriod = timePeriod;
            ConditionalFormatType = XLConditionalFormatType.TimePeriod;
            return Style;
        }

        public IXLStyle WhenContains(String value)
        {
            Values.Initialize(XLFormula.Text(value));
            ConditionalFormatType = XLConditionalFormatType.ContainsText;
            Operator = XLCFOperator.Contains;
            return Style;
        }

        public IXLStyle WhenNotContains(String value)
        {
            Values.Initialize(XLFormula.Text(value));
            ConditionalFormatType = XLConditionalFormatType.NotContainsText;
            Operator = XLCFOperator.NotContains;
            return Style;
        }

        public IXLStyle WhenStartsWith(String value)
        {
            Values.Initialize(XLFormula.Text(value));
            ConditionalFormatType = XLConditionalFormatType.StartsWith;
            Operator = XLCFOperator.StartsWith;
            return Style;
        }

        public IXLStyle WhenEndsWith(String value)
        {
            Values.Initialize(XLFormula.Text(value));
            ConditionalFormatType = XLConditionalFormatType.EndsWith;
            Operator = XLCFOperator.EndsWith;
            return Style;
        }

        public IXLStyle WhenEquals(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.Equal;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotEquals(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.NotEqual;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenGreaterThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.GreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenLessThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.LessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrGreaterThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.EqualOrGreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrLessThan(String value)
        {
            Values.Initialize(new XLFormula { Value = value });
            Operator = XLCFOperator.EqualOrLessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenBetween(String minValue, String maxValue)
        {
            Values.Initialize(new XLFormula { Value = minValue });
            Values.Add(new XLFormula { Value = maxValue });
            Operator = XLCFOperator.Between;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotBetween(String minValue, String maxValue)
        {
            Values.Initialize(new XLFormula { Value = minValue });
            Values.Add(new XLFormula { Value = maxValue });
            Operator = XLCFOperator.NotBetween;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEquals(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.Equal;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotEquals(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.NotEqual;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenGreaterThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.GreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenLessThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.LessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrGreaterThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.EqualOrGreaterThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenEqualOrLessThan(Double value)
        {
            Values.Initialize(new XLFormula(value));
            Operator = XLCFOperator.EqualOrLessThan;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenBetween(Double minValue, Double maxValue)
        {
            Values.Initialize(new XLFormula(minValue));
            Values.Add(new XLFormula(maxValue));
            Operator = XLCFOperator.Between;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenNotBetween(Double minValue, Double maxValue)
        {
            Values.Initialize(new XLFormula(minValue));
            Values.Add(new XLFormula(maxValue));
            Operator = XLCFOperator.NotBetween;
            ConditionalFormatType = XLConditionalFormatType.CellIs;
            return Style;
        }

        public IXLStyle WhenIsDuplicate()
        {
            ConditionalFormatType = XLConditionalFormatType.IsDuplicate;
            return Style;
        }

        public IXLStyle WhenIsUnique()
        {
            ConditionalFormatType = XLConditionalFormatType.IsUnique;
            return Style;
        }

        public IXLStyle WhenIsTrue(String formula)
        {
            String f = formula.TrimStart()[0] == '=' ? formula : "=" + formula;
            Values.Initialize(new XLFormula { Value = f });
            ConditionalFormatType = XLConditionalFormatType.Expression;
            return Style;
        }

        public IXLStyle WhenIsTop(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items)
        {
            Values.Initialize(new XLFormula(value));
            Percent = topBottomType == XLTopBottomType.Percent;
            ConditionalFormatType = XLConditionalFormatType.Top10;
            Bottom = false;
            return Style;
        }

        public IXLStyle WhenIsBottom(Int32 value, XLTopBottomType topBottomType = XLTopBottomType.Items)
        {
            Values.Initialize(new XLFormula(value));
            Percent = topBottomType == XLTopBottomType.Percent;
            ConditionalFormatType = XLConditionalFormatType.Top10;
            Bottom = true;
            return Style;
        }

        public IXLCFColorScaleMin ColorScale()
        {
            ConditionalFormatType = XLConditionalFormatType.ColorScale;
            return new XLCFColorScaleMin(this);
        }

        public IXLCFDataBarMin DataBar(XLColor color, Boolean showBarOnly = false)
        {
            Colors.Initialize(color);
            ShowBarOnly = showBarOnly;
            Gradient = true;
            ConditionalFormatType = XLConditionalFormatType.DataBar;
            return new XLCFDataBarMin(this);
        }

        public IXLCFDataBarMin DataBar(XLColor positiveColor, XLColor negativeColor, Boolean showBarOnly = false)
        {
            Colors.Initialize(positiveColor);
            Colors.Add(negativeColor);
            ShowBarOnly = showBarOnly;
            Gradient = true;
            ConditionalFormatType = XLConditionalFormatType.DataBar;
            return new XLCFDataBarMin(this);
        }

        public IXLCFIconSet IconSet(XLIconSetStyle iconSetStyle, Boolean reverseIconOrder = false, Boolean showIconOnly = false)
        {
            IconSetOperators.Clear();
            Values.Clear();
            ContentTypes.Clear();
            ConditionalFormatType = XLConditionalFormatType.IconSet;
            IconSetStyle = iconSetStyle;
            ReverseIconOrder = reverseIconOrder;
            ShowIconOnly = showIconOnly;
            return new XLCFIconSet(this);
        }
    }
}
