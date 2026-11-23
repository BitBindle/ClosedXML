using System;
using System.Diagnostics;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An API object to modify a rich string.
    /// </summary>
    [DebuggerDisplay("{Text}")]
    internal class XLRichString : IXLRichString
    {
        private readonly XLWorkbookStyles _styles;
        private readonly IXLWithRichString _withRichString;
        private readonly Action _onChange;
        private string _text;
        private XLFontFormatValue _font;

        internal XLRichString(String text, XLFontFormatValue font, IXLWithRichString withRichString, XLWorkbookStyles styles, Action? onChange)
        {
            _text = text;
            _font = font;
            _withRichString = withRichString;
            _styles = styles;
            _onChange = onChange ?? (() => { });
        }

        public String Text
        {
            get => _text;
            set
            {
                _text = value;
                _onChange();
            }
        }

        public IXLRichString AddText(String text)
        {
            return _withRichString.AddText(text);
        }

        public IXLRichString AddNewLine()
        {
            return AddText("\n");
        }

        public Boolean Bold
        {
            get => _font.Bold;
            set => ChangeFont(f => f with { Bold = value });
        }

        public Boolean Italic
        {
            get => _font.Italic;
            set => ChangeFont(f => f with { Italic = value });
        }

        public XLFontUnderlineValues Underline
        {
            get => _font.Underline;
            set => ChangeFont(f => f with { Underline = value });
        }

        public Boolean Strikethrough
        {
            get => _font.Strikethrough;
            set => ChangeFont(f => f with { Strikethrough = value });
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get => _font.VerticalAlignment;
            set => ChangeFont(f => f with { VerticalAlignment = value });
        }

        public Boolean Shadow
        {
            get => _font.Shadow;
            set => ChangeFont(f => f with { Shadow = value });
        }

        public Double FontSize
        {
            get => _font.Size.Points;
            set => ChangeFont(f => f with { Size = XLFontSize.FromPoints(value) });
        }

        public XLColor FontColor
        {
            get => _font.Color;
            set => ChangeFont(f => f with { Color = value });
        }

        public String FontName
        {
            get => _font.Name.Text;
            set => ChangeFont(f => f with { Name = value });
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get => _font.Family;
            set => ChangeFont(f => f with { Family = value });
        }

        public XLFontCharSet FontCharSet
        {
            get => _font.Charset;
            set => ChangeFont(f => f with { Charset = value });
        }

        public XLFontScheme FontScheme
        {
            get => _font.Scheme;
            set => ChangeFont(f => f with { Scheme = value });
        }

        internal XLFontFormatValue Font => _font;

        public IXLRichString SetBold()
        {
            Bold = true; return this;
        }

        public IXLRichString SetBold(Boolean value)
        {
            Bold = value; return this;
        }

        public IXLRichString SetItalic()
        {
            Italic = true; return this;
        }

        public IXLRichString SetItalic(Boolean value)
        {
            Italic = value; return this;
        }

        public IXLRichString SetUnderline()
        {
            Underline = XLFontUnderlineValues.Single; return this;
        }

        public IXLRichString SetUnderline(XLFontUnderlineValues value)
        {
            Underline = value; return this;
        }

        public IXLRichString SetStrikethrough()
        {
            Strikethrough = true; return this;
        }

        public IXLRichString SetStrikethrough(Boolean value)
        {
            Strikethrough = value; return this;
        }

        public IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        {
            VerticalAlignment = value; return this;
        }

        public IXLRichString SetShadow()
        {
            Shadow = true; return this;
        }

        public IXLRichString SetShadow(Boolean value)
        {
            Shadow = value; return this;
        }

        public IXLRichString SetFontSize(Double value)
        {
            FontSize = value; return this;
        }

        public IXLRichString SetFontColor(XLColor value)
        {
            FontColor = value; return this;
        }

        public IXLRichString SetFontName(String value)
        {
            FontName = value; return this;
        }

        public IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        {
            FontFamilyNumbering = value; return this;
        }

        public IXLRichString SetFontCharSet(XLFontCharSet value)
        {
            FontCharSet = value; return this;
        }

        public IXLRichString SetFontScheme(XLFontScheme value)
        {
            FontScheme = value; return this;
        }

        public override bool Equals(object? obj) => Equals(obj as XLRichString);

        public Boolean Equals(IXLRichString? other) => Equals(other as XLRichString);

        public Boolean Equals(XLRichString? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            return Text == other.Text && _font.Equals(other._font);
        }

        public override int GetHashCode()
        {
            // Since all properties of type are mutable, can't have different hashcode for any instance.
            // Don't ever use this class in a dictionary, e.g. SST.
            return 4; // Chosen by fair dice roll. Guaranteed to be random.
        }

        private void ChangeFont(Func<XLFontFormatValue, XLFontFormatValue> modifyFont)
        {
            _font = _styles.GetRegisteredFontFormat(_font, modifyFont);
            _onChange();
        }
    }
}
