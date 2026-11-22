using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel
{
    internal class XLFormattedText<T> : IXLFormattedText<T>
    {
        protected readonly XLWorkbookStyles Styles;

        /// <summary>
        /// Font used for a new rich text run, never modified. It is generally provided by a container of the formatted text.
        /// </summary>
        private readonly XLFontFormatValue _defaultFont;
        private readonly List<XLRichString> _richTexts = new();
        private XLPhonetics? _phonetics;

        protected XLFormattedText(XLFontFormatValue defaultFont, XLWorkbookStyles styles)
        {
            Debug.Assert(styles.Fonts.ContainsValue(defaultFont));
            _defaultFont = defaultFont;
            Styles = styles;
        }

        IXLPhonetics IXLFormattedText<T>.Phonetics => Phonetics;

        public Int32 Count => _richTexts.Count;

        public int Length { get; private set; }

        public String Text => ToString();

        public Boolean HasPhonetics => _phonetics is not null;

        /// <inheritdoc cref="IXLFormattedText{T}.Phonetics"/>
        internal XLPhonetics Phonetics
        {
            get => _phonetics ??= new XLPhonetics(_defaultFont, _defaultFont, Styles, OnContentChanged);
            init => _phonetics = value;
        } 

        public IXLRichString AddText(String text)
        {
            var richText = new XLRichString(text, _defaultFont, this, Styles, OnContentChanged);
            AddText(richText);
            OnContentChanged();
            return richText;
        }

        public IXLRichString AddText(String text, IXLFontBase font)
        {
            var richFont = XLFontFormatValue.FromFontBase(font, Styles);
            var richText = new XLRichString(text, richFont, this, Styles, OnContentChanged);
            AddText(richText);
            OnContentChanged();
            return richText;
        }

        public IXLRichString AddNewLine()
        {
            return AddText(Environment.NewLine);
        }

        public IXLFormattedText<T> ClearText()
        {
            ClearContent();
            OnContentChanged();
            return this;
        }

        public IXLFormattedText<T> ClearFont()
        {
            String text = Text;
            ClearContent();
            AddText(text);
            return this;
        }

        public override string ToString()
        {
            var sb = new StringBuilder(_richTexts.Count);
            _richTexts.ForEach(rt => sb.Append(rt.Text));
            return sb.ToString();
        }

        public IXLFormattedText<T> Substring(Int32 index)
        {
            return Substring(index, Length - index);
        }

        public IXLFormattedText<T> Substring(Int32 index, Int32 length)
        {
            if (index + 1 > Length || (Length - index + 1) < length || length <= 0)
                throw new IndexOutOfRangeException("Index and length must refer to a location within the string.");

            var newRichTexts = new List<XLRichString>();
            var retVal = new XLFormattedText<T>(_defaultFont, Styles);

            var lastPosition = 0;
            foreach (var rt in _richTexts)
            {
                if (lastPosition >= index + 1 + length) // We already have what we need
                {
                    newRichTexts.Add(rt);
                }
                else if (lastPosition + rt.Text.Length >= index + 1) // Eureka!
                {
                    var startIndex = index - lastPosition;

                    if (startIndex > 0)
                        newRichTexts.Add(new XLRichString(rt.Text[..startIndex], rt.Font, this, Styles, OnContentChanged));
                    else if (startIndex < 0)
                        startIndex = 0;

                    var leftToTake = length - retVal.Length;
                    if (leftToTake > rt.Text.Length - startIndex)
                        leftToTake = rt.Text.Length - startIndex;

                    var newRt = new XLRichString(rt.Text.Substring(startIndex, leftToTake), rt.Font, this, Styles, OnContentChanged);
                    newRichTexts.Add(newRt);
                    retVal.AddText(newRt);

                    if (startIndex + leftToTake < rt.Text.Length)
                        newRichTexts.Add(new XLRichString(rt.Text.Substring(startIndex + leftToTake), rt.Font, this, Styles, OnContentChanged));
                }
                else // We haven't reached the desired position yet
                {
                    newRichTexts.Add(rt);
                }
                lastPosition += rt.Text.Length;
            }

            _richTexts.Clear();
            _richTexts.AddRange(newRichTexts);
            OnContentChanged();
            return retVal;
        }

        public IXLFormattedText<T> CopyFrom(IXLFormattedText<T> original)
        {
            ClearContent();
            foreach (var richText in original)
            {
                var copyFont = XLFontFormatValue.FromFontBase(richText, Styles);
                var copyText = new XLRichString(richText.Text, copyFont, this, Styles, OnContentChanged);
                AddText(copyText);
            }

            OnContentChanged();
            return this;
        }

        public List<XLRichString>.Enumerator GetEnumerator() => _richTexts.GetEnumerator();

        IEnumerator<IXLRichString> IEnumerable<IXLRichString>.GetEnumerator() => GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

        public Boolean Bold { set { _richTexts.ForEach(rt => rt.Bold = value); } }
        public Boolean Italic { set { _richTexts.ForEach(rt => rt.Italic = value); } }
        public XLFontUnderlineValues Underline { set { _richTexts.ForEach(rt => rt.Underline = value); } }
        public Boolean Strikethrough { set { _richTexts.ForEach(rt => rt.Strikethrough = value); } }
        public XLFontVerticalTextAlignmentValues VerticalAlignment { set { _richTexts.ForEach(rt => rt.VerticalAlignment = value); } }
        public Boolean Shadow { set { _richTexts.ForEach(rt => rt.Shadow = value); } }
        public Double FontSize { set { _richTexts.ForEach(rt => rt.FontSize = value); } }
        public XLColor FontColor { set { _richTexts.ForEach(rt => rt.FontColor = value); } }
        public String FontName { set { _richTexts.ForEach(rt => rt.FontName = value); } }
        public XLFontFamilyNumberingValues FontFamilyNumbering { set { _richTexts.ForEach(rt => rt.FontFamilyNumbering = value); } }

        public IXLFormattedText<T> SetBold() { Bold = true; return this; }
        public IXLFormattedText<T> SetBold(Boolean value) { Bold = value; return this; }
        public IXLFormattedText<T> SetItalic() { Italic = true; return this; }
        public IXLFormattedText<T> SetItalic(Boolean value) { Italic = value; return this; }
        public IXLFormattedText<T> SetUnderline() { Underline = XLFontUnderlineValues.Single; return this; }
        public IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value) { Underline = value; return this; }
        public IXLFormattedText<T> SetStrikethrough() { Strikethrough = true; return this; }
        public IXLFormattedText<T> SetStrikethrough(Boolean value) { Strikethrough = value; return this; }
        public IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return this; }
        public IXLFormattedText<T> SetShadow() { Shadow = true; return this; }
        public IXLFormattedText<T> SetShadow(Boolean value) { Shadow = value; return this; }
        public IXLFormattedText<T> SetFontSize(Double value) { FontSize = value; return this; }
        public IXLFormattedText<T> SetFontColor(XLColor value) { FontColor = value; return this; }
        public IXLFormattedText<T> SetFontName(String value) { FontName = value; return this; }
        public IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value) { FontFamilyNumbering = value; return this; }

        public bool Equals(IXLFormattedText<T>? other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            if (Count != other.Count)
                return false;

            if (!_richTexts.SequenceEqual(other))
                return false;

            return (_phonetics is null && !other.HasPhonetics) || Phonetics.Equals(other.Phonetics);
        }

        protected void AddText(XLRichString richText)
        {
            _richTexts.Add(richText);
            Length += richText.Text.Length;
        }

        /// <summary>
        /// This method is called every time the formatted text is changed (new runs, font props, phonetics...).
        /// </summary>
        protected virtual void OnContentChanged()
        {
            // Do nothing, intended to be overriden.
        }

        private void ClearContent()
        {
            _richTexts.Clear();
            Length = 0;
        }
    }
}
