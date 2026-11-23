using System;
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace ClosedXML.Utils
{
    internal static class XmlEncoder
    {
        /// <summary>
        /// The characters that the encoder can change: the start of an escape sequence, the carriage
        /// return, the characters that XML doesn't allow and the surrogates. The other characters
        /// stay the same.
        /// </summary>
        private static readonly SearchValues<char> SpecialChars = SearchValues.Create(GetSpecialChars());

        /// <summary>
        /// The special characters and the line feed.
        /// </summary>
        private static readonly SearchValues<char> SpecialCharsAndLineFeed = SearchValues.Create([.. GetSpecialChars(), '\n']);

        /// <summary>
        /// Encodes the text as an <c>ST_Xstring</c>.
        /// </summary>
        /// <param name="encodeStr">The text to encode.</param>
        /// <param name="escapeLineFeed">
        /// If <c>true</c>, the method also writes a line feed as an escape sequence. Excel does this for a
        /// text in an attribute, for example the name of a table column.
        /// </param>
        public static string EncodeString(string encodeStr, bool escapeLineFeed = false)
        {
            // Most strings have no special characters. The fast path doesn't allocate memory.
            var specialChars = escapeLineFeed ? SpecialCharsAndLineFeed : SpecialChars;
            if (encodeStr.AsSpan().IndexOfAny(specialChars) < 0)
                return encodeStr;

            var sb = new StringBuilder(encodeStr.Length + 16);
            var len = encodeStr.Length;
            for (var i = 0; i < len; ++i)
            {
                var currentChar = encodeStr[i];
                if (currentChar == '_')
                {
                    // An escaped underscore keeps the decoder from reading an escape sequence.
                    if (StartsEscapeSequence(encodeStr, i, escapeLineFeed))
                        sb.Append("_x005F_");
                    else
                        sb.Append('_');
                }
                else if (currentChar == '\r' || (escapeLineFeed && currentChar == '\n'))
                {
                    // An XML reader changes a carriage return to a line feed. Excel writes a line feed in
                    // an attribute as an escape sequence, because an XML reader changes it to a space.
                    AppendEscapeSequence(sb, currentChar);
                }
                else if (XmlConvert.IsXmlChar(currentChar))
                {
                    sb.Append(currentChar);
                }
                else if (i + 1 < len && XmlConvert.IsXmlSurrogatePair(encodeStr[i + 1], currentChar))
                {
                    sb.Append(currentChar);
                    sb.Append(encodeStr[++i]);
                }
                else
                {
                    AppendEscapeSequence(sb, currentChar);
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Does an escape sequence (<c>_xHHHH_</c>) start at the <paramref name="index"/> of the
        /// text after the encoding? The last underscore of the sequence can also be the first
        /// character of an escape sequence that the encoder writes for a character.
        /// </summary>
        private static bool StartsEscapeSequence(string text, int index, bool escapeLineFeed)
        {
            if (index + 6 >= text.Length || text[index + 1] != 'x')
                return false;

            for (var i = index + 2; i < index + 6; ++i)
            {
                if (!char.IsAsciiHexDigit(text[i]))
                    return false;
            }

            var lastChar = text[index + 6];
            if (lastChar == '_' || lastChar == '\r' || (escapeLineFeed && lastChar == '\n'))
                return true;

            // The encoder keeps a valid surrogate pair.
            var isSurrogatePair = index + 7 < text.Length && XmlConvert.IsXmlSurrogatePair(text[index + 7], lastChar);
            return !isSurrogatePair && !XmlConvert.IsXmlChar(lastChar);
        }

        private static void AppendEscapeSequence(StringBuilder sb, char c)
        {
            sb.Append(CultureInfo.InvariantCulture, $"_x{(int)c:X4}_");
        }

        private static char[] GetSpecialChars()
        {
            var chars = new List<char> { '_', '\r' };
            for (var c = char.MinValue; c < char.MaxValue; ++c)
            {
                if (!XmlConvert.IsXmlChar(c))
                    chars.Add(c);
            }

            // The loop doesn't test the last character, U+FFFF, which XML doesn't allow.
            chars.Add(char.MaxValue);
            return chars.ToArray();
        }
    }
}
