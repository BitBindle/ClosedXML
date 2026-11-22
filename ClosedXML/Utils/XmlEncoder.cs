using System;
using System.Buffers;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ClosedXML.Utils
{
    internal static partial class XmlEncoder
    {
        [GeneratedRegex("_(x[\\dA-Fa-f]{4})_")]
        private static partial Regex XHHHHRegex { get; }

        /// <summary>
        /// The characters that the encoder changes: the start of an escape sequence, the characters
        /// that XML doesn't allow and the surrogates. The other characters stay the same.
        /// </summary>
        private static readonly SearchValues<char> SpecialChars = SearchValues.Create(GetSpecialChars());

        public static string EncodeString(string encodeStr)
        {
            // Most strings have no special characters. The fast path doesn't allocate memory.
            if (encodeStr.AsSpan().IndexOfAny(SpecialChars) < 0)
                return encodeStr;

            encodeStr = XHHHHRegex.Replace(encodeStr, "_x005F_$1_");

            var sb = new StringBuilder(encodeStr.Length);
            var len = encodeStr.Length;
            for (var i = 0; i < len; ++i)
            {
                var currentChar = encodeStr[i];
                if (XmlConvert.IsXmlChar(currentChar))
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
                    sb.Append(XmlConvert.EncodeName(currentChar.ToString()));
                }
            }

            return sb.ToString();
        }

        private static char[] GetSpecialChars()
        {
            var chars = new List<char> { '_' };
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
