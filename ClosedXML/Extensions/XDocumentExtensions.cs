// Keep this file CodeMaid organised and cleaned
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ClosedXML.Excel
{
    internal static class XDocumentExtensions
    {
        public static XDocument? Load(Stream stream)
        {
            using (XmlReader reader = XmlReader.Create(stream))
            {
                try
                {
                    return XDocument.Load(reader);
                }
                catch (XmlException)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Same as <see cref="XNode.ToString()"/> (indented, no XML declaration), but the line ends are always CRLF.
        /// <c>XNode.ToString()</c> uses the line ends of the platform, so the output would differ between Windows and Linux.
        /// </summary>
        public static string ToStringCrLf(this XNode node)
        {
            var settings = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                Indent = true,
                NewLineChars = "\r\n",
                ConformanceLevel = ConformanceLevel.Auto
            };
            var sb = new StringBuilder();
            using (var writer = XmlWriter.Create(sb, settings))
            {
                node.WriteTo(writer);
            }
            return sb.ToString();
        }
    }
}
