using System.IO;
using ClosedXML.IO;
using ClosedXML.Utils;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class XmlEncoderTest
    {
        [Test]
        public void TestControlChars()
        {
            Assert.AreEqual("_x0001_ _x0002_ _x0003_ _x0004_", XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004"));
            Assert.AreEqual("_x0005_ _x0006_ _x0007_ _x0008_", XmlEncoder.EncodeString("\u0005 \u0006 \u0007 \u0008"));
        }

        [TestCase("")]
        [TestCase("Text with spaces, tabs\t and new lines\n")]
        [TestCase("\u65E5\u672C\u8A9E \u00E9 \uFFFD")]
        [Issue("74")]
        public void Text_without_special_chars_is_the_same_instance(string text)
        {
            Assert.AreSame(text, XmlEncoder.EncodeString(text));
        }

        [TestCase("_x0041_", "_x005F_x0041_")]
        [TestCase("a_b", "a_b")]
        [TestCase("a\u000Bb", "a_x000B_b")]
        [Issue("74")]
        public void Text_with_special_chars_is_encoded(string text, string expected)
        {
            Assert.AreEqual(expected, XmlEncoder.EncodeString(text));
        }

        [TestCase("a\rb", "a_x000D_b")]
        [TestCase("a\r\nb", "a_x000D_\nb")]
        [TestCase("_x005F_x000D_", "_x005F_x005F_x005F_x000D_")]
        [TestCase("_x00aF_", "_x005F_x00aF_")]
        [TestCase("_x0041\r", "_x005F_x0041_x000D_")]
        [TestCase("_x0041\u0001", "_x005F_x0041_x0001_")]
        [TestCase("_x0041a", "_x0041a")]
        [TestCase("_x0041\U0001F600", "_x0041\U0001F600")]
        [TestCase("_X0041_", "_X0041_")]
        [TestCase("_x00G1_", "_x00G1_")]
        [TestCase("_x0041", "_x0041")]
        [Issue("96")]
        public void Encoded_text_has_the_same_decoded_text(string text, string expected)
        {
            var encoded = XmlEncoder.EncodeString(text);

            Assert.AreEqual(expected, encoded);
            Assert.AreEqual(text, XStringConvert.Decode(encoded));
        }

        [Test]
        [Issue("96")]
        public void Escape_text_before_a_surrogate_without_a_pair_is_escaped()
        {
            // An attribute cannot hold a surrogate without a pair, so the test makes the text.
            var text = "_x0041" + (char)0xD800;

            var encoded = XmlEncoder.EncodeString(text);

            Assert.AreEqual("_x005F_x0041_xD800_", encoded);
            Assert.AreEqual(text, XStringConvert.Decode(encoded));
        }

        [Test]
        [Issue("74")]
        public void Surrogates_and_non_characters_are_encoded()
        {
            // An attribute can't have a string with an unpaired surrogate, so the strings are in the code.
            Assert.AreEqual("_xD800_", XmlEncoder.EncodeString("\uD800"));
            Assert.AreEqual("_xDC00_a", XmlEncoder.EncodeString("\uDC00a"));
            Assert.AreEqual("_xFFFE__xFFFF_", XmlEncoder.EncodeString("\uFFFE\uFFFF"));

            var pair = "\uD83D\uDE00";
            Assert.AreEqual(pair, XmlEncoder.EncodeString(pair));
        }

        [Test]
        public void AstralUnicodeCharsAreWrittenWithoutOpenXmlEncoding()
        {
            using var sr = new StreamReader(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Unicode\let_it_go_in_emoji.txt")));
            var surrogateEmoji = sr.ReadToEnd();

            TestHelper.CreateAndCompare(wb =>
            {
                var ws = wb.AddWorksheet();

                var cell = ws.FirstCell();
                cell.Value = "This emoji version of Let It Go from Frozen:";
                cell.CellBelow().Value = surrogateEmoji;
            }, @"Other\Unicode\let_it_go_in_emoji-outputfile.xlsx");
        }
    }
}
