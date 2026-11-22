using System;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ClosedXML.Excel;
using ClosedXML.Extensions;
using NUnit.Framework;

namespace ClosedXML.Tests.Extensions
{
    [TestFixture]
    public class XmlWriterExtensionsTests
    {
        [TestCase(0)]
        [TestCase(7)]
        [TestCase(-42)]
        [TestCase(int.MinValue)]
        [TestCase(int.MaxValue)]
        [Issue("74")]
        public void Int_text_is_same_as_text_of_WriteValue(int value)
        {
            var expected = Write(w => w.WriteValue(value));

            Assert.AreEqual(expected, Write(w => w.WriteNumberValue(value)));
        }

        [TestCase(0u)]
        [TestCase(7u)]
        [TestCase(uint.MaxValue)]
        [Issue("74")]
        public void UInt_text_is_same_as_text_of_WriteValue(uint value)
        {
            var expected = Write(w => w.WriteValue((long)value));

            Assert.AreEqual(expected, Write(w => w.WriteNumberValue(value)));
        }

        [TestCase(0d)]
        [TestCase(0.1)]
        [TestCase(-123.45)]
        [TestCase(1234567890123456d)]
        [TestCase(double.MaxValue)]
        [TestCase(double.Epsilon)]
        [TestCase(double.NaN)]
        [TestCase(double.NegativeInfinity)]
        [Issue("74")]
        public void Double_text_is_round_trip_text(double value)
        {
            var expected = $"<a>{ObjectExtensions.ToRoundTripString(value)}</a>";

            Assert.AreEqual(expected, Write(w => w.WriteNumberValue(value)));
        }

        [Test]
        [Issue("74")]
        public void Number_attributes_have_the_number_text()
        {
            var xml = Write(w =>
            {
                w.WriteAttribute("i", -5);
                w.WriteAttribute("u", 5u);
                w.WriteAttribute("d", 0.1 + 0.2);
            });

            Assert.AreEqual("<a i=\"-5\" u=\"5\" d=\"0.30000000000000004\" />", xml);
        }

        [Test]
        [Issue("74")]
        public void Threads_write_numbers_at_the_same_time()
        {
            // Each thread has its own number buffer, so the text of one thread doesn't change
            // the text of another thread.
            Parallel.For(0, 8, thread =>
            {
                for (var i = 0; i < 1000; i++)
                {
                    var value = thread * 1_000_000 + i;
                    var xml = Write(w =>
                    {
                        w.WriteNumberValue(value);
                        w.WriteString(" ");
                        w.WriteNumberValue(value + 0.5);
                    });
                    var expected = string.Create(CultureInfo.InvariantCulture, $"<a>{value} {value + 0.5}</a>");
                    if (xml != expected)
                        Assert.Fail($"Expected {expected}, but was {xml}.");
                }
            });
        }

        private static string Write(Action<XmlWriter> write)
        {
            var sb = new StringBuilder();
            var settings = new XmlWriterSettings { OmitXmlDeclaration = true };
            using (var w = XmlWriter.Create(sb, settings))
            {
                w.WriteStartElement("a");
                write(w);
                w.WriteEndElement();
            }

            return sb.ToString();
        }
    }
}
