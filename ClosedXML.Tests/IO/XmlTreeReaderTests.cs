using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;

namespace ClosedXML.Tests.IO;

[TestFixture]
internal class XmlTreeReaderTests
{
    [Test]
    public void Can_transparently_processes_MCE()
    {
        const string xml = $"""
                            <font xmlns="{OpenXmlConst.Main2006SsNs}"
                                   xmlns:mc="{OpenXmlConst.MarkupCompatibilityNs}">
                              <mc:AlternateContent>
                                <mc:Choice xmlns:cs="http://example.com/custom" Requires="cs">
                                  <cs:bold weight="10"/>
                                </mc:Choice>
                                <mc:Fallback>
                                  <b/>
                                </mc:Fallback>
                              </mc:AlternateContent>
                            </font>
                            """;
        using var reader = CreateReader(xml);
        reader.Open("font", OpenXmlConst.Main2006SsNs);
        reader.Open("b", OpenXmlConst.Main2006SsNs);
        reader.Close("b", OpenXmlConst.Main2006SsNs);
        reader.Close("font", OpenXmlConst.Main2006SsNs);
    }

    [Test]
    public void GetContent_reads_xml_in_element()
    {
        const string xml =
            """
            <root>
              Hello <![CDATA[world]]>
              ! 
            </root>
            """;
        using var reader = CreateReader(xml);
        reader.Open("root", string.Empty);
        var content = reader.GetContent();
        reader.Close("root", string.Empty);

        Assert.That(content, Is.EqualTo("\n  Hello world\n  ! \n"));
    }

    [Test]
    public void LeaveOpen_true_does_not_close_wrapped_XmlReader()
    {
        using var xmlReader = CreateXmlReader("<root/>");
        using (new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance, true, leaveOpen: true))
        {
        }

        Assert.That(xmlReader.ReadState, Is.Not.EqualTo(ReadState.Closed));
    }

    [Test]
    public void LeaveOpen_false_closes_wrapped_XmlReader()
    {
        var xmlReader = CreateXmlReader("<root/>");
        using (new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance, true, leaveOpen: false))
        {
        }

        Assert.That(xmlReader.ReadState, Is.EqualTo(ReadState.Closed));
    }

    [Test]
    public void TryOpen_opens_current_element_of_already_positioned_XmlReader()
    {
        using var xmlReader = CreateXmlReader("""
                                              <root>
                                                <child/>
                                              </root>
                                              """);
        xmlReader.Read();
        xmlReader.Read();
        Assert.That(xmlReader.NodeType, Is.EqualTo(XmlNodeType.Element));
        Assert.That(xmlReader.LocalName, Is.EqualTo("child"));

        using var treeReader = new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance);

        Assert.That(treeReader.TryOpen("root", string.Empty), Is.False);
        Assert.That(treeReader.TryOpen("child", string.Empty), Is.True);
    }

    [Test]
    public void TryOpen_opens_next_element_when_wrapped_XmlReader_is_not_on_element()
    {
        using var xmlReader = CreateXmlReader("""
                                              <root>
                                                text
                                                <child/>
                                              </root>
                                              """);
        xmlReader.Read();
        xmlReader.Read();
        Assert.That(xmlReader.NodeType, Is.Not.EqualTo(XmlNodeType.Element));
        Assert.That(xmlReader.NodeType, Is.EqualTo(XmlNodeType.Text));

        using var treeReader = new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance);

        Assert.That(treeReader.TryOpen("root", string.Empty), Is.False);
        Assert.That(treeReader.TryOpen("child", string.Empty), Is.True);
    }

    [TestCase(1)]
    [TestCase(2)]
    [TestCase(3)]
    [TestCase(64)]
    [Issue("73")]
    public void GetContent_with_buffer_copies_all_text_and_grows_the_buffer(int bufferLength)
    {
        // The emoji is a surrogate pair. The text is longer than each start buffer.
        var text = string.Concat(Enumerable.Repeat("ab&amp;&#x1F600;", 30));
        var expected = string.Concat(Enumerable.Repeat("ab&\U0001F600", 30)) + "<c>";
        var xml = $"<root>{text}<![CDATA[<c>]]></root>";
        using var reader = CreateReader(xml);
        reader.Open("root", string.Empty);

        var buffer = new char[bufferLength];
        var length = reader.GetContent(ref buffer);
        reader.Close("root", string.Empty);

        Assert.That(new string(buffer, 0, length), Is.EqualTo(expected));
    }

    [Test]
    [Issue("73")]
    public void GetContent_skips_entity_reference_of_reader_that_does_not_expand_entities()
    {
        const string xml = """<!DOCTYPE root [<!ENTITY e "x">]><root>a&e;b</root>""";
        using var xmlReader = new XmlTextReader(new StringReader(xml))
        {
            DtdProcessing = DtdProcessing.Parse,
            EntityHandling = EntityHandling.ExpandCharEntities,
        };
        using var reader = new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance);
        reader.Open("root", string.Empty);

        var buffer = new char[4];
        var length = reader.GetContent(ref buffer);
        reader.Close("root", string.Empty);

        Assert.That(new string(buffer, 0, length), Is.EqualTo("ab"));
    }

    [Test]
    [Issue("73")]
    public void GetContent_with_buffer_of_empty_element_returns_zero()
    {
        using var reader = CreateReader("<root></root>");
        reader.Open("root", string.Empty);

        var buffer = new char[4];
        var length = reader.GetContent(ref buffer);
        reader.Close("root", string.Empty);

        Assert.That(length, Is.Zero);
        Assert.That(buffer.Length, Is.EqualTo(4));
    }

    [Test]
    [Issue("73")]
    public void GetContent_with_buffer_throws_on_nested_element()
    {
        using var reader = CreateReader("<root>a<b/></root>");
        reader.Open("root", string.Empty);

        var buffer = new char[4];
        Assert.That(() => reader.GetContent(ref buffer), Throws.TypeOf<PartStructureException>());
    }

    [Test]
    [Issue("73")]
    public void TryGetAttributeChars_reads_value_longer_than_the_buffer()
    {
        var value = string.Concat(Enumerable.Repeat("x&amp;&#x1F600;", 20));
        var expected = string.Concat(Enumerable.Repeat("x&\U0001F600", 20));
        const string ns = "http://example.com/attributes";
        var xml = $"<root xmlns:q=\"{ns}\" long=\"{value}\" q:short=\"1\"><child/></root>";
        using var reader = CreateReader(xml);
        reader.Open("root", string.Empty);

        Assert.That(reader.TryGetAttributeChars("long", null, out var longValue), Is.True);
        Assert.That(longValue.ToString(), Is.EqualTo(expected));
        Assert.That(reader.TryGetAttributeChars("short", ns, out var shortValue), Is.True);
        Assert.That(shortValue.ToString(), Is.EqualTo("1"));
        Assert.That(reader.TryGetAttributeChars("short", null, out _), Is.False);
        Assert.That(reader.TryGetAttributeChars("missing", null, out _), Is.False);

        // The reader stays on the element.
        reader.Open("child", string.Empty);
        reader.Close("child", string.Empty);
        reader.Close("root", string.Empty);
    }

    private static XmlReader CreateXmlReader(string xml)
    {
        return XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        {
            IgnoreWhitespace = true,
            IgnoreComments = true,
            DtdProcessing = DtdProcessing.Prohibit,
        });
    }

    private static XmlTreeReader CreateReader(string xml)
    {
        return new XmlTreeReader(new MemoryStream(Encoding.UTF8.GetBytes(xml)), XmlToEnumMapper.Instance, true);
    }
}
