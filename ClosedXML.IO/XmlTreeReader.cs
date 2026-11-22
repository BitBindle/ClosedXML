using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Xml;
using static ClosedXML.IO.XmlTreeNodeType;

namespace ClosedXML.IO;

/// <summary>
/// <para>
/// Reader that expects that XML document consists of elements in a tree-like fashion. XML
/// shouldn't be mixed, text should be only in the leaves (e.g. <c>&lt;f&gt;ABS(A1)*A2&lt;/f&gt;</c>).
/// </para>
/// <para>
/// The schema of XML should be mostly by elements, with choices and sequences.
/// </para>
/// <para>
/// In case of some specialities or mixed content, use <c>XDocument.Load(_reader.ReadSubtree())</c>
/// and parse the result.
/// </para>
/// <para>
/// All <c>Get*</c> methods read values from attributes of current element.
/// </para>
/// <para>
/// The reader is always at either start element or end element. Any API that moves the reader will
/// end on either start or end element. The empty elements (e.g. &lt;br/&gt;) behave same way as
/// non-empty elements (that is different from <see cref="XmlReader"/> that checks
/// <see cref="XmlReader.IsEmptyElement"/>).
/// The reader element is used for one of two purposes:
/// <list type="bullet">
/// <item>
/// Element is being processed, i.e. parser logic has correctly identified the element (by
/// name) and will use parser logic to extract data from element (mostly by reading attributes,
/// potentially content if it is a leaf).
/// </item>
/// <item>
/// Element is used as a lookahead. The parsing logic is using it to determine how to parse the rest
/// of document. Example:
/// <example>Stylesheet parser has processed element <![CDATA[<name/>]]> from <![CDATA[<font><name val="Arial"/></font>]]>).
/// It reads next element <![CDATA[</font>]]> as a lookahead. The parsing logic now has to
/// determine what to do. Font could have additional properties (e.g. <![CDATA[<b/>]]>) or the
/// schema could end. Parser will use <c>reader.TryOpen("b")</c> to check if there is a bold
/// element. If there isn't, it uses <c>reader.TryClose("font")</c> to check that <c>font</c>
/// should close. If neither is true, XML doesn't match expected schema and parser will likely
/// throw an exception.
/// </example>
/// </item>
/// </list>
/// </para>
/// </summary>
public sealed class XmlTreeReader : IDisposable
{
    private static readonly XmlReaderSettings Settings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        Async = false,
        DtdProcessing = DtdProcessing.Prohibit,
        CloseInput = true
    };

    /// <summary>
    /// The reader that holds the current element. The current node must always be either
    /// <see cref="OpenElement"/> or <see cref="CloseElement"/> after an API method call.
    /// </summary>
    private readonly IXmlReader _reader;

    private readonly IEnumMapper _enumMapper;

    /// <summary>
    /// What is current state of parser:
    /// <list type="bullet">
    /// <item>
    ///   <term>false</term>
    ///   <description>
    ///   Current element is being processed. we can get value of attributes.
    ///   </description>
    /// </item>
    /// <item>
    ///   <term>true</term>
    ///   <description>
    ///   Current element is not being processed. The only thing we are interested in is a name and
    ///   open/close. We are using it to determine how to parse the remainder of the file.
    ///   Trying to get attribute value will throw.
    ///   </description>
    /// </item>
    /// </list>
    /// </summary>
    private bool _inLookup;

    private readonly List<string> _context = new();

    /// <summary>
    /// The white space characters of XML. <see cref="XmlConvert"/> removes them before it parses a value.
    /// </summary>
    private const string XmlWhitespace = " \t\n\r";

    /// <summary>
    /// The number styles of <see cref="XmlConvert.ToInt32(string)"/> and <see cref="XmlConvert.ToInt64(string)"/>.
    /// </summary>
    private const NumberStyles XmlIntegerStyle = NumberStyles.AllowLeadingSign | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite;

    /// <summary>
    /// The number styles of <see cref="XmlConvert.ToDouble(string)"/>, after it removes the white space.
    /// </summary>
    private const NumberStyles XmlDoubleStyle = NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent;

    /// <summary>
    /// The buffer that keeps the value of the last attribute that <see cref="TryGetAttributeChars"/> read.
    /// </summary>
    private char[] _attributeBuffer = new char[32];
    private char[] _contentBuffer = new char[64];

    public XmlTreeReader(Stream stream, IEnumMapper enumMapper, bool strictAttributeParsing)
        : this(XmlReader.Create(stream, Settings), enumMapper, strictAttributeParsing, leaveOpen: false)
    {
    }

    /// <summary>
    /// Create a new instance MCE-aware reader that wraps around the <see cref="XmlReader"/>. The reader will throw
    /// on malformed attributes and will close the underlaying stream when disposed.
    /// </summary>
    /// <param name="xmlReader">Wrapped XML reader.</param>
    /// <param name="enumMapper">A mapper for converting enum values to </param>
    public XmlTreeReader(XmlReader xmlReader, IEnumMapper enumMapper)
        : this(xmlReader, enumMapper, true, false)
    {
    }

    /// <summary>
    /// Create a new instance MCE-aware reader that wraps around the <see cref="XmlReader"/>.
    /// </summary>
    /// <param name="xmlReader">Wrapped XML reader.</param>
    /// <param name="enumMapper">A mapper for converting enum values to </param>
    /// <param name="strictAttributeParsing">Should malformed attributes throw an exception (<c>false</c>) or be treated as missing attributes (<c>false</c>)?</param>
    /// <param name="leaveOpen">Should the <paramref name="xmlReader"/> be closed and disposed when this instance is disposed of?</param>
    public XmlTreeReader(XmlReader xmlReader, IEnumMapper enumMapper, bool strictAttributeParsing, bool leaveOpen)
    {
        _reader = new MceXmlReader(xmlReader, new MceSettings
        {
            SignalMismatch = info => throw PartStructureException.MceError(info.LineInfo, "Mismatch between consuming application capability and document requirements.")
        }, leaveOpen);
        _enumMapper = enumMapper;
        StrictAttributeParsing = strictAttributeParsing;

        // The usage pattern of XmlTreeReader expects to call TryOpen as the first call. Set the lookup state so it
        // tries to open the current or closest next element.
        _inLookup = xmlReader.NodeType is XmlNodeType.Element;
    }

    /// <summary>
    /// Returns names of elements from currently processed element to root. The <c>On*Parsed</c>
    /// hook is called after processed element is closed. Therefore when the context is inspected
    /// in the <c>On*Parsed</c> hook, it doesn't contain the name of element that was just processed.
    /// </summary>
    public IReadOnlyList<string> Context => _context;

    /// <summary>
    /// Should attributes with values outside of their simple type be treated as errors and throw
    /// an exception or be treated as missing and just return null? Excel generally ignores
    /// unparseable attributes.
    /// </summary>
    public bool StrictAttributeParsing { get; }

    /// <summary>
    /// Get the local name of current element (lookup/processing).
    /// </summary>
    internal string ElementName => _reader.LocalName;

    internal LineInfo LineInfo => _reader.LineInfo;

    /// <summary>
    /// Read next element. Check lookup element is <paramref name="localName"/>. If it is, open the
    /// element and return true. Otherwise, return false (element doesn't change).
    /// </summary>
    public bool TryOpen(string localName, string namespaceUri)
    {
        SwitchToLookup();
        Debug.Assert(_reader.NodeType is OpenElement or CloseElement);

        if (_reader.NodeType == OpenElement && _reader.LocalName == localName && _reader.NamespaceUri == namespaceUri)
        {
            // Element has been opened, so it should be processed.
            SwitchToProcessing();
            _context.Add(localName);
            return true;
        }

        return false;
    }

    public bool TryClose(string localName, string namespaceUri)
    {
        SwitchToLookup();
        Debug.Assert(_reader.NodeType is OpenElement or CloseElement);

        if (_reader.NodeType == OpenElement || _reader.LocalName != localName || _reader.NamespaceUri != namespaceUri)
            return false;

        // Element has been closed, so it should be processed. Though closing elements are not
        // really processed, but we don't want to switch back to lookup. We just want to mark it
        // as "done." Be lazy, e.g. last element of a document doesn't have next element. If
        // parsing logic needs further elements, it will read them when they are needed.
        SwitchToProcessing();
        _context.RemoveAt(_context.Count - 1);
        return true;
    }

    /// <summary>
    /// Assert that we are at the element with <paramref name="localName"/>. Doesn't move anywhere.
    /// </summary>
    public void Open(string localName, string namespaceUri)
    {
        if (!TryOpen(localName, namespaceUri))
            throw PartStructureException.ExpectedElementNotFound($"Expected opening element '{localName}', but reader is currently on {_reader.NodeType} node '{_reader.LocalName}'.", this);
    }

    /// <summary>
    /// Close the next unprocessed node. If the node doesn't match the <paramref name="localName"/>,
    /// throw an exception.
    /// </summary>
    public void Close(string localName, string namespaceUri)
    {
        if (!TryClose(localName, namespaceUri))
            throw PartStructureException.ExpectedElementNotFound($"Expected closing element '{localName}', but reader is currently on {_reader.NodeType} node '{_reader.LocalName}'.", this);
    }

    /// <summary>
    /// Skip subtree that start on the current element. After subtree is read, the reader is
    /// on an ending element of a subtree in a processed state.
    /// </summary>
    /// <exception cref="InvalidOperationException">Reader isn't on opening element.</exception>
    public void Skip(string elementName)
    {
        ThrowOnNonStartElement();
        var startDepth = _reader.Depth;
        do
        {
            if (!_reader.Read())
                throw InvalidXml();
        } while (_reader.NodeType != CloseElement || _reader.Depth > startDepth);

        _inLookup = false;
    }

    public bool? GetOptionalBool(string attributeName, string? ns = null)
    {
        if (!TryGetAttributeChars(attributeName, ns, out var value))
            return null;

        // The same values as XmlConvert.ToBoolean.
        var trimmed = value.Trim(XmlWhitespace);
        if (trimmed is "1" or "true")
            return true;

        if (trimmed is "0" or "false")
            return false;

        if (StrictAttributeParsing)
            ThrowAttributeFormatException(attributeName, value.ToString());

        return null;
    }

    public byte? GetOptionalUByte(string attributeName, string? ns = null)
    {
        ThrowOnNonStartElement();
        byte? number = null;
        if (_reader.GetAttribute(attributeName, ns) is { } value)
        {
            try
            {
                number = XmlConvert.ToByte(value);
            }
            catch (OverflowException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
        }

        return number;
    }

    public int? GetOptionalInt(string attributeName, string? ns = null)
    {
        if (!TryGetAttributeChars(attributeName, ns, out var value))
            return null;

        if (int.TryParse(value, XmlIntegerStyle, NumberFormatInfo.InvariantInfo, out var number))
            return number;

        if (StrictAttributeParsing)
            ThrowAttributeFormatException(attributeName, value.ToString());

        return null;
    }

    public uint? GetOptionalUInt(string attributeName, string? ns = null)
    {
        if (!TryGetAttributeChars(attributeName, ns, out var value))
            return null;

        if (long.TryParse(value, XmlIntegerStyle, NumberFormatInfo.InvariantInfo, out var number) && number is >= 0 and <= uint.MaxValue)
            return (uint)number;

        if (StrictAttributeParsing)
            ThrowAttributeFormatException(attributeName, value.ToString());

        return null;
    }

    public double? GetOptionalDouble(string attributeName, string? ns = null)
    {
        if (!TryGetAttributeChars(attributeName, ns, out var value))
            return null;

        // XmlConvert.ToDouble also accepts INF and -INF, but this method does not accept an
        // infinity or NaN.
        if (double.TryParse(value.Trim(XmlWhitespace), XmlDoubleStyle, NumberFormatInfo.InvariantInfo, out var number) && double.IsFinite(number))
            return number;

        if (StrictAttributeParsing)
            ThrowAttributeFormatException(attributeName, value.ToString());

        return null;
    }

    /// <summary>
    /// Get the characters of an attribute of the current element. The method does not make a new
    /// string. The <paramref name="value"/> is valid only until the next call to this reader.
    /// </summary>
    /// <returns>True, if the element has the attribute. False, if the element does not have the attribute.</returns>
    public bool TryGetAttributeChars(string attributeName, string? ns, out ReadOnlySpan<char> value)
    {
        ThrowOnNonStartElement();
        if (!_reader.TryCopyAttribute(attributeName, ns, ref _attributeBuffer, out var length))
        {
            value = default;
            return false;
        }

        value = _attributeBuffer.AsSpan(0, length);
        return true;
    }

    /// <summary>
    /// Try to read <c>xsd:dateTime</c> from an attribute of current element.
    /// </summary>
    /// <returns>Read datetime or null if attribute is not present.</returns>
    public DateTime? GetOptionalDateTime(string attributeName, string? ns = null)
    {
        ThrowOnNonStartElement();
        DateTime? dateTime = null;
        if (_reader.GetAttribute(attributeName, ns) is { } value)
        {
            try
            {
                dateTime = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind);
            }
            catch (FormatException e)
            {
                if (StrictAttributeParsing)
                    ThrowAttributeFormatException(attributeName, value, e);
            }
        }

        return dateTime;
    }

    public string? GetOptionalString(string attributeName, string? ns = null)
    {
        ThrowOnNonStartElement();
        return _reader.GetAttribute(attributeName, ns);
    }

    public TEnum? GetOptionalEnum<TEnum>(string attributeName, string? ns = null)
        where TEnum : struct, Enum
    {
        ThrowOnNonStartElement();
        var enumString = _reader.GetAttribute(attributeName, ns);

        if (enumString is null)
        {
            return null;
        }

        if (!_enumMapper.TryGetEnum<TEnum>(enumString, out var enumValue))
        {
            if (StrictAttributeParsing)
                ThrowAttributeFormatException(attributeName, enumString);

            return null;
        }

        return enumValue;
    }

    /// <summary>
    /// Get the content of current element. Once it ends, it should be on
    /// the <see cref="XmlTreeNodeType.CloseElement"/> of the tree.
    /// </summary>
    /// <exception cref="PartStructureException">If another element is found inside the current element.</exception>>
    public string GetContent()
    {
        var length = GetContent(ref _contentBuffer);
        return length == 0 ? string.Empty : new string(_contentBuffer, 0, length);
    }

    /// <summary>
    /// Copy the content of the current element to the start of the <paramref name="buffer"/>. The
    /// method does not make a new string. If the buffer is too small, the method puts a larger
    /// buffer in <paramref name="buffer"/>. Once it ends, it should be on the
    /// <see cref="XmlTreeNodeType.CloseElement"/> of the tree.
    /// </summary>
    /// <returns>The number of characters of the content.</returns>
    /// <exception cref="PartStructureException">If another element is found inside the current element.</exception>
    public int GetContent(ref char[] buffer)
    {
        ThrowOnNonStartElement();

        var length = 0;
        while (_reader.Read())
        {
            switch (_reader.NodeType)
            {
                case OpenElement:
                    throw PartStructureException.UnexpectedElementFound(_reader.LocalName);
                case CloseElement:
                    _inLookup = true;
                    return length;
                case Text:
                    length += _reader.CopyValue(ref buffer, length);
                    break;
                default:
                    throw new UnreachableException();
            }
        }

        throw InvalidXml();
    }

    public void Dispose()
    {
        _reader.Dispose();
    }

    private void SwitchToProcessing()
    {
        if (_inLookup)
            _inLookup = false;
    }

    private void SwitchToLookup()
    {
        // When switching to lookup, current node and all its attributes should have already been processed.
        if (_inLookup)
            return;

        // Read next element.
        MoveToNextElement();
        _inLookup = true;
    }

    /// <summary>
    /// Move from current opening/closing element to next opening/closing element.
    /// </summary>
    private void MoveToNextElement()
    {
        while (_reader.Read())
        {
            if (_reader.NodeType is OpenElement or CloseElement)
            {
                return;
            }
        }

        throw InvalidXml();
    }

    private void ThrowOnNonStartElement()
    {
        if (_reader.NodeType != OpenElement || _inLookup)
            throw new InvalidOperationException("To read content/attribute, the reader must be on start element and in processing state.");
    }

    private void ThrowAttributeFormatException(string attributeName, string attributeValue, Exception? exception = null)
    {
        throw PartStructureException.InvalidAttributeFormat(attributeName, attributeValue, this, exception);
    }

    private Exception InvalidXml()
    {
        // This should never happen. The underlaying reader should throw when it detects
        // invalid XML (no root, unpaired elements, multiple elements in root)
        return new UnreachableException($"{_reader.LineInfo}: Invalid XML.");
    }
}
