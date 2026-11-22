using System;
using System.Xml;

namespace ClosedXML.IO;

/// <summary>
/// A simplified XML reader that reads the content and hides full complexity the <see cref="XmlReader"/>.
/// </summary>
public interface IXmlReader : IDisposable
{
    /// <summary>
    /// Read next node. If no more nodes can be read, return <c>false</c>.
    /// </summary>
    bool Read();

    /// <summary>
    /// A node reader is currently on.
    /// </summary>
    XmlTreeNodeType NodeType { get; }

    /// <summary>
    /// Get depth of current element.
    /// </summary>
    int Depth { get; }

    /// <summary>
    /// Get position of current element.
    /// </summary>
    LineInfo LineInfo { get; }

    /// <summary>
    /// Name of an open/close element. If not on an element, return an empty string.
    /// </summary>
    /// <remarks>The name is atomized.</remarks>
    string LocalName { get; }

    /// <summary>
    /// Namespace of an open/close element. If not on an element, return an empty string.
    /// </summary>
    /// <remarks>The namespace is atomized.</remarks>
    string NamespaceUri { get; }

    /// <summary>
    /// Value of a <see cref="XmlTreeNodeType.Text"/> node. Empty string for other node types.
    /// </summary>
    string Value { get; }

    /// <summary>
    /// Get attribute value. If attribute is not found or reader is not on open element, return null.
    /// </summary>
    string? GetAttribute(string attributeName, string? namespaceUri);

    /// <summary>
    /// Copy the value of an attribute to the start of the <paramref name="buffer"/>. If the buffer
    /// is too small, the method puts a larger buffer in <paramref name="buffer"/>.
    /// </summary>
    /// <returns>
    /// True, if the method found the attribute. False, if the attribute is not found or the reader
    /// is not on an open element.
    /// </returns>
    bool TryCopyAttribute(string attributeName, string? namespaceUri, ref char[] buffer, out int length)
    {
        if (GetAttribute(attributeName, namespaceUri) is not { } value)
        {
            length = 0;
            return false;
        }

        length = CopyToBuffer(value, ref buffer, 0);
        return true;
    }

    /// <summary>
    /// Copy the value of a <see cref="XmlTreeNodeType.Text"/> node to the <paramref name="buffer"/>
    /// at the <paramref name="index"/>. If the buffer is too small, the method puts a larger buffer
    /// in <paramref name="buffer"/>. After this call, the value of the node is not available again.
    /// </summary>
    /// <returns>The number of copied characters.</returns>
    int CopyValue(ref char[] buffer, int index)
    {
        return CopyToBuffer(Value, ref buffer, index);
    }

    /// <summary>
    /// Copy the <paramref name="value"/> to the <paramref name="buffer"/> at the <paramref name="index"/>.
    /// If the buffer is too small, put a larger buffer in <paramref name="buffer"/>.
    /// </summary>
    /// <returns>The number of copied characters.</returns>
    protected static int CopyToBuffer(string value, ref char[] buffer, int index)
    {
        var requiredLength = index + value.Length;
        if (buffer.Length < requiredLength)
            Array.Resize(ref buffer, Math.Max(buffer.Length * 2, requiredLength));

        value.CopyTo(0, buffer, index, value.Length);
        return value.Length;
    }
}
