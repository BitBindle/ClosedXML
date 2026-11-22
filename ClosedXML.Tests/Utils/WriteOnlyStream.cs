using System;
using System.IO;

namespace ClosedXML.Tests.Utils
{
    /// <summary>
    /// A stream that the caller can only write to, as the response body of a web server.
    /// It cannot read and it cannot seek. It keeps the written bytes for the test.
    /// </summary>
    internal sealed class WriteOnlyStream : Stream
    {
        private readonly MemoryStream _written = new();

        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <summary>
        /// Get a readable copy of the written bytes.
        /// </summary>
        public MemoryStream GetWrittenBytes() => new(_written.ToArray());

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => _written.Write(buffer, offset, count);
    }
}
