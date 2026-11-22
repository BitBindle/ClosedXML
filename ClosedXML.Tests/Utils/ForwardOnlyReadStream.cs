using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Tests.Utils
{
    /// <summary>
    /// A stream that the caller can only read from the start to the end, as a network stream. It
    /// cannot seek and it cannot write. A synchronous read throws, as in a web server that does not
    /// allow synchronous I/O.
    /// </summary>
    internal sealed class ForwardOnlyReadStream : Stream
    {
        private readonly MemoryStream _content;

        public ForwardOnlyReadStream(byte[] content)
        {
            _content = new MemoryStream(content, false);
        }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous operations are not allowed.");

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            return ValueTask.FromResult(_content.Read(buffer.Span));
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _content.Dispose();

            base.Dispose(disposing);
        }
    }
}
