using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Tests.Utils
{
    /// <summary>
    /// A stream that the caller can only write to, as the response body of a web server.
    /// It cannot read and it cannot seek. It keeps the written bytes for the test.
    /// </summary>
    internal sealed class WriteOnlyStream : Stream
    {
        private readonly MemoryStream _written = new();

        /// <summary>
        /// When false, a synchronous write or flush throws, as in a web server that does not allow
        /// synchronous I/O.
        /// </summary>
        public bool AllowSynchronousIO { get; init; } = true;

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

        public override void Flush() => ThrowIfSynchronousIONotAllowed();

        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count)
        {
            ThrowIfSynchronousIONotAllowed();
            _written.Write(buffer, offset, count);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            _written.Write(buffer, offset, count);
            return Task.CompletedTask;
        }

        private void ThrowIfSynchronousIONotAllowed()
        {
            if (!AllowSynchronousIO)
                throw new InvalidOperationException("Synchronous operations are not allowed.");
        }
    }
}
