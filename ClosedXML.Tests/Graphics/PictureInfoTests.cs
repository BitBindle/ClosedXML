using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;
using NUnit.Framework;

namespace ClosedXML.Tests.Graphics
{
    [TestFixture]
    public class PictureInfoTests
    {
        [Test]
        public void CanReadPng()
        {
            AssertRasterImage("SampleImagePng.png", XLPictureFormat.Png, new Size(252, 152), 96, 96);
        }

        [TestCase("SampleImageJfif.jpg", 176, 270, 96, 96)]
        [TestCase("jpeg-rgb.jpg", 200, 200, 0, 0)] // Adobe JPG, has APP14 marker right after SOI instead of APP0
        public void CanReadJfif(string filename, int widthPx, int heightPx, int dpiX, int dpiY)
        {
            AssertRasterImage($"Jpg.{filename}", XLPictureFormat.Jpeg, new Size(widthPx, heightPx), dpiX, dpiY);
        }

        [Test]
        public void CanReadExif()
        {
            AssertRasterImage("SampleImageExif.jpg", XLPictureFormat.Jpeg, new Size(252, 152), 0, 0);
        }

        [TestCase(0xE1, "http://ns.adobe.com/xap/1.0/\0")]
        [TestCase(0xE2, "ICC_PROFILE\0")]
        [TestCase(-1, "")]
        public void Jpeg_without_jfif_or_exif_segment_is_read(int appMarker, string identifier)
        {
            // JFIF and EXIF segments are optional. Encoders can put XMP or ICC segments first, or no
            // application segment at all, so that a quantization table follows the start of image.
            var jfif = ReadResource("Jpg.SampleImageJfif.jpg");
            var segments = new List<byte>();
            if (appMarker >= 0)
            {
                var payload = Encoding.ASCII.GetBytes(identifier + "data");
                var length = payload.Length + 2;
                segments.AddRange(new byte[] { 0xFF, (byte)appMarker, (byte)(length >> 8), (byte)length });
                segments.AddRange(payload);
            }

            var jpeg = ReplaceApplicationSegments(jfif, segments.ToArray());
            using var stream = new MemoryStream(jpeg);
            var info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(stream, XLPictureFormat.Unknown);

            Assert.AreEqual(XLPictureFormat.Jpeg, info.Format);
            Assert.AreEqual(new Size(176, 270), info.SizePx);
            Assert.AreEqual(0, info.DpiX);
            Assert.AreEqual(0, info.DpiY);
        }

        [Test]
        public void CanReadGif87Image()
        {
            AssertRasterImage("SampleImageGif87a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
        }

        [Test]
        public void CanReadGif89Image()
        {
            AssertRasterImage("SampleImageGif89a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
        }

        [TestCase("SampleImageBmpWin24bit.bmp")]
        [TestCase("SampleImageBmpWin8bit.bmp")]
        [TestCase("SampleImageBmpWin4bit.bmp")]
        [TestCase("SampleImageBmpWin24bit.bmp")]
        public void CanReadBmpImageV3AndFurther(string imageName)
        {
            AssertRasterImage(imageName, XLPictureFormat.Bmp, new Size(167, 51), 80.645d, 80.645d);
        }

        [Test]
        public void CanReadBmpV1()
        {
            AssertRasterImage("SampleImageBmpV1.bmp", XLPictureFormat.Bmp, new Size(150, 50), 0, 0);
        }

        [Test]
        public void CanReadTiffWithBigEndianEncoding()
        {
            AssertRasterImage("SampleImageTiffBigEndian.tiff", XLPictureFormat.Tiff, new Size(130, 45), 96, 96);
        }

        [Test]
        public void CanReadTiffWithLittleEndianEncoding()
        {
            AssertRasterImage("SampleImageTiffLittleEndian.tiff", XLPictureFormat.Tiff, new Size(130, 45), 96, 96);
        }

        [Test]
        public void CanReadPcx()
        {
            AssertRasterImage("SampleImagePcx.pcx", XLPictureFormat.Pcx, new Size(100, 50), 96, 96);
        }

        [Test]
        public void CanReadWmfWithPlaceableHeader()
        {
            AssertVectorImage("SampleImagePlaceableWmf.wmf", XLPictureFormat.Wmf, new Size(1000, 500));
        }

        [Test]
        public void CanReadWmfWithOriginalHeader()
        {
            AssertVectorImage("SampleImageOriginalWmf.wmf", XLPictureFormat.Wmf, new Size(12496, 6247));
        }

        [Test]
        public void CanReadEmf()
        {
            AssertVectorImage("SampleImageEmf.emf", XLPictureFormat.Emf, new Size(28844, 28938));
        }

        [Test]
        public void CanReadExtendedWebp()
        {
            AssertRasterImage("SampleImageWebpExtendedFormat.webp", XLPictureFormat.Webp, new Size(188, 231), 72, 72);
        }

        [Test]
        public void CanReadLossyWebp()
        {
            AssertRasterImage("SampleImageWebpLossy.webp", XLPictureFormat.Webp, new Size(278, 90), 72, 72);
        }

        [Test]
        public void CanReadLosslessWebp()
        {
            AssertRasterImage("SampleImageWebpLossless.webp", XLPictureFormat.Webp, new Size(395, 136), 72, 72);
        }

        [Issue("17")]
        [TestCase("SampleImagePng.png", XLPictureFormat.Png)]
        [TestCase("Jpg.SampleImageJfif.jpg", XLPictureFormat.Jpeg)]
        [TestCase("SampleImageGif89a.gif", XLPictureFormat.Gif)]
        [TestCase("SampleImageBmpWin24bit.bmp", XLPictureFormat.Bmp)]
        [TestCase("SampleImageTiffLittleEndian.tiff", XLPictureFormat.Tiff)]
        [TestCase("SampleImagePcx.pcx", XLPictureFormat.Pcx)]
        [TestCase("SampleImagePlaceableWmf.wmf", XLPictureFormat.Wmf)]
        [TestCase("SampleImageOriginalWmf.wmf", XLPictureFormat.Wmf)]
        [TestCase("SampleImageEmf.emf", XLPictureFormat.Emf)]
        [TestCase("SampleImageWebpLossless.webp", XLPictureFormat.Webp)]
        public void Picture_info_is_same_when_stream_returns_one_byte_per_read(string imageName, XLPictureFormat expectedFormat)
        {
            // Stream.Read can return fewer bytes than the buffer length before the end of the stream.
            using var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"ClosedXML.Tests.Resource.Images.{imageName}");
            using var memoryStream = new MemoryStream();
            resourceStream.CopyTo(memoryStream);
            var expectedInfo = DefaultGraphicEngine.Instance.Value.GetPictureInfo(memoryStream, XLPictureFormat.Unknown);

            using var oneBytePerReadStream = new OneBytePerReadStream(memoryStream.ToArray());
            var info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(oneBytePerReadStream, XLPictureFormat.Unknown);

            Assert.AreEqual(expectedFormat, info.Format);
            Assert.AreEqual(expectedInfo.SizePx, info.SizePx);
            Assert.AreEqual(expectedInfo.SizePhys, info.SizePhys);
            Assert.AreEqual(expectedInfo.DpiX, info.DpiX);
            Assert.AreEqual(expectedInfo.DpiY, info.DpiY);
        }

        private static byte[] ReadResource(string imageName)
        {
            using var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"ClosedXML.Tests.Resource.Images.{imageName}");
            using var memoryStream = new MemoryStream();
            resourceStream.CopyTo(memoryStream);
            return memoryStream.ToArray();
        }

        /// <summary>
        /// Remove the APPn segments of a JPEG and put <paramref name="newSegments"/> after the start of image.
        /// </summary>
        private static byte[] ReplaceApplicationSegments(byte[] jpeg, byte[] newSegments)
        {
            var result = new List<byte> { jpeg[0], jpeg[1] };
            result.AddRange(newSegments);
            var position = 2;
            while (true)
            {
                var marker = jpeg[position + 1];
                var segmentLength = 2 + ((jpeg[position + 2] << 8) | jpeg[position + 3]);
                var isApplicationSegment = marker >= 0xE0 && marker <= 0xEF;
                if (!isApplicationSegment)
                {
                    // The other segments, the scan data and the end of image stay as they are.
                    result.AddRange(new ArraySegment<byte>(jpeg, position, jpeg.Length - position));
                    return result.ToArray();
                }

                position += segmentLength;
            }
        }

        private static void AssertRasterImage(string imageName, XLPictureFormat expectedFormat, Size expectedPxSize, double expectedDpiX, double expectedDpiY)
        {
            AssertImage(imageName, expectedFormat, expectedPxSize, Size.Empty, expectedDpiX, expectedDpiY);
        }

        private static void AssertVectorImage(string imageName, XLPictureFormat expectedFormat, Size expectedHiMetricSize)
        {
            AssertImage(imageName, expectedFormat, Size.Empty, expectedHiMetricSize, 0, 0);
        }

        private static void AssertImage(string imageName, XLPictureFormat expectedFormat, Size expectedPxSize, Size expectedHiMetricSize, double expectedDpiX, double expectedDpiY)
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"ClosedXML.Tests.Resource.Images.{imageName}");
            var info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(stream, XLPictureFormat.Unknown);

            Assert.AreEqual(expectedFormat, info.Format);
            Assert.AreEqual(expectedPxSize, info.SizePx);
            Assert.AreEqual(expectedHiMetricSize, info.SizePhys);

            // Some DPI is stored as pixels per meter, causing a rounding errors.
            Assert.AreEqual(expectedDpiX, info.DpiX, 0.02);
            Assert.AreEqual(expectedDpiY, info.DpiY, 0.02);
        }

        private sealed class OneBytePerReadStream : MemoryStream
        {
            public OneBytePerReadStream(byte[] data) : base(data, false)
            {
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                return base.Read(buffer, offset, Math.Min(count, 1));
            }

            public override int Read(Span<byte> buffer)
            {
                return base.Read(buffer.Slice(0, Math.Min(buffer.Length, 1)));
            }
        }
    }
}
