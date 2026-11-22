// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Graphics;

namespace ClosedXML.Excel.Drawings
{
    [DebuggerDisplay("{Name}")]
    internal sealed class XLPicture : IXLPicture, IDisposable
    {
        private Int32 _height;
        private Int32 _id;
        private String _name = string.Empty;
        private Int32 _width;
        private bool _disposed;

        internal XLPicture(XLWorksheet worksheet, Stream stream)
            : this(worksheet, stream, XLPictureFormat.Unknown)
        {
        }

        internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
            : this(worksheet, stream, format, copyStream: true)
        {
        }

        /// <summary>
        /// Make a picture that uses the <paramref name="imageStream"/> as its image data. The method
        /// does not copy the data. The picture owns the stream and disposes it.
        /// </summary>
        internal static XLPicture FromOwnedStream(IXLWorksheet worksheet, MemoryStream imageStream)
        {
            return new XLPicture(worksheet, imageStream, XLPictureFormat.Unknown, copyStream: false);
        }

        private XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format, bool copyStream)
        {
            this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            this.Placement = XLPicturePlacement.MoveAndSize;
            this.Markers = new Dictionary<XLMarkerPosition, XLMarker?>()
            {
                [XLMarkerPosition.TopLeft] = null,
                [XLMarkerPosition.BottomRight] = null
            };

            // Calculate default picture ID
            var allPictures = worksheet.Workbook.Worksheets.SelectMany(ws => ws.Pictures);
            var freeId = allPictures.Select(x => x.Id).DefaultIfEmpty(0).Max() + 1;
            _id = freeId;

            var info = worksheet.Workbook.GraphicEngine.GetPictureInfo(stream, format);
            Init(info);

            if (copyStream)
            {
                ImageStream = new MemoryStream();
                stream.Position = 0;
                stream.CopyTo(ImageStream);
            }
            else
            {
                ImageStream = (MemoryStream)stream;
            }

            ImageStream.Seek(0, SeekOrigin.Begin);
        }

        public IXLCell? BottomRightCell
        {
            get => Markers[XLMarkerPosition.BottomRight]?.Cell;
            private set
            {
                ArgumentNullException.ThrowIfNull(value);
                if (!value.Worksheet.Equals(this.Worksheet))
                    throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

                Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
            }
        }

        public XLPictureFormat Format { get; private set; } = XLPictureFormat.Unknown;

        public Int32 Height
        {
            get => _height;
            set
            {
                if (Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the height, the placement should be FreeFloating or Move");

                _height = value;
            }
        }

        public Int32 Id
        {
            get => _id;
            internal set
            {
                if ((Worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
                    throw new ArgumentException($"The picture ID '{value}' already exists.");

                _id = value;
            }
        }

        public MemoryStream ImageStream
        {
            get
            {
                ThrowIfDisposed();
                return field;
            }
            private init;
        }

        public Int32 Left
        {
            get => Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0;
            set
            {
                if (Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new System.Drawing.Point(value, Top));
            }
        }

        public String Name
        {
            get => _name;
            set
            {
                if (_name == value)
                    return;

                if ((Worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? this) != this)
                    throw new ArgumentException($"The picture name '{value}' already exists.");

                SetName(value);
            }
        }

        public Int32 OriginalHeight { get; private set; }

        public Int32 OriginalWidth { get; private set; }

        public XLPicturePlacement Placement { get; set; }

        public Int32 Top
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the top offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new System.Drawing.Point(this.Left, value));
            }
        }

        public IXLCell? TopLeftCell
        {
            get
            {
                return Markers[XLMarkerPosition.TopLeft]?.Cell;
            }

            private set
            {
                ArgumentNullException.ThrowIfNull(value);
                if (!value.Worksheet.Equals(this.Worksheet))
                    throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

                this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
            }
        }

        public Int32 Width
        {
            get { return _width; }
            set
            {
                if (this.Placement == XLPicturePlacement.MoveAndSize)
                    throw new ArgumentException("To set the width, the placement should be FreeFloating or Move");
                _width = value;
            }
        }

        public IXLWorksheet Worksheet { get; }

        internal IDictionary<XLMarkerPosition, XLMarker?> Markers { get; private set; }

        internal String? RelId { get; set; }

        /// <summary>
        /// Create a copy of the picture on a different worksheet.
        /// </summary>
        /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
        /// <returns>A created copy of the picture.</returns>
        public IXLPicture CopyTo(IXLWorksheet targetSheet)
        {
            return CopyTo((XLWorksheet)targetSheet);
        }

        public void Delete()
        {
            Worksheet.Pictures.Delete(this.Name);
        }

        /// <summary>
        /// Create a copy of the picture on the same worksheet.
        /// </summary>
        /// <returns>A created copy of the picture.</returns>
        public IXLPicture Duplicate()
        {
            return CopyTo(Worksheet);
        }

        public System.Drawing.Point GetOffset(XLMarkerPosition position)
        {
            return GetMarker(position).Offset;
        }

        public IXLPicture MoveTo(Int32 left, Int32 top)
        {
            this.Placement = XLPicturePlacement.FreeFloating;
            this.Left = left;
            this.Top = top;
            return this;
        }

        public IXLPicture MoveTo(IXLCell cell)
        {
            return MoveTo(cell, 0, 0);
        }

        public IXLPicture MoveTo(IXLCell cell, Int32 xOffset, Int32 yOffset)
        {
            return MoveTo(cell, new System.Drawing.Point(xOffset, yOffset));
        }

        public IXLPicture MoveTo(IXLCell cell, System.Drawing.Point offset)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            this.Placement = XLPicturePlacement.Move;
            this.TopLeftCell = cell;
            GetMarker(XLMarkerPosition.TopLeft).Offset = offset;
            return this;
        }

        public IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell)
        {
            return MoveTo(fromCell, 0, 0, toCell, 0, 0);
        }

        public IXLPicture MoveTo(IXLCell fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXLCell toCell, Int32 toCellXOffset, Int32 toCellYOffset)
        {
            return MoveTo(fromCell, new System.Drawing.Point(fromCellXOffset, fromCellYOffset), toCell, new System.Drawing.Point(toCellXOffset, toCellYOffset));
        }

        public IXLPicture MoveTo(IXLCell fromCell, System.Drawing.Point fromOffset, IXLCell toCell, System.Drawing.Point toOffset)
        {
            if (fromCell == null) throw new ArgumentNullException(nameof(fromCell));
            if (toCell == null) throw new ArgumentNullException(nameof(toCell));
            this.Placement = XLPicturePlacement.MoveAndSize;

            this.TopLeftCell = fromCell;
            GetMarker(XLMarkerPosition.TopLeft).Offset = fromOffset;

            this.BottomRightCell = toCell;
            GetMarker(XLMarkerPosition.BottomRight).Offset = toOffset;

            return this;
        }

        public IXLPicture Scale(Double factor, Boolean relativeToOriginal = false)
        {
            return this.ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);
        }

        public IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false)
        {
            this.Height = Convert.ToInt32((relativeToOriginal ? this.OriginalHeight : this.Height) * factor);
            return this;
        }

        public IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false)
        {
            this.Width = Convert.ToInt32((relativeToOriginal ? this.OriginalWidth : this.Width) * factor);
            return this;
        }

        public IXLPicture WithPlacement(XLPicturePlacement value)
        {
            this.Placement = value;
            return this;
        }

        public IXLPicture WithSize(Int32 width, Int32 height)
        {
            this.Width = width;
            this.Height = height;
            return this;
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            ImageStream.Dispose();
            _disposed = true;
        }

        internal IXLPicture CopyTo(XLWorksheet targetSheet)
        {
            ThrowIfDisposed();
            if (targetSheet == null)
                targetSheet = (XLWorksheet)Worksheet;

            IXLPicture newPicture;
            if (targetSheet == Worksheet)
                newPicture = targetSheet.AddPicture(ImageStream, Format);
            else
                newPicture = targetSheet.AddPicture(ImageStream, Format, Name);

            newPicture = newPicture
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .WithSize(Width, Height)
                    .WithPlacement(Placement);

            switch (Placement)
            {
                case XLPicturePlacement.FreeFloating:
                    newPicture.MoveTo(Left, Top);
                    break;

                case XLPicturePlacement.Move:
                    newPicture.MoveTo(targetSheet.Cell(GetMarker(XLMarkerPosition.TopLeft).Cell.Address), GetOffset(XLMarkerPosition.TopLeft));
                    break;

                case XLPicturePlacement.MoveAndSize:
                    newPicture.MoveTo(targetSheet.Cell(GetMarker(XLMarkerPosition.TopLeft).Cell.Address), GetOffset(XLMarkerPosition.TopLeft), targetSheet.Cell(GetMarker(XLMarkerPosition.BottomRight).Cell.Address),
                        GetOffset(XLMarkerPosition.BottomRight));
                    break;
            }

            return newPicture;
        }

        private XLMarker GetMarker(XLMarkerPosition position)
        {
            return Markers[position] ?? throw new InvalidOperationException($"The picture has no {position} anchor.");
        }

        internal void SetName(string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("Picture names cannot be null or empty.");

            _name = value;
        }

        private void Init(XLPictureInfo info)
        {
            Format = info.Format;
            var size = info.GetSizePx(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
            _width = OriginalWidth = size.Width;
            _height = OriginalHeight = size.Height;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(XLPicture));
        }
    }
}
