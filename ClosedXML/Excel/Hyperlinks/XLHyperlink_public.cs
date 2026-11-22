using System;
using System.Diagnostics.CodeAnalysis;

namespace ClosedXML.Excel
{
    public partial class XLHyperlink
    {
        private Uri? _externalAddress;
        private String? _internalAddress;
        private String _tooltip = String.Empty;

        public XLHyperlink(String address)
        {
            SetValues(address, String.Empty);
        }

        public XLHyperlink(String address, String tooltip)
        {
            SetValues(address, tooltip);
        }

        public XLHyperlink(IXLCell cell)
        {
            SetValues(cell, String.Empty);
        }

        public XLHyperlink(IXLCell cell, String tooltip)
        {
            SetValues(cell, tooltip);
        }

        public XLHyperlink(IXLRangeBase range)
        {
            SetValues(range, String.Empty);
        }

        public XLHyperlink(IXLRangeBase range, String tooltip)
        {
            SetValues(range, tooltip);
        }

        public XLHyperlink(Uri uri)
        {
            SetValues(uri, String.Empty);
        }

        public XLHyperlink(Uri uri, String tooltip)
        {
            SetValues(uri, tooltip);
        }

        public Boolean IsExternal { get; set; }

        /// <summary>
        /// The external address of the hyperlink, or <c>null</c> if the hyperlink is not external.
        /// </summary>
        /// <exception cref="ArgumentNullException">The value to set is <c>null</c>.</exception>
        [DisallowNull]
        public Uri? ExternalAddress
        {
            get
            {
                return IsExternal ? _externalAddress : null;
            }
            set
            {
                _externalAddress = value ?? throw new ArgumentNullException(nameof(value));
                IsExternal = true;
            }
        }

        /// <summary>
        /// Gets top left cell of a hyperlink range. Return <c>null</c>,
        /// if the hyperlink isn't in a worksheet.
        /// </summary>
        public IXLCell? Cell
        {
            get
            {
                if (Container is null)
                    return null;

                return Container.GetCell(this);
            }
        }

        /// <summary>
        /// The internal address of the hyperlink, or <c>null</c> if the hyperlink is external
        /// or has no address.
        /// </summary>
        /// <exception cref="ArgumentNullException">The value to set is <c>null</c>.</exception>
        [DisallowNull]
        public String? InternalAddress
        {
            get
            {
                if (IsExternal || _internalAddress is null)
                    return null;
                if (_internalAddress.Contains('!'))
                {
                    return _internalAddress[0] != '\''
                               ? String.Concat(
                                    _internalAddress
                                        .Substring(0, _internalAddress.IndexOf('!'))
                                        .EscapeSheetName(),
                                    '!',
                                    _internalAddress.Substring(_internalAddress.IndexOf('!') + 1))
                               : _internalAddress;
                }

                if (Container is null)
                    throw new InvalidOperationException("Hyperlink is not attached to a worksheet.");

                var sheetName = Container.WorksheetName;
                return String.Concat(
                    sheetName.EscapeSheetName(),
                    '!',
                    _internalAddress);
            }
            set
            {
                _internalAddress = value ?? throw new ArgumentNullException(nameof(value));
                IsExternal = false;
            }
        }

        /// <summary>
        /// Tooltip displayed when user hovers over the hyperlink range. If not specified,
        /// the link target is displayed in the tooltip. A <c>null</c> value sets an empty tooltip.
        /// </summary>
        [AllowNull]
        public String Tooltip
        {
            get => _tooltip;
            set => _tooltip = value ?? String.Empty;
        }

        /// <inheritdoc cref="IXLHyperlinks.Delete(XLHyperlink)"/>
        public void Delete()
        {
            Container?.Delete(this);
        }
    }
}
