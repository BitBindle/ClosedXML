#nullable disable

using System;

namespace ClosedXML.Excel
{
    internal class XLCFDataBar : IXLCFDataBar
    {
        private readonly XLConditionalFormat _conditionalFormat;

        public XLCFDataBar(XLConditionalFormat conditionalFormat)
        {
            _conditionalFormat = conditionalFormat;
        }

        public Boolean Gradient => _conditionalFormat.Gradient;

        public IXLCFDataBar SetGradient(Boolean value = true)
        {
            _conditionalFormat.Gradient = value;
            return this;
        }
    }
}
