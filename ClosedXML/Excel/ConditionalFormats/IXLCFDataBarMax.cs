using System;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMax
    {
        IXLCFDataBar Maximum(XLCFContentType type, String value);

        IXLCFDataBar Maximum(XLCFContentType type, Double value);

        IXLCFDataBar HighestValue();
    }
}
