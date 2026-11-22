using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLPrintAreas : IXLPrintAreas
    {
        List<IXLRange> ranges = new List<IXLRange>();
        private readonly XLWorksheet? worksheet;
        public XLPrintAreas(XLWorksheet? worksheet)
        {
            this.worksheet = worksheet;
        }

        public XLPrintAreas(XLPrintAreas defaultPrintAreas, XLWorksheet? worksheet)
        {
            ranges = defaultPrintAreas.ranges.ToList();
            this.worksheet = worksheet;
        }

        /// <summary>
        /// The worksheet of the print areas. The default page setup of a workbook has no worksheet.
        /// </summary>
        private XLWorksheet RequiredWorksheet => worksheet ?? throw new InvalidOperationException("The print areas have no worksheet.");

        public void Clear()
        {
            ranges.Clear();
        }

        public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            ranges.Add(RequiredWorksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
        }

        public void Add(string rangeAddress)
        {
            ranges.Add(RequiredWorksheet.Range(rangeAddress) ?? throw new ArgumentException($"The range '{rangeAddress}' is not valid.", nameof(rangeAddress)));
        }

        public void Add(string firstCellAddress, string lastCellAddress)
        {
            ranges.Add(RequiredWorksheet.Range(firstCellAddress, lastCellAddress));
        }

        public void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            ranges.Add(RequiredWorksheet.Range(firstCellAddress, lastCellAddress));
        }

        public IEnumerator<IXLRange> GetEnumerator()
        {
            return ranges.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
