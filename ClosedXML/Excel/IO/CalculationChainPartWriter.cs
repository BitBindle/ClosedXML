using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using static ClosedXML.Excel.XLWorkbook;

namespace ClosedXML.Excel.IO
{
    internal class CalculationChainPartWriter
    {
        internal static void GenerateContent(WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
        {
            var calculationChainPart = workbookPart.CalculationChainPart ??
                                       workbookPart.AddNewPart<CalculationChainPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
            var calculationChain = calculationChainPart.CalculationChain ??= new CalculationChain();
            calculationChain.RemoveAllChildren<CalculationCell>();

            foreach (var worksheet in workbook.WorksheetsInternal)
            {
                foreach (var c in worksheet.Internals.CellsCollection.GetCells())
                {
                    if (c.Formula is not { } formula)
                        continue;

                    if (formula.Type == FormulaType.DataTable)
                    {
                        // Do nothing, Excel doesn't generate calc chain for data table
                    }
                    else if (c.HasArrayFormula)
                    {
                        var formulaReference = c.FormulaReference ??= c.AsRange().RangeAddress;
                        if (formulaReference.FirstAddress.Equals(c.Address))
                        {
                            var cc = new CalculationCell
                            {
                                CellReference = c.Address.ToString(),
                                SheetId = (Int32)worksheet.SheetId
                            };

                            cc.Array = true;
                            calculationChain.AppendChild(cc);

                            foreach (var childCell in worksheet.Range(formulaReference).Cells())
                            {
                                calculationChain.AppendChild(new CalculationCell
                                {
                                    CellReference = childCell.Address.ToString(),
                                    SheetId = (Int32)worksheet.SheetId,
                                });
                            }
                        }
                    }
                    else
                    {
                        calculationChain.AppendChild(new CalculationCell
                        {
                            CellReference = c.Address.ToString(),
                            SheetId = (Int32)worksheet.SheetId
                        });
                    }
                }
            }

            if (!calculationChain.Any())
                workbookPart.DeletePart(calculationChainPart);
        }
    }
}
