using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using System;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelList : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.List;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            ListData elementData = (ListData)element.GetData();
            IExcelElementStyle style = element.ExcelStyle;

            int row = style.Cell.NextRow(NextRow(documentSection));
            int col = style.Cell.NextColumn(NextColumn(documentSection, row));

            int index = 1;
            foreach(var item in elementData.Items)
            {
                CellStyle(documentSection, row, col, style);
                var orderPrefix = (elementData.Ordered) ? $"{index}. " : "";
                documentSection.Cells[row, col].Value = $"{orderPrefix}{item.Text}";
                index++;
                row++;
            }

            
        }
    }
}
