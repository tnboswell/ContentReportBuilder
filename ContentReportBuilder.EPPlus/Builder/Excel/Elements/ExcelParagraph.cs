using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using System;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelParagraph : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Paragraph;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            ParagraphData elementData = (ParagraphData)element.GetData();
            IExcelElementStyle style = element.ExcelStyle;

            int row = style.Cell.NextRow(NextRow(documentSection));
            int col = style.Cell.NextColumn(NextColumn(documentSection, row));
            CellStyle(documentSection, row, col, style);
            documentSection.Cells[row, col].Value = ToDataType(style.Cell.DataType, elementData.Text);
        }
    }
}
