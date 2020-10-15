using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelHeading : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Heading;

        /// <summary>
        /// Builds an excel paragraph as specified by the Element model using EPPlus.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="documentSection"></param>
        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            HeadingData elementData = (HeadingData)element.GetData();
            IExcelElementStyle style = element.ExcelStyle;

            
            int row = style.Cell.NextRow(NextRow(documentSection));
            int col = style.Cell.NextColumn(NextColumn(documentSection, row));
            CellStyle(documentSection, row, col, style);
            documentSection.Cells[row, col].Value = elementData.Text;

        }


    }
}
