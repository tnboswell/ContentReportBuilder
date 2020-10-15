using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using System;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelImage : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Image;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            ImageDataBase elementData = (ImageDataBase)element.GetData();

            IExcelElementStyle style = element.ExcelStyle;

            int row = style.Cell.NextRow(NextRow(documentSection));
            int col = style.Cell.NextColumn(NextColumn(documentSection, row));

            using (System.Drawing.Image image = System.Drawing.Image.FromStream(elementData.GetData()))
            {
                
                var drawingNumber = documentSection.Drawings.Count + 1;
                var excelImage = documentSection.Drawings.AddPicture($"Drawing {drawingNumber}", image);
                //add the image to row, column
                excelImage.SetPosition(row - 1 , 0, col - 1, 0);
                excelImage.SetSize(elementData.Width, elementData.Height);
            }

        }
    }
}
