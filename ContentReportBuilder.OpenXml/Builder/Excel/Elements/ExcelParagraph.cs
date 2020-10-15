using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ContentReportBuilder.OpenXml.Builder.Excel.Elements
{
    public class ExcelParagraph : DocumentElementBuilderBase<Sheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Paragraph;

        protected override void Build(DocumentModelElement element, Sheet documentSection)
        {
            ParagraphData elementData = (ParagraphData)element.GetData();
            throw new NotImplementedException();
        }
    }
}
