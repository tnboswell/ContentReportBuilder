using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ContentReportBuilder.OpenXml.Builder.Excel.Elements
{
    public class ExcelHeading : DocumentElementBuilderBase<Sheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Heading;

        protected override void Build(DocumentModelElement element, Sheet documentSection)
        {
            throw new NotImplementedException();
        }
    }
}
