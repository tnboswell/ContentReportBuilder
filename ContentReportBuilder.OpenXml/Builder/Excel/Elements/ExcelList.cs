using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ContentReportBuilder.OpenXml.Builder.Excel.Elements
{
    public class ExcelList : DocumentElementBuilderBase<Sheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.List;

        protected override void Build(DocumentModelElement element, Sheet documentSection)
        {
            ListData elementData = (ListData)element.GetData();
            throw new NotImplementedException();
        }
    }
}
