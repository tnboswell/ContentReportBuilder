using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ContentReportBuilder.OpenXml.Builder.Excel.Elements
{
    public class ExcelImage : DocumentElementBuilderBase<Sheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Image;

        protected override void Build(DocumentModelElement element, Sheet documentSection)
        {
            ImageDataBase elementData = (ImageDataBase)element.GetData();
            throw new NotImplementedException();
        }
    }
}
