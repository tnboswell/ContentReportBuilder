using DocumentFormat.OpenXml.Wordprocessing;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System;
using ContentReportBuilder.Builder;

namespace ContentReportBuilder.OpenXml.Builder.Word.Elements
{
    public class WordImage : DocumentElementBuilderBase<Body>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Image;

        protected override void Build(DocumentModelElement element, Body documentSection)
        {
            ImageDataBase elementData = (ImageDataBase)element.GetData();
            throw new NotImplementedException();
        }
    }
}
