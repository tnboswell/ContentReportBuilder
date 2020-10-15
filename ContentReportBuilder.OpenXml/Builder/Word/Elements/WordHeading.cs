using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using DocumentFormat.OpenXml.Wordprocessing;
using System;


namespace ContentReportBuilder.OpenXml.Builder.Word.Elements
{
    public class WordHeading : DocumentElementBuilderBase<Body>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Heading;

        protected override void Build(DocumentModelElement element, Body documentSection)
        {
            HeadingData elementData = (HeadingData)element.GetData();
            throw new NotImplementedException();
        }
    }
}
