using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordHeading : DocumentElementBuilder<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Heading;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            HeadingData elementData = (HeadingData)element.GetData();
            var heading = documentSection.InsertParagraph(elementData.Text);
            ParagraphFormating(element.WordStyle, heading);
        }

    }
}
